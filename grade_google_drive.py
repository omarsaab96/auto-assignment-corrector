from __future__ import annotations

import argparse
import gc
import re
import tempfile
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from openpyxl import load_workbook

from grade_assignments import DEFAULT_TEMPLATE_NAME, Difference, grade_workbook, template_value_cache, write_summary


SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FOLDER_MIME_TYPE = "application/vnd.google-apps.folder"
DEFAULT_CORRECTED_FOLDER_PREFIX = "corrected"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Grade XLSX submissions stored in Google Drive and upload corrected copies.",
    )
    parser.add_argument(
        "--students-folder-id",
        required=True,
        help="Google Drive folder ID containing student submissions.",
    )
    parser.add_argument(
        "--template-file-id",
        default=None,
        help="Google Drive file ID of the template solution workbook.",
    )
    parser.add_argument(
        "--template-name",
        default=DEFAULT_TEMPLATE_NAME,
        help="Template file name to search for when --template-file-id is not supplied.",
    )
    parser.add_argument(
        "--corrected-folder-name",
        default=None,
        help="Name of the output folder created inside the students folder. Default: corrected-YYYYMMDD-HHMMSS",
    )
    parser.add_argument(
        "--credentials",
        type=Path,
        default=Path("credentials.json"),
        help="OAuth client credentials JSON from Google Cloud. Default: ./credentials.json",
    )
    parser.add_argument(
        "--token",
        type=Path,
        default=Path("token.json"),
        help="Saved OAuth token path. Default: ./token.json",
    )
    parser.add_argument(
        "--ignore-case",
        action="store_true",
        help="Ignore letter case when comparing text cells.",
    )
    parser.add_argument(
        "--no-trim-text",
        dest="trim_text",
        action="store_false",
        help="Treat leading and trailing spaces as differences.",
    )
    parser.set_defaults(trim_text=True)
    return parser.parse_args()


def timestamped_corrected_folder_name(now: datetime | None = None) -> str:
    current_time = now or datetime.now()
    return f"{DEFAULT_CORRECTED_FOLDER_PREFIX}-{current_time:%Y%m%d-%H%M%S}"


def resolve_corrected_folder_name(folder_name: str | None) -> str:
    name = (folder_name or "").strip()
    if not name or name.casefold() == DEFAULT_CORRECTED_FOLDER_PREFIX:
        return timestamped_corrected_folder_name()
    if re.search(r"\d{8}-\d{6}$", name):
        return name
    return f"{name}-{timestamped_corrected_folder_name().removeprefix(DEFAULT_CORRECTED_FOLDER_PREFIX + '-')}"


def authenticate(credentials_path: Path, token_path: Path):
    credentials = None
    if token_path.exists():
        credentials = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if credentials and credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
        except RefreshError:
            credentials = None
            token_path.unlink(missing_ok=True)

    if not credentials or not credentials.valid:
        if not credentials_path.is_file():
            raise FileNotFoundError(f"Google OAuth credentials file not found: {credentials_path}")
        flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
        credentials = flow.run_local_server(port=0)
        token_path.write_text(credentials.to_json(), encoding="utf-8")

    return build("drive", "v3", credentials=credentials)


def drive_list(service, query: str, fields: str):
    results = []
    request = service.files().list(
        q=query,
        spaces="drive",
        fields=f"nextPageToken, files({fields})",
        pageSize=1000,
    )
    while request is not None:
        response = request.execute()
        results.extend(response.get("files", []))
        request = service.files().list_next(request, response)
    return results


def get_file(service, file_id: str, fields: str = "id, name, mimeType"):
    return service.files().get(fileId=file_id, fields=fields).execute()


def find_template_by_name(service, template_name: str):
    escaped_name = template_name.replace("'", "\\'")
    query = (
        f"name = '{escaped_name}' and "
        f"mimeType = '{XLSX_MIME_TYPE}' and "
        "trashed = false"
    )
    matches = drive_list(service, query, "id, name, modifiedTime")
    if not matches:
        raise FileNotFoundError(f"Template file not found in Google Drive: {template_name}")
    if len(matches) > 1:
        ids = ", ".join(f"{item['name']} ({item['id']})" for item in matches)
        raise RuntimeError(f"Multiple template files matched. Use --template-file-id. Matches: {ids}")
    return matches[0]


def escape_drive_query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_or_create_child_folder(service, parent_folder_id: str, folder_name: str) -> str:
    escaped_name = escape_drive_query_value(folder_name)
    query = (
        f"'{parent_folder_id}' in parents and "
        f"name = '{escaped_name}' and "
        f"mimeType = '{FOLDER_MIME_TYPE}' and "
        "trashed = false"
    )
    matches = drive_list(service, query, "id, name")
    if matches:
        return matches[0]["id"]

    metadata = {
        "name": folder_name,
        "mimeType": FOLDER_MIME_TYPE,
        "parents": [parent_folder_id],
    }
    folder = service.files().create(body=metadata, fields="id").execute()
    return folder["id"]


def list_student_workbooks(
    service,
    students_folder_id: str,
    template_file_id: str | None,
    excluded_folder_names: set[str] | None = None,
):
    excluded_names = {name.casefold() for name in (excluded_folder_names or {DEFAULT_CORRECTED_FOLDER_PREFIX})}
    workbooks = []

    def collect(folder_id: str, folder_path: tuple[str, ...]) -> None:
        query = (
            f"'{folder_id}' in parents and "
            f"(mimeType = '{XLSX_MIME_TYPE}' or mimeType = '{FOLDER_MIME_TYPE}') and "
            "trashed = false"
        )
        children = drive_list(service, query, "id, name, mimeType, modifiedTime")
        for child in sorted(children, key=lambda item: (item["mimeType"] != FOLDER_MIME_TYPE, item["name"].casefold())):
            if child["mimeType"] == FOLDER_MIME_TYPE:
                if child["name"].casefold() in excluded_names:
                    continue
                collect(child["id"], (*folder_path, child["name"]))
                continue

            if child["name"].startswith("~$") or child["id"] == template_file_id:
                continue

            relative_parts = (*folder_path, child["name"])
            workbooks.append(
                {
                    **child,
                    "folder_path": list(folder_path),
                    "relative_path": "/".join(relative_parts),
                    "student_name": folder_path[0] if folder_path else "",
                }
            )

    collect(students_folder_id, ())
    return sorted(workbooks, key=lambda item: item["relative_path"].casefold())


def download_file(service, file_id: str, output_path: Path) -> None:
    request = service.files().get_media(fileId=file_id)
    with output_path.open("wb") as output_file:
        downloader = MediaIoBaseDownload(output_file, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def find_or_create_corrected_folder(service, students_folder_id: str, folder_name: str) -> str:
    return find_or_create_child_folder(service, students_folder_id, folder_name)


def corrected_destination_folder(service, corrected_root_id: str, student_file: dict) -> str:
    folder_id = corrected_root_id
    for folder_name in student_file.get("folder_path", []):
        folder_id = find_or_create_child_folder(service, folder_id, folder_name)
    return folder_id


def upload_or_replace_file(service, folder_id: str, path: Path, mime_type: str) -> str:
    escaped_name = escape_drive_query_value(path.name)
    query = (
        f"'{folder_id}' in parents and "
        f"name = '{escaped_name}' and "
        "trashed = false"
    )
    matches = drive_list(service, query, "id, name")
    media = MediaFileUpload(str(path), mimetype=mime_type, resumable=True)

    if matches:
        file_id = matches[0]["id"]
        service.files().update(fileId=file_id, media_body=media, fields="id").execute()
        return file_id

    metadata = {"name": path.name, "parents": [folder_id]}
    uploaded = service.files().create(body=metadata, media_body=media, fields="id").execute()
    return uploaded["id"]


def upload_summary_text(service, folder_id: str, path: Path) -> str:
    return upload_or_replace_file(service, folder_id, path, "text/csv")


def make_grade_args(args: argparse.Namespace):
    return SimpleNamespace(ignore_case=args.ignore_case, trim_text=args.trim_text)


def main() -> int:
    args = parse_args()
    corrected_folder_name = resolve_corrected_folder_name(args.corrected_folder_name)
    service = authenticate(args.credentials.resolve(), args.token.resolve())
    corrected_folder_id = find_or_create_corrected_folder(
        service,
        args.students_folder_id,
        corrected_folder_name,
    )

    template_file = (
        get_file(service, args.template_file_id)
        if args.template_file_id
        else find_template_by_name(service, args.template_name)
    )
    student_files = list_student_workbooks(
        service,
        args.students_folder_id,
        template_file["id"],
        excluded_folder_names={corrected_folder_name, DEFAULT_CORRECTED_FOLDER_PREFIX},
    )

    with tempfile.TemporaryDirectory(prefix="assignment-corrector-") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        students_dir = temp_dir / "students"
        corrected_dir = temp_dir / "corrected"
        students_dir.mkdir()
        corrected_dir.mkdir()

        template_path = temp_dir / template_file["name"]
        download_file(service, template_file["id"], template_path)
        template_wb = load_workbook(template_path, read_only=True, keep_links=False)
        try:
            template_value_cache(template_wb)
            all_differences: list[Difference] = []
            grade_args = make_grade_args(args)

            for student_file in student_files:
                student_source_dir = students_dir / student_file["id"]
                student_source_dir.mkdir()
                student_path = student_source_dir / student_file["name"]
                download_file(service, student_file["id"], student_path)
                differences = grade_workbook(template_wb, student_path, corrected_dir, grade_args)
                labeled_differences = [
                    replace(diff, workbook=student_file["relative_path"])
                    for diff in differences
                ]
                all_differences.extend(labeled_differences)
                corrected_path = corrected_dir / student_path.name
                destination_folder_id = corrected_destination_folder(service, corrected_folder_id, student_file)
                upload_or_replace_file(service, destination_folder_id, corrected_path, XLSX_MIME_TYPE)
                print(f"{student_file['relative_path']}: {len(differences)} difference(s)")
                gc.collect()
        finally:
            template_wb.close()
            gc.collect()

        summary_path = write_summary(corrected_dir, all_differences)
        upload_summary_text(service, corrected_folder_id, summary_path)

    print(f"Graded {len(student_files)} workbook(s).")
    print(f"Corrected Google Drive folder ID: {corrected_folder_id}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

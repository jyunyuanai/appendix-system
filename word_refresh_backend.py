from __future__ import annotations

import base64
import json
import tempfile
from contextlib import nullcontext
from pathlib import Path
from urllib import error as urllib_error
from urllib import request as urllib_request


def word_bookmark_adjusted_page_number(document, bookmark_name: str) -> int | None:
    try:
        bookmark_range = document.Bookmarks(bookmark_name).Range
        page_number = bookmark_range.Information(3)
    except Exception:
        return None

    if not page_number:
        return None
    return int(page_number) - 1 if int(page_number) > 1 else 1


def word_find_generated_toc_table(document):
    try:
        for table in document.Tables:
            if table.Rows.Count < 2 or table.Columns.Count < 3:
                continue
            header_text = table.Cell(1, 2).Range.Text.replace("\r", "").replace("\x07", "").strip()
            if "抽查" in header_text and "名稱" in header_text:
                return table
    except Exception:
        return None
    return None


def word_set_toc_cell_text(cell, text: str) -> None:
    range_object = cell.Range
    range_object.End -= 1
    range_object.Text = text


def word_apply_toc_page_ranges(
    document,
    appendix_number: int,
    toc_page_range_bookmarks: list[tuple[str, str]],
) -> None:
    toc_table = word_find_generated_toc_table(document)
    if toc_table is None:
        return

    try:
        row_count = toc_table.Rows.Count
    except Exception:
        return

    for pair_index, (start_bookmark_name, end_bookmark_name) in enumerate(
        toc_page_range_bookmarks,
        start=2,
    ):
        if pair_index > row_count:
            break

        start_page = word_bookmark_adjusted_page_number(document, start_bookmark_name)
        end_page = word_bookmark_adjusted_page_number(document, end_bookmark_name)
        if start_page is None and end_page is None:
            continue
        if start_page is None:
            start_page = end_page
        if end_page is None:
            end_page = start_page

        if start_page == end_page:
            page_text = f"附錄 {appendix_number}-{start_page}"
        else:
            page_text = f"附錄 {appendix_number}-{start_page}至附錄 {appendix_number}-{end_page}"

        try:
            word_set_toc_cell_text(toc_table.Cell(pair_index, 3), page_text)
        except Exception:
            continue


def word_update_header_footer_fields(document) -> None:
    try:
        for section in document.Sections:
            for footer_index in (1, 2, 3):
                section.Footers(footer_index).Range.Fields.Update()
                section.Headers(footer_index).Range.Fields.Update()
    except Exception:
        pass


def refresh_docx_fields_with_local_word(
    docx_bytes: bytes,
    appendix_number: int | None = None,
    toc_page_range_bookmarks: list[tuple[str, str]] | None = None,
    word_com_lock=None,
) -> bytes:
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return docx_bytes

    lock_context = word_com_lock if word_com_lock is not None else nullcontext()

    with lock_context:
        pythoncom.CoInitialize()
        word = None
        document = None
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                input_path = Path(temp_dir) / "input.docx"
                output_path = Path(temp_dir) / "output.docx"
                input_path.write_bytes(docx_bytes)

                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                try:
                    word.ScreenUpdating = False
                except Exception:
                    pass
                try:
                    word.AutomationSecurity = 3
                except Exception:
                    pass

                document = word.Documents.Open(
                    str(input_path),
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                    OpenAndRepair=False,
                )
                try:
                    document.Repaginate()
                except Exception:
                    pass

                if appendix_number is not None and toc_page_range_bookmarks:
                    word_apply_toc_page_ranges(
                        document,
                        appendix_number,
                        toc_page_range_bookmarks,
                    )
                    try:
                        document.Repaginate()
                    except Exception:
                        pass

                word_update_header_footer_fields(document)

                document.SaveAs2(str(output_path), FileFormat=16)
                document.Close(False)
                document = None
                return output_path.read_bytes()
        except Exception:
            return docx_bytes
        finally:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def build_refresh_service_url(service_url: str) -> str:
    base_url = service_url.rstrip("/")
    if base_url.endswith("/refresh"):
        return base_url
    return f"{base_url}/refresh"


def refresh_docx_fields_via_service(
    docx_bytes: bytes,
    service_url: str,
    appendix_number: int | None = None,
    toc_page_range_bookmarks: list[tuple[str, str]] | None = None,
    service_token: str = "",
    timeout_seconds: int = 300,
) -> bytes | None:
    endpoint = build_refresh_service_url(service_url)
    payload = {
        "docx_base64": base64.b64encode(docx_bytes).decode("ascii"),
        "appendix_number": appendix_number,
        "toc_page_range_bookmarks": list(toc_page_range_bookmarks or []),
    }
    request_data = json.dumps(payload).encode("utf-8")
    request_headers = {"Content-Type": "application/json"}
    if service_token:
        request_headers["Authorization"] = f"Bearer {service_token}"

    request = urllib_request.Request(
        endpoint,
        data=request_data,
        headers=request_headers,
        method="POST",
    )

    try:
        with urllib_request.urlopen(request, timeout=timeout_seconds) as response:
            if response.status != 200:
                return None
            response_payload = json.loads(response.read().decode("utf-8"))
    except (urllib_error.URLError, TimeoutError, ValueError, OSError):
        return None

    docx_base64 = response_payload.get("docx_base64")
    if not docx_base64:
        return None

    try:
        return base64.b64decode(docx_base64)
    except Exception:
        return None

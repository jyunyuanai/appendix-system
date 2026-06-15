from __future__ import annotations

import base64
import copy
import hashlib
import hmac
import io
import os
import re
import shutil
import subprocess
import tempfile
import threading
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import streamlit as st
from docx import Document
from lxml import etree
from word_refresh_backend import refresh_docx_fields_via_service
from word_refresh_backend import refresh_docx_fields_with_local_word


APP_DIR = Path(__file__).resolve().parent
APPENDIX_NAMES = ["附錄一", "附錄二", "附錄三", "附錄四", "附錄五", "附錄六"]
DEFAULT_WORK_ITEMS = ["放樣", "開挖", "回填", "便道"]
UI_MASCOT_IMAGE = APP_DIR / "assets" / "ui_mascots.png"
PREPARE_CACHE_VERSION = 50
MAX_PREPARE_WORKERS = 4
MAX_OUTPUT_WORKERS = 6
APPENDIX_CODE_PREFIXES = ["A", "B", "C", "D", "E", "F"]
APPENDIX_CODE_LABEL_TEXTS = {
    "編號",
    "表單編號",
    "抽查編號",
    "查驗編號",
    "檢驗編號",
    "紀錄編號",
    "記錄編號",
    "文件編號",
}
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WORD_NAMESPACES = {"w": WORD_NAMESPACE, "r": RELATIONSHIP_NAMESPACE}
W_P = f"{{{WORD_NAMESPACE}}}p"
W_TBL = f"{{{WORD_NAMESPACE}}}tbl"
W_TR = f"{{{WORD_NAMESPACE}}}tr"
W_TC = f"{{{WORD_NAMESPACE}}}tc"
W_T = f"{{{WORD_NAMESPACE}}}t"
W_SECTPR = f"{{{WORD_NAMESPACE}}}sectPr"
WORD_COM_LOCK = threading.Lock()


def get_secret_setting(name: str, default: str = "") -> str:
    env_value = os.environ.get(name)
    if env_value is not None and str(env_value).strip():
        return str(env_value).strip()

    try:
        secret_value = st.secrets.get(name, default)
    except Exception:
        secret_value = default

    if secret_value is None:
        return default
    return str(secret_value).strip()


def require_app_password() -> None:
    app_password = get_secret_setting("APP_PASSWORD")
    if not app_password:
        return
    if st.session_state.get("app_password_ok"):
        return

    st.title("監造附錄六合一產出系統")
    st.caption("請先輸入密碼")

    with st.form("app_password_form"):
        password_input = st.text_input("密碼", type="password")
        submitted = st.form_submit_button("登入", use_container_width=True)

    if submitted:
        if hmac.compare_digest(password_input, app_password):
            st.session_state["app_password_ok"] = True
            st.rerun()
        else:
            st.error("密碼錯誤。")

    st.stop()
TOC_WORK_ITEM_SUFFIXES = [
    "工程施工安全衛生抽查管理標準表",
    "工程施工安全衛生查驗管理標準表",
    "工程施工安全衛生檢驗管理標準表",
    "工程施工安全衛生抽查標準表",
    "工程施工安全衛生查驗標準表",
    "工程施工安全衛生檢驗標準表",
    "施工安全衛生抽查管理標準表",
    "施工安全衛生查驗管理標準表",
    "施工安全衛生檢驗管理標準表",
    "施工安全衛生抽查標準表",
    "施工安全衛生查驗標準表",
    "施工安全衛生檢驗標準表",
    "工程施工安全衛生抽查程序流程圖",
    "工程施工安全衛生查驗程序流程圖",
    "工程施工安全衛生檢驗程序流程圖",
    "工程施工安全衛生抽查紀錄表",
    "工程施工安全衛生查驗紀錄表",
    "工程施工安全衛生檢驗紀錄表",
    "施工安全衛生抽查程序流程圖",
    "施工安全衛生查驗程序流程圖",
    "施工安全衛生檢驗程序流程圖",
    "施工安全衛生抽查紀錄表",
    "施工安全衛生查驗紀錄表",
    "施工安全衛生檢驗紀錄表",
    "安全衛生抽查程序流程圖",
    "安全衛生查驗程序流程圖",
    "安全衛生檢驗程序流程圖",
    "安全衛生抽查紀錄表",
    "安全衛生查驗紀錄表",
    "安全衛生檢驗紀錄表",
    "安全衛生抽查管理標準表",
    "安全衛生查驗管理標準表",
    "安全衛生檢驗管理標準表",
    "工程施工抽查程序流程圖",
    "工程施工抽查程序",
    "工程施工查驗程序",
    "工程施工檢驗程序",
    "工程施工抽查紀錄表",
    "工程施工查驗紀錄表",
    "工程施工檢驗紀錄表",
    "工程施工安全衛生抽查程序",
    "工程施工安全衛生查驗程序",
    "工程施工安全衛生檢驗程序",
    "工程施工抽查管理標準表",
    "工程施工查驗管理標準表",
    "工程施工檢驗管理標準表",
    "工程施工安全衛生抽查管理標準表",
    "工程施工安全衛生查驗管理標準表",
    "工程施工安全衛生檢驗管理標準表",
    "工程施工安全衛生抽查標準表",
    "工程施工安全衛生查驗標準表",
    "工程施工安全衛生檢驗標準表",
    "工程施工抽查表",
    "工程施工查驗表",
    "工程施工檢驗表",
    "工程抽查紀錄表",
    "工程查驗紀錄表",
    "工程檢驗紀錄表",
    "工程抽查管理標準表",
    "工程查驗管理標準表",
    "工程檢驗管理標準表",
    "工程安全衛生抽查標準表",
    "工程查驗管理標準表",
    "工程檢驗管理標準表",
    "施工抽查程序流程圖",
    "施工抽查程序",
    "施工查驗程序",
    "施工檢驗程序",
    "施工抽查紀錄表",
    "施工查驗紀錄表",
    "施工檢驗紀錄表",
    "施工安全衛生抽查程序",
    "施工安全衛生查驗程序",
    "施工安全衛生檢驗程序",
    "施工抽查管理標準表",
    "施工查驗管理標準表",
    "施工檢驗管理標準表",
    "施工安全衛生抽查標準表",
    "施工安全衛生查驗標準表",
    "施工安全衛生檢驗標準表",
    "施工抽查表",
    "施工查驗表",
    "施工檢驗表",
    "施工品質抽查紀錄表",
    "施工品質查驗紀錄表",
    "施工品質檢驗紀錄表",
    "施工品質抽查表",
    "安全衛生抽查標準表",
    "安全衛生抽查程序",
    "安全衛生查驗程序",
    "安全衛生檢驗程序",
    "安全衛生查驗標準表",
    "安全衛生檢驗標準表",
    "抽查程序流程圖",
    "抽查程序",
    "抽查紀錄表",
    "查驗紀錄表",
    "檢驗紀錄表",
    "抽查管理標準表",
    "抽查標準表",
    "抽查表",
    "查驗程序",
    "查驗管理標準表",
    "查驗標準表",
    "查驗表",
    "檢驗程序",
    "檢驗管理標準表",
    "檢驗標準表",
    "檢驗表",
    "管理標準表",
    "標準表",
    "流程圖",
]


def bytes_digest(file_bytes: bytes) -> str:
    return hashlib.blake2b(file_bytes, digest_size=16).hexdigest()


def worker_count(item_count: int, limit: int) -> int:
    return max(1, min(item_count, limit))


def appendix_position_from_filename(file_name: str) -> int | None:
    text = Path(file_name).stem
    number_texts = {
        "1": 0,
        "01": 0,
        "１": 0,
        "一": 0,
        "壹": 0,
        "2": 1,
        "02": 1,
        "２": 1,
        "二": 1,
        "貳": 1,
        "3": 2,
        "03": 2,
        "３": 2,
        "三": 2,
        "參": 2,
        "4": 3,
        "04": 3,
        "４": 3,
        "四": 3,
        "肆": 3,
        "5": 4,
        "05": 4,
        "５": 4,
        "五": 4,
        "伍": 4,
        "6": 5,
        "06": 5,
        "６": 5,
        "六": 5,
        "陸": 5,
    }
    match = re.search(r"附錄\s*[-_（(]*\s*([0０]?[1-6１-６一二三四五六壹貳參肆伍陸])", text)
    if not match:
        return None
    return number_texts.get(match.group(1))


def toc_source_position_from_payloads(payloads: list[dict]) -> int:
    for payload in payloads:
        if appendix_position_from_filename(payload["name"]) == 0:
            return int(payload["position"])
    return 0


def uploaded_file_payloads(uploaded_files) -> list[dict]:
    payloads_by_position = {}
    fallback_payloads = []
    for index, uploaded_file in enumerate(uploaded_files[: len(APPENDIX_NAMES)]):
        file_bytes = uploaded_file.getvalue()
        payload = {
            "position": index,
            "name": uploaded_file.name,
            "bytes": file_bytes,
            "size": len(file_bytes),
            "type": getattr(uploaded_file, "type", ""),
            "digest": bytes_digest(file_bytes),
        }
        file_position = appendix_position_from_filename(uploaded_file.name)
        if file_position is not None and file_position not in payloads_by_position:
            payload["position"] = file_position
            payloads_by_position[file_position] = payload
        else:
            fallback_payloads.append(payload)

    available_positions = [
        position
        for position in range(len(APPENDIX_NAMES))
        if position not in payloads_by_position
    ]
    for payload, position in zip(fallback_payloads, available_positions):
        payload["position"] = position
        payloads_by_position[position] = payload

    return [
        payloads_by_position[position]
        for position in sorted(payloads_by_position)
    ]


def uploaded_payloads_signature(payloads: list[dict]) -> tuple[tuple[str, int, str, str], ...]:
    return tuple(
        (payload["name"], payload["size"], payload["type"], payload["digest"])
        for payload in payloads
    )


def sanitize_filename_part(name: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", name).strip(" ._")
    return sanitized or "完成版"


def image_data_url(path: Path) -> str | None:
    if not path.exists():
        return None

    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def render_page_header() -> None:
    mascot_data_url = image_data_url(UI_MASCOT_IMAGE)
    wallpaper_background = (
        f'linear-gradient(rgba(255, 249, 245, 0.68), rgba(255, 255, 255, 0.74)), url("{mascot_data_url}")'
        if mascot_data_url
        else "linear-gradient(180deg, #fff9f5 0%, #ffffff 38%)"
    )
    wallpaper_size = "100% 100%, 360px auto" if mascot_data_url else "auto"
    wallpaper_repeat = "no-repeat, repeat" if mascot_data_url else "no-repeat"
    wallpaper_position = "center, left top" if mascot_data_url else "center"

    st.markdown(
        f"""
        <style>
            .block-container {{
                max-width: 1500px;
                padding-top: 2rem;
            }}
            html,
            body,
            .stApp {{
                background: #fff9f5;
            }}
            .stApp {{
                isolation: isolate;
            }}
            .stApp::before {{
                content: "";
                position: fixed;
                inset: 0;
                z-index: 0;
                pointer-events: none;
                background-image: {wallpaper_background};
                background-size: {wallpaper_size};
                background-repeat: {wallpaper_repeat};
                background-position: {wallpaper_position};
            }}
            .stApp > div {{
                position: relative;
                z-index: 1;
            }}
            [data-testid="stAppViewContainer"] {{
                background: transparent;
            }}
            header[data-testid="stHeader"] {{
                background: transparent;
                height: 0;
            }}
            header[data-testid="stHeader"] > div,
            div[data-testid="stToolbar"],
            div[data-testid="stDecoration"],
            #MainMenu,
            footer {{
                display: none;
            }}
            .mascot-hero {{
                margin: 0 0 1.4rem 0;
                padding: 0.8rem 0 0.4rem;
            }}
            .mascot-hero__title {{
                margin: 0;
                color: #23263a;
                font-size: 2.3rem;
                font-weight: 800;
                letter-spacing: 0;
            }}
            .mascot-hero__subtitle {{
                margin: 0.65rem 0 0;
                color: #6f7480;
                font-size: 1rem;
            }}
            div[data-testid="stFileUploaderDropzone"] {{
                background: #fff;
                border: 1px solid #efd7ca;
                border-radius: 14px;
            }}
            div[data-testid="stVerticalBlockBorderWrapper"] {{
                border-color: #efd7ca;
                border-radius: 14px;
                background: rgba(255, 255, 255, 0.8);
            }}
        </style>
        <div class="mascot-hero">
            <div>
                <h1 class="mascot-hero__title">監造附錄六合一產出系統</h1>
                <p class="mascot-hero__subtitle">匯入六份 Word，選取工項，快速產出六份完成版附錄。</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def clean_toc_text(text: str) -> str:
    text = text.replace("\u3000", " ")
    text = re.sub(r"\b(PAGEREF|HYPERLINK)\s+_[A-Za-z0-9_]+", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bTOC\s+\\[^\r\n\t ]+", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b_(Toc|Ref)[A-Za-z0-9_]*\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\\[A-Za-z]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s*[.．。·‧…⋯]{2,}\s*\d+\s*$", "", text)
    text = re.sub(r"\s+\d+\s*$", "", text)
    return text.strip(" -－—:：、")


def to_work_item_name(toc_text: str) -> str:
    cleaned_text = clean_toc_text(toc_text).replace("舖面", "鋪面")
    work_item = cleaned_text
    for suffix in TOC_WORK_ITEM_SUFFIXES:
        if work_item.endswith(suffix):
            work_item = work_item[: -len(suffix)]
            break
    for suffix in ("工程施工安全衛生", "施工安全衛生"):
        if work_item.endswith(suffix):
            work_item = work_item[: -len(suffix)]
            break
    if work_item.endswith("鋪面") and cleaned_text.startswith(f"{work_item}工程"):
        work_item = f"{work_item}工程"
    if work_item.endswith("工程"):
        keep_engineering_suffix = work_item.endswith("鋪面工程")
        if not keep_engineering_suffix:
            work_item = work_item[:-2]
    return work_item.strip(" -－—:：、")


def is_toc_candidate(text: str, style_name: str) -> bool:
    if not text or "目錄" in text or "分節符號" in text:
        return False
    if re.search(r"\bTOC\s*\\", text, flags=re.IGNORECASE):
        return True
    if style_name.startswith("TOC") or style_name.startswith("目錄"):
        return True
    has_page_number = bool(re.search(r"(\s|[.．。·‧…⋯])\d+\s*$", text))
    has_work_item_suffix = any(suffix in text for suffix in TOC_WORK_ITEM_SUFFIXES)
    has_leader = bool(re.search(r"([.．。·‧…⋯]{2,}|\t)", text))
    looks_like_toc_entry = has_page_number and has_leader and ("工程" in text or "表" in text)
    return has_page_number and (has_work_item_suffix or looks_like_toc_entry)


def is_toc_page_separator(text: str) -> bool:
    cleaned_text = clean_toc_text(text)
    if not cleaned_text:
        return True
    return "分節符號" in cleaned_text or "換頁" in cleaned_text


def extract_work_items_from_docx_bytes(docx_bytes: bytes) -> list[str]:
    work_items = []

    try:
        root = read_document_root(docx_bytes)
        toc_source = xml_extract_toc_source(root)
        if toc_source:
            if toc_source["type"] == "table":
                table = etree.fromstring(toc_source["table_xml"])
                for row in table.xpath("./w:tr", namespaces=WORD_NAMESPACES):
                    text = xml_row_text(row)
                    if is_toc_header_row(text):
                        continue
                    work_item = to_work_item_name(strip_toc_row_noise(text))
                    if work_item:
                        work_items.append(work_item)
            else:
                for entry in toc_source["entries"]:
                    work_item = to_work_item_name(strip_toc_row_noise(entry["text"]))
                    if work_item:
                        work_items.append(work_item)

            if work_items:
                # 回傳前才去重，避免表格型目錄被第一筆提前 return 而漏抓後面工項。
                return unique_work_items(work_items)
    except Exception:
        work_items = []

    document = Document(io.BytesIO(docx_bytes))
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        style_name = paragraph.style.name if paragraph.style is not None else ""
        if not is_toc_candidate(text, style_name):
            continue

        work_item = to_work_item_name(strip_toc_row_noise(text))
        if work_item:
            work_items.append(work_item)

    return unique_work_items(work_items)


@st.cache_data(show_spinner=False)


def normalize_doc_to_docx_bytes(file_name: str, file_bytes: bytes) -> bytes:
    suffix = Path(file_name).suffix.lower()
    if suffix == ".docx" or zipfile.is_zipfile(io.BytesIO(file_bytes)):
        return file_bytes
    if suffix != ".doc":
        raise ValueError("僅支援 .doc 與 .docx 檔案")

    conversion_errors = []
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        source_path = temp_path / "source.doc"
        source_path.write_bytes(file_bytes)

        try:
            converted_by_word = try_convert_doc_with_word(source_path, temp_path)
        except Exception as error:
            converted_by_word = None
            conversion_errors.append(f"Microsoft Word 轉檔失敗：{error}")
        if converted_by_word and converted_by_word.exists():
            return converted_by_word.read_bytes()

        try:
            converted_by_soffice = try_convert_doc_with_soffice(source_path, temp_path)
        except Exception as error:
            converted_by_soffice = None
            conversion_errors.append(f"LibreOffice 轉檔失敗：{error}")
        if converted_by_soffice and converted_by_soffice.exists():
            return converted_by_soffice.read_bytes()

    detail = "；".join(conversion_errors) if conversion_errors else "未找到可用轉檔工具"
    raise RuntimeError(f"無法轉換 .doc。{detail}")


def try_convert_doc_with_word(source_path: Path, output_dir: Path) -> Path | None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        raise RuntimeError(f"pywin32 未安裝或無法載入：{error}") from error

    output_path = output_dir / "converted.docx"
    word = None
    document = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        word.Options.SaveNormalPrompt = False
        word.Options.ConfirmConversions = False

        document = word.Documents.Open(
            str(source_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        document.SaveAs2(str(output_path), FileFormat=16)
        return output_path if output_path.exists() else None
    except Exception as error:
        raise RuntimeError(error) from error
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
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def try_convert_doc_with_soffice(source_path: Path, output_dir: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None

    subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(output_dir),
            str(source_path),
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    output_path = output_dir / f"{source_path.stem}.docx"
    return output_path if output_path.exists() else None


def prepare_word_payload(payload: dict, read_toc: bool = True) -> dict:
    """
    初始上傳時只需要從附錄一讀工項選單。
    附錄二～附錄六先不讀目錄，避免多花時間。
    後面正式產出時，仍然會針對每一份 Word 建立索引，所以不影響準確度。
    """
    try:
        docx_bytes = normalize_doc_to_docx_bytes(payload["name"], payload["bytes"])
        file_work_items = extract_work_items_from_docx_bytes(docx_bytes) if read_toc else []

        return {
            "position": payload["position"],
            "name": payload["name"],
            "digest": payload["digest"],
            "docx_digest": bytes_digest(docx_bytes),
            "docx_bytes": docx_bytes,
            "work_items": file_work_items,
            "index": None,
            "error": None,
        }
    except Exception as error:
        return {
            "position": payload["position"],
            "name": payload["name"],
            "digest": payload["digest"],
            "docx_digest": None,
            "docx_bytes": None,
            "work_items": [],
            "index": None,
            "error": str(error),
        }


def build_prepared_index(prepared_upload: dict, boundary_work_items: tuple[str, ...]) -> dict:
    if prepared_upload["docx_bytes"] is None:
        return prepared_upload

    indexed_upload = dict(prepared_upload)
    try:
        indexed_upload["index"] = build_appendix_index_from_docx_bytes(
            prepared_upload["docx_bytes"],
            list(boundary_work_items),
        )
    except Exception as error:
        indexed_upload["error"] = f"建立快速索引失敗：{error}"
    return indexed_upload


def appendix_boundary_work_items(
    prepared_upload: dict,
    output_work_items: tuple[str, ...],
) -> tuple[str, ...]:
    return tuple(
        unique_work_items(
            list(prepared_upload.get("work_items") or [])
            + list(output_work_items)
        )
    )


def build_single_output(
    index: int,
    prepared_upload: dict,
    work_items: tuple[str, ...],
    project_name: str = "",
) -> tuple[int, bytes, int, list[str]]:
    output_bytes, selected_element_count, missing_work_items = build_appendix_docx_fast(
        prepared_upload,
        list(work_items),
        project_name,
    )
    return index, output_bytes, selected_element_count, missing_work_items


def normalize_match_text(text: str) -> str:
    return re.sub(r"[\s　,，、;；:：.．。·‧…⋯\-－—_()（）]", "", text)


def strip_toc_row_noise(text: str) -> str:
    cleaned = clean_toc_text(text)
    cleaned = re.sub(r"^\s*[A-Za-z]+\s*\d+\s*", "", cleaned)
    cleaned = re.sub(r"^\s*\d+\s*", "", cleaned)
    cleaned = re.sub(r"\s*附錄\s*\d+\s*-\s*\d+.*$", "", cleaned)
    return cleaned.strip(" -－—:：、")


def canonical_work_item_name(text: str) -> str:
    cleaned = to_work_item_name(strip_toc_row_noise(text))
    normalized = normalize_match_text(cleaned)
    normalized = normalized.replace("舖面", "鋪面")
    default_aliases = {
        "施工放樣": "放樣",
        "基礎放樣": "放樣",
        "放樣施工": "放樣",
        "施工開挖": "開挖",
        "開挖土方": "開挖",
        "土方開挖": "開挖",
        "施工回填": "回填",
        "回填土方": "回填",
        "土方回填": "回填",
        "施工便道": "便道",
        "便道施工": "便道",
        "施工便道設施": "便道",
        "便道設施": "便道",
        "擋土墻": "擋土牆",
        # 部分附錄（如附錄4）的目錄標題只寫「瀝青混凝土」，
        # 但使用者選的工項是完整的「瀝青混凝土鋪面工程」，
        # 兩者視為同一工項，避免該工項被當成找不到而漏抓。
        "瀝青混凝土": "瀝青混凝土鋪面工程",
        # 附錄4目錄寫成「土包袋程」（疑似「土包袋工程」誤植漏字），
        # 附錄6目錄則寫成「土包袋收邊工程」，
        # 與使用者選的「土包袋」視為同一工項，避免漏抓。
        "土包袋程": "土包袋",
        "土包袋收邊": "土包袋",
    }
    if normalized in ("瀝青混凝土鋪面", "混凝土鋪面"):
        return f"{normalized}工程"
    return default_aliases.get(normalized, normalized)


def work_item_identity_key(text: str) -> str:
    cleaned = to_work_item_name(strip_toc_row_noise(text))
    return normalize_match_text(cleaned).replace("舖面", "鋪面")


def unique_work_items(work_items: list[str]) -> list[str]:
    unique_items = []
    seen = set()
    for work_item in work_items:
        cleaned = to_work_item_name(strip_toc_row_noise(work_item))
        identity_key = work_item_identity_key(cleaned)
        if not cleaned or not identity_key or identity_key in seen:
            continue
        seen.add(identity_key)
        unique_items.append(cleaned)
    return unique_items


def normalize_work_item_list(work_items: list[str]) -> list[tuple[str, str]]:
    return [
        (work_item, normalized_work_item)
        for work_item in dict.fromkeys(work_items)
        if (normalized_work_item := canonical_work_item_name(work_item))
    ]


def normalized_title_work_item(text: str) -> str:
    return canonical_work_item_name(text)


def normalized_work_item_matches_title(
    normalized_title: str,
    normalized_work_item: str,
) -> bool:
    if not normalized_title or not normalized_work_item:
        return False
    if normalized_title == normalized_work_item:
        return True

    return False


def normalized_section_work_item_matches_title(
    normalized_title: str,
    normalized_work_item: str,
) -> bool:
    return normalized_work_item_matches_title(normalized_title, normalized_work_item)


def normalized_text_contains_work_item_near_start(
    normalized_text: str,
    normalized_work_item: str,
) -> bool:
    if not normalized_text or not normalized_work_item:
        return False

    # 只允許「真正從開頭開始」的工項名稱，避免把
    # 「瀝青混凝土鋪面工程」誤判成「混凝土鋪面工程」。
    leading_text = re.sub(r"^(?:[A-Fa-f]\d{1,3}|\d{1,3})", "", normalized_text)
    return leading_text.startswith(normalized_work_item)


def work_item_from_near_start_text(
    text: str,
    normalized_work_items: list[tuple[str, str]],
) -> str | None:
    normalized_text = normalize_match_text(text).replace("舖面", "鋪面")
    matches = [
        (len(normalized_work_item), work_item)
        for work_item, normalized_work_item in normalized_work_items
        if normalized_text_contains_work_item_near_start(
            normalized_text,
            normalized_work_item,
        )
    ]
    if not matches:
        return None
    return max(matches, key=lambda item: item[0])[1]


def work_item_title_matches(text: str, work_item: str) -> bool:
    normalized_title = normalized_title_work_item(text)
    normalized_work_item = canonical_work_item_name(work_item)
    return normalized_work_item_matches_title(normalized_title, normalized_work_item)


def work_item_matches(text: str, work_item: str) -> bool:
    return work_item_title_matches(text, work_item)


def same_work_item_identity(left: str, right: str) -> bool:
    left_key = work_item_identity_key(left)
    right_key = work_item_identity_key(right)
    return bool(left_key and right_key and left_key == right_key)


def same_work_item_name(left: str, right: str) -> bool:
    if same_work_item_identity(left, right):
        return True
    return canonical_work_item_name(left) == canonical_work_item_name(right)


def is_default_work_item_option(work_item: str) -> bool:
    return any(
        same_work_item_name(work_item, default_work_item)
        for default_work_item in DEFAULT_WORK_ITEMS
    )


def is_toc_header_row(text: str) -> bool:
    return "目錄" in text or ("頁碼" in text and ("序號" in text or "名稱" in text))


def row_matches_work_item(text: str, work_items: list[str]) -> bool:
    return any(work_item_matches(text, work_item) for work_item in work_items)


def toc_entry_matches_work_item_identity(entry_text: str, work_item: str) -> bool:
    entry_work_item = to_work_item_name(strip_toc_row_noise(entry_text))
    return same_work_item_identity(entry_work_item, work_item)


def read_document_root(docx_bytes: bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as docx_package:
        return etree.fromstring(docx_package.read("word/document.xml"))


def xml_element_text(element) -> str:
    pieces = []
    for node in element.iter():
        if node.tag == W_T and node.text:
            pieces.append(node.text)
        elif node.tag == f"{{{WORD_NAMESPACE}}}tab":
            pieces.append("\t")
    return "".join(pieces).strip()


XML_DRAWING_OBJECT_TAGS = {
    f"{{{WORD_NAMESPACE}}}drawing",
    f"{{{WORD_NAMESPACE}}}pict",
    f"{{{WORD_NAMESPACE}}}object",
}


def xml_direct_text(element) -> str:
    """
    取得段落「本身」的文字，排除流程圖等內嵌物件（drawing/pict/object，
    例如文字方塊）裡的文字，避免標題段落因為附帶流程圖而被誤判成
    過長文字、導致標題比對失敗。
    """

    pieces = []

    def walk(node):
        if node.tag in XML_DRAWING_OBJECT_TAGS:
            return
        if node.tag == W_T and node.text:
            pieces.append(node.text)
        elif node.tag == f"{{{WORD_NAMESPACE}}}tab":
            pieces.append("\t")
        for child in node:
            walk(child)

    walk(element)
    return "".join(pieces).strip()


def xml_paragraph_style(element) -> str:
    style = element.find("./w:pPr/w:pStyle", namespaces=WORD_NAMESPACES)
    if style is None:
        return ""
    return style.get(f"{{{WORD_NAMESPACE}}}val", "")


def xml_row_text(row) -> str:
    return xml_element_text(row)


def xml_is_toc_table(element) -> bool:
    text = xml_element_text(element)
    if not text:
        return False

    has_toc_header = "頁碼" in text and ("序號" in text or "名稱" in text)
    has_appendix_page = bool(re.search(r"附錄\s*\d+\s*-\s*\d+", text))
    rows = element.xpath("./w:tr", namespaces=WORD_NAMESPACES)
    has_toc_rows = sum(
        1
        for row in rows
        if any(suffix in xml_row_text(row) for suffix in TOC_WORK_ITEM_SUFFIXES)
    )
    return has_toc_header and (has_appendix_page or has_toc_rows >= 2)


def xml_is_toc_block(element) -> bool:
    if element.tag == W_TBL:
        return xml_is_toc_table(element)
    if element.tag != W_P:
        return False
    return is_toc_candidate(xml_element_text(element), xml_paragraph_style(element))


def xml_leading_table_text(element) -> str:
    texts = []
    for paragraph in element.xpath(".//w:p", namespaces=WORD_NAMESPACES):
        text = xml_element_text(paragraph)
        if text:
            texts.append(text)
            if len(texts) >= 8:
                return " ".join(texts)
    return " ".join(texts)


def xml_title_candidate_texts(element) -> list[tuple[str, str]]:
    if element.tag == W_P:
        return [(xml_direct_text(element), xml_paragraph_style(element))]

    title_candidates = []
    for paragraph in element.xpath(".//w:p", namespaces=WORD_NAMESPACES):
        text = xml_direct_text(paragraph)
        if text:
            title_candidates.append((text, xml_paragraph_style(paragraph)))
    return title_candidates


def text_looks_like_section_title(
    text: str,
    style_name: str = "",
    require_suffix: bool = False,
) -> bool:
    cleaned_text = clean_toc_text(text)
    if not cleaned_text or len(cleaned_text) > 160:
        return False

    has_title_suffix = any(suffix in cleaned_text for suffix in TOC_WORK_ITEM_SUFFIXES)
    if require_suffix:
        return has_title_suffix

    has_title_style = style_name.startswith("Heading") or style_name.startswith("標題")
    return has_title_suffix or has_title_style


def xml_section_title_work_items(element) -> list[str]:
    work_items = []
    for text, style_name in xml_title_candidate_texts(element):
        if element.tag == W_P and is_toc_candidate(text, style_name):
            continue
        if not text_looks_like_section_title(
            text,
            style_name,
            require_suffix=(element.tag == W_TBL),
        ):
            continue

        work_item = to_work_item_name(clean_toc_text(text))
        if work_item:
            work_items.append(work_item)
    return unique_work_items(work_items)


def xml_leading_block_text(element) -> str:
    if element.tag == W_P:
        return xml_direct_text(element)
    if element.tag == W_TBL:
        return xml_leading_table_text(element)
    return xml_element_text(element)


def xml_detect_section_work_item(
    element,
    normalized_work_items: list[tuple[str, str]],
) -> str | None:
    detected_work_item = xml_detect_any_section_work_item(element)
    if detected_work_item:
        normalized_title = canonical_work_item_name(detected_work_item)
        matched_work_items = [
            work_item
            for work_item, normalized_work_item in normalized_work_items
            if normalized_section_work_item_matches_title(
                normalized_title,
                normalized_work_item,
            )
        ]
        if matched_work_items:
            return max(matched_work_items, key=len)

    leading_match = work_item_from_near_start_text(
        xml_leading_block_text(element),
        normalized_work_items,
    )
    if leading_match:
        return leading_match

    return None


def xml_detect_any_section_work_item(element) -> str | None:
    work_items = xml_section_title_work_items(element)
    return work_items[0] if work_items else None


def xml_set_cell_text_preserving_first_run(cell, text: str) -> None:
    text_nodes = cell.xpath(".//w:t", namespaces=WORD_NAMESPACES)
    if text_nodes:
        text_nodes[0].text = text
        for node in text_nodes[1:]:
            node.text = ""
        return

    paragraph = cell.find("./w:p", namespaces=WORD_NAMESPACES)
    if paragraph is None:
        paragraph = etree.SubElement(cell, f"{{{WORD_NAMESPACE}}}p")

    run = paragraph.find("./w:r", namespaces=WORD_NAMESPACES)
    if run is None:
        run = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}r")

    text_node = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
    text_node.text = text


def xml_apply_font_to_cell(cell, font_name: str = "標楷體") -> None:
    for run in cell.xpath(".//w:r", namespaces=WORD_NAMESPACES):
        run_properties = run.find("./w:rPr", namespaces=WORD_NAMESPACES)
        if run_properties is None:
            run_properties = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
            run.insert(0, run_properties)

        run_fonts = run_properties.find("./w:rFonts", namespaces=WORD_NAMESPACES)
        if run_fonts is None:
            run_fonts = etree.Element(f"{{{WORD_NAMESPACE}}}rFonts")
            run_properties.insert(0, run_fonts)

        for attribute_name in ("ascii", "hAnsi", "eastAsia", "cs"):
            run_fonts.set(f"{{{WORD_NAMESPACE}}}{attribute_name}", font_name)


def xml_set_cell_text_and_font(cell, text: str, font_name: str = "標楷體") -> None:
    xml_set_cell_text_preserving_first_run(cell, text)
    xml_apply_font_to_cell(cell, font_name)


def appendix_code_for_section(appendix_number: int, section_index: int) -> str:
    if 1 <= appendix_number <= len(APPENDIX_CODE_PREFIXES):
        prefix = APPENDIX_CODE_PREFIXES[appendix_number - 1]
    else:
        prefix = chr(ord("A") + max(0, appendix_number - 1))
    return f"{prefix}{section_index:02d}"


def xml_cell_is_empty_or_old_appendix_code(cell) -> bool:
    text = normalize_match_text(xml_element_text(cell))
    if not text:
        return True
    return bool(re.fullmatch(r"[A-Fa-f]?\d{1,3}", text) or re.fullmatch(r"[A-Fa-f]\d{2}", text))


def xml_cell_is_appendix_code_label(cell) -> bool:
    text = normalize_match_text(xml_element_text(cell))
    return text in APPENDIX_CODE_LABEL_TEXTS


def xml_appendix_code_label_text(cell, code_text: str) -> str:
    original_text = xml_element_text(cell).strip()
    if not original_text:
        return code_text

    new_text = re.sub(
        r"([^\r\n]*?編\s*號)\s*[：:]+\s*(?:[A-Za-z]?\d{1,3}\s*)*$",
        rf"\g<1>：{code_text}",
        original_text,
    )
    if new_text != original_text:
        return new_text

    return f"{original_text}：{code_text}"


def xml_replace_code_text_inside_cell(cell, code_text: str) -> bool:
    original_text = xml_element_text(cell)
    if "編號" not in original_text:
        return False

    compact_text = normalize_match_text(original_text)
    if compact_text in APPENDIX_CODE_LABEL_TEXTS:
        return False

    new_text = re.sub(
        r"([^\r\n]*?編\s*號)\s*[：:]+\s*(?:[A-Za-z]?\d{1,3}\s*)*",
        rf"\g<1>：{code_text}",
        original_text,
        count=1,
    )
    if new_text == original_text:
        return False

    xml_set_cell_text_and_font(cell, new_text)
    return True


def xml_fill_appendix_code_in_record_table(
    element_xml: bytes,
    appendix_number: int,
    section_index: int,
) -> bytes:
    """
    附錄二、附錄五的正文表格如果有「編號」欄位，依選取順序填入：
    附錄二 = B01、B02...
    附錄五 = E01、E02...
    """
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    if element.tag != W_TBL:
        return element_xml

    code_text = appendix_code_for_section(appendix_number, section_index)
    changed = False
    label_adjacent_filled = False

    rows = element.xpath("./w:tr", namespaces=WORD_NAMESPACES)

    # 1) 常見格式：左格是「編號」，右格是空白或舊編號。
    for row in rows:
        cells = row.xpath("./w:tc", namespaces=WORD_NAMESPACES)
        for cell_index, cell in enumerate(cells):
            if xml_cell_is_appendix_code_label(cell):
                if cell_index + 1 < len(cells):
                    target_cell = cells[cell_index + 1]
                    if xml_cell_is_empty_or_old_appendix_code(target_cell):
                        xml_set_cell_text_and_font(target_cell, code_text)
                        changed = True
                        label_adjacent_filled = True
                else:
                    xml_set_cell_text_and_font(
                        cell,
                        xml_appendix_code_label_text(cell, code_text),
                    )
                    changed = True
                    label_adjacent_filled = True
            elif xml_replace_code_text_inside_cell(cell, code_text):
                changed = True

    # 2) 另一種格式：第一列某欄是「編號」，下面資料列同欄空白。
    #    如果前面已經填過「編號右側欄位」，就不要再把下方明細列誤填一次。
    if label_adjacent_filled:
        return etree.tostring(element) if changed else element_xml

    for row_index, row in enumerate(rows):
        cells = row.xpath("./w:tc", namespaces=WORD_NAMESPACES)
        code_column_indexes = [
            cell_index
            for cell_index, cell in enumerate(cells)
            if xml_cell_is_appendix_code_label(cell)
        ]
        if not code_column_indexes:
            continue

        for data_row in rows[row_index + 1 :]:
            data_cells = data_row.xpath("./w:tc", namespaces=WORD_NAMESPACES)
            for code_column_index in code_column_indexes:
                if code_column_index >= len(data_cells):
                    continue
                target_cell = data_cells[code_column_index]
                if xml_cell_is_empty_or_old_appendix_code(target_cell):
                    xml_set_cell_text_and_font(target_cell, code_text)
                    changed = True
            # 只補第一筆資料列，避免把檢查項目明細列全部誤填。
            break

    return etree.tostring(element) if changed else element_xml


def xml_fill_project_name_in_record_table(element_xml: bytes, project_name: str) -> bytes:
    project_name = project_name.strip()
    if not project_name:
        return element_xml

    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    if element.tag != W_TBL:
        return element_xml

    for row in element.xpath("./w:tr", namespaces=WORD_NAMESPACES):
        cells = row.xpath("./w:tc", namespaces=WORD_NAMESPACES)
        for cell_index, cell in enumerate(cells):
            if normalize_match_text(xml_element_text(cell)) != "工程名稱":
                continue

            if cell_index + 1 < len(cells):
                xml_set_cell_text_and_font(cells[cell_index + 1], project_name)
                return etree.tostring(element)

            xml_set_cell_text_and_font(cell, f"工程名稱：{project_name}")
            return etree.tostring(element)

    return element_xml


def xml_extract_toc_source(root):
    body = root.find(".//w:body", namespaces=WORD_NAMESPACES)
    if body is None:
        return None

    children = [child for child in body if child.tag != W_SECTPR]
    for index, element in enumerate(children):
        if element.tag != W_TBL or not xml_is_toc_table(element):
            continue

        prefix_xml = []
        for previous_element in children[max(0, index - 4) : index]:
            if previous_element.tag == W_P and "目錄" in xml_element_text(previous_element):
                prefix_xml.append(etree.tostring(previous_element))

        suffix_xml = []
        next_index = index + 1
        if next_index < len(children):
            next_element = children[next_index]
            if next_element.tag == W_P and not xml_element_text(next_element):
                suffix_xml.append(etree.tostring(next_element))

        return {
            "type": "table",
            "prefix_xml": prefix_xml,
            "table_xml": etree.tostring(element),
            "suffix_xml": suffix_xml,
        }

    prefix_xml = []
    entries = []
    suffix_xml = []
    toc_started = False

    for element in children:
        if element.tag != W_P:
            if toc_started:
                break
            continue

        text = xml_element_text(element)
        style_name = xml_paragraph_style(element)

        if "目錄" in text:
            toc_started = True
            prefix_xml.append(etree.tostring(element))
            continue

        if not toc_started:
            continue

        if not text:
            suffix_xml.append(etree.tostring(element))
            continue

        if not is_toc_candidate(text, style_name):
            if is_toc_page_separator(text):
                suffix_xml.append(etree.tostring(element))
                continue
            break

        if re.search(r"\bTOC\s*\\", text, flags=re.IGNORECASE):
            continue

        entries.append({"xml": etree.tostring(element), "text": text})

    if entries:
        return {
            "type": "paragraphs",
            "prefix_xml": prefix_xml,
            "entries": entries,
            "suffix_xml": suffix_xml[:1],
        }

    return None


def toc_entry_text_with_display_work_item(
    entry_text: str,
    actual_work_item: str,
    display_work_item: str,
) -> str:
    if same_work_item_name(actual_work_item, display_work_item):
        return entry_text

    cleaned_entry = strip_toc_row_noise(entry_text)
    actual_name = to_work_item_name(cleaned_entry)
    if actual_name and cleaned_entry.startswith(actual_name):
        return f"{display_work_item}{cleaned_entry[len(actual_name):]}"

    for suffix in TOC_WORK_ITEM_SUFFIXES:
        if cleaned_entry.endswith(suffix):
            return f"{display_work_item}{suffix}"

    return cleaned_entry.replace(actual_work_item, display_work_item, 1)


def toc_entry_text_from_template(
    source_entry_texts: list[str],
    display_work_item: str,
) -> str:
    for entry_text in source_entry_texts:
        template_work_item = to_work_item_name(strip_toc_row_noise(entry_text))
        if template_work_item:
            return toc_entry_text_with_display_work_item(
                entry_text,
                template_work_item,
                display_work_item,
            )
    return display_work_item


def xml_filter_toc_elements(
    toc_source,
    work_items: list[str],
    appendix_number: int,
    bookmark_names: list[str] | None = None,
    display_work_items: list[str] | None = None,
) -> list[bytes]:
    if not toc_source:
        return []

    source_entry_texts = toc_entry_texts(toc_source)
    selected_entry_texts = []
    used_entry_indexes = set()
    for work_item_index, work_item in enumerate(work_items):
        display_work_item = (
            display_work_items[work_item_index]
            if display_work_items and work_item_index < len(display_work_items)
            else work_item
        )
        allow_reused_source = not same_work_item_name(work_item, display_work_item)
        matched_entry = None
        for matcher in (
            toc_entry_matches_work_item_identity,
            lambda entry_text, selected_work_item: row_matches_work_item(
                entry_text,
                [selected_work_item],
            ),
        ):
            for entry_index, entry_text in enumerate(source_entry_texts):
                if entry_index in used_entry_indexes and not allow_reused_source:
                    continue
                if matcher(entry_text, work_item):
                    matched_entry = (entry_index, entry_text)
                    break
            if matched_entry is not None:
                break

        if matched_entry is not None:
            entry_index, entry_text = matched_entry
            selected_entry_texts.append(
                toc_entry_text_with_display_work_item(
                    entry_text,
                    work_item,
                    display_work_item,
                )
            )
            if not allow_reused_source:
                used_entry_indexes.add(entry_index)
            continue

        selected_entry_texts.append(
            toc_entry_text_from_template(source_entry_texts, display_work_item)
        )

    if not selected_entry_texts:
        return []

    return (
        toc_source["prefix_xml"]
        + [xml_generated_toc_table(selected_entry_texts, appendix_number, bookmark_names)]
        + toc_source["suffix_xml"]
    )


def xml_text_paragraph(
    text: str,
    bold: bool = False,
    align: str | None = None,
    font_name: str | None = None,
):
    paragraph = etree.Element(f"{{{WORD_NAMESPACE}}}p")
    if align:
        paragraph_properties = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}pPr")
        justification = etree.SubElement(paragraph_properties, f"{{{WORD_NAMESPACE}}}jc")
        justification.set(f"{{{WORD_NAMESPACE}}}val", align)

    run = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}r")
    if bold or font_name:
        run_properties = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}rPr")
        if font_name:
            run_fonts = etree.SubElement(run_properties, f"{{{WORD_NAMESPACE}}}rFonts")
            for attribute_name in ("ascii", "hAnsi", "eastAsia", "cs"):
                run_fonts.set(f"{{{WORD_NAMESPACE}}}{attribute_name}", font_name)
        if bold:
            etree.SubElement(run_properties, f"{{{WORD_NAMESPACE}}}b")
    text_node = etree.SubElement(run, f"{{{WORD_NAMESPACE}}}t")
    text_node.text = text
    return paragraph


def xml_toc_cell(text: str, width: int, bold: bool = False, shaded: bool = False):
    cell = etree.Element(f"{{{WORD_NAMESPACE}}}tc")
    cell_properties = etree.SubElement(cell, f"{{{WORD_NAMESPACE}}}tcPr")
    cell_width = etree.SubElement(cell_properties, f"{{{WORD_NAMESPACE}}}tcW")
    cell_width.set(f"{{{WORD_NAMESPACE}}}w", str(width))
    cell_width.set(f"{{{WORD_NAMESPACE}}}type", "dxa")
    if shaded:
        shading = etree.SubElement(cell_properties, f"{{{WORD_NAMESPACE}}}shd")
        shading.set(f"{{{WORD_NAMESPACE}}}fill", "D9D9D9")
    cell.append(xml_text_paragraph(text, bold=bold, align="center", font_name="標楷體"))
    return cell


def xml_toc_page_reference_cell(
    bookmark_name: str,
    appendix_number: int,
    fallback_page_number: int,
    width: int,
):
    cell = etree.Element(f"{{{WORD_NAMESPACE}}}tc")
    cell_properties = etree.SubElement(cell, f"{{{WORD_NAMESPACE}}}tcPr")
    cell_width = etree.SubElement(cell_properties, f"{{{WORD_NAMESPACE}}}tcW")
    cell_width.set(f"{{{WORD_NAMESPACE}}}w", str(width))
    cell_width.set(f"{{{WORD_NAMESPACE}}}type", "dxa")

    paragraph = etree.SubElement(cell, f"{{{WORD_NAMESPACE}}}p")
    paragraph_properties = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}pPr")
    justification = etree.SubElement(paragraph_properties, f"{{{WORD_NAMESPACE}}}jc")
    justification.set(f"{{{WORD_NAMESPACE}}}val", "center")

    prefix_run = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}r")
    run_properties = etree.SubElement(prefix_run, f"{{{WORD_NAMESPACE}}}rPr")
    run_fonts = etree.SubElement(run_properties, f"{{{WORD_NAMESPACE}}}rFonts")
    for attribute_name in ("ascii", "hAnsi", "eastAsia", "cs"):
        run_fonts.set(f"{{{WORD_NAMESPACE}}}{attribute_name}", "標楷體")
    prefix_text = etree.SubElement(prefix_run, f"{{{WORD_NAMESPACE}}}t")
    prefix_text.text = f"附錄 {appendix_number}-"

    page_field = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}fldSimple")
    page_field.set(f"{{{WORD_NAMESPACE}}}instr", f" PAGEREF {bookmark_name} \\h ")
    page_field.set(f"{{{WORD_NAMESPACE}}}dirty", "true")
    page_run = etree.SubElement(page_field, f"{{{WORD_NAMESPACE}}}r")
    page_run_properties = etree.SubElement(page_run, f"{{{WORD_NAMESPACE}}}rPr")
    page_run_fonts = etree.SubElement(page_run_properties, f"{{{WORD_NAMESPACE}}}rFonts")
    for attribute_name in ("ascii", "hAnsi", "eastAsia", "cs"):
        page_run_fonts.set(f"{{{WORD_NAMESPACE}}}{attribute_name}", "標楷體")
    page_text = etree.SubElement(page_run, f"{{{WORD_NAMESPACE}}}t")
    page_text.text = str(fallback_page_number)

    return cell


def xml_toc_row(values: list[str], widths: list[int], bold: bool = False, shaded: bool = False):
    row = etree.Element(f"{{{WORD_NAMESPACE}}}tr")
    for value, width in zip(values, widths):
        row.append(xml_toc_cell(value, width, bold=bold, shaded=shaded))
    return row


def xml_set_all_text_font(element_xml: bytes, font_name: str) -> bytes:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    for run in element.xpath(".//w:r", namespaces=WORD_NAMESPACES):
        run_properties = run.find("./w:rPr", namespaces=WORD_NAMESPACES)
        if run_properties is None:
            run_properties = etree.Element(f"{{{WORD_NAMESPACE}}}rPr")
            run.insert(0, run_properties)

        run_fonts = run_properties.find("./w:rFonts", namespaces=WORD_NAMESPACES)
        if run_fonts is None:
            run_fonts = etree.Element(f"{{{WORD_NAMESPACE}}}rFonts")
            run_properties.insert(0, run_fonts)

        for attribute_name in ("ascii", "hAnsi", "eastAsia", "cs"):
            run_fonts.set(f"{{{WORD_NAMESPACE}}}{attribute_name}", font_name)

    return etree.tostring(element)


def xml_replace_work_item_text(
    element_xml: bytes,
    actual_work_item: str,
    requested_work_item: str,
) -> bytes:
    """
    把備援抓到的工項名稱改成使用者實際選的工項名稱。
    例如：
    原本模板：混凝土鋪面工程
    使用者選：瀝青混凝土鋪面工程
    輸出時改成：瀝青混凝土鋪面工程

    同時處理 Word 文字被拆成多個 run 的情況。
    """
    if same_work_item_name(actual_work_item, requested_work_item):
        return element_xml

    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    actual_full = canonical_work_item_name(actual_work_item)
    requested_full = canonical_work_item_name(requested_work_item)

    actual_base = actual_full[:-2] if actual_full.endswith("工程") else actual_full
    requested_base = requested_full[:-2] if requested_full.endswith("工程") else requested_full

    replace_pairs = [
        (actual_full, requested_full),
        (actual_full.replace("鋪", "舖"), requested_full),
        (actual_base, requested_base),
        (actual_base.replace("鋪", "舖"), requested_base),
    ]

    def replace_work_item_once(text: str) -> str:
        for old_text, new_value in replace_pairs:
            if old_text and old_text in text:
                return text.replace(old_text, new_value)
        return text

    # 以段落合併文字處理，避免 Word 拆成不同 run 時漏換；
    # 每段只套用第一個命中的規則，避免「瀝青」被重複前綴。
    paragraphs = []
    if element.tag == W_P:
        paragraphs.append(element)
    paragraphs.extend(element.xpath(".//w:p", namespaces=WORD_NAMESPACES))

    for paragraph in paragraphs:
        text_nodes = paragraph.xpath(".//w:t", namespaces=WORD_NAMESPACES)
        if not text_nodes:
            continue

        combined_text = "".join(node.text or "" for node in text_nodes)
        new_combined_text = replace_work_item_once(combined_text)

        if new_combined_text != combined_text:
            text_nodes[0].text = new_combined_text
            for node in text_nodes[1:]:
                node.text = ""

    return etree.tostring(element)

def appendix_toc_name_header(appendix_number: int) -> str:
    if appendix_number in (2, 5):
        return "抽查紀錄表名稱"
    if appendix_number in (3, 6):
        return "抽查管理標準表名稱"
    return "抽查程序流程圖名稱"


def xml_generated_toc_table(
    entry_texts: list[str],
    appendix_number: int,
    bookmark_names: list[str] | None = None,
) -> bytes:
    table = etree.Element(f"{{{WORD_NAMESPACE}}}tbl")
    table_properties = etree.SubElement(table, f"{{{WORD_NAMESPACE}}}tblPr")
    table_alignment = etree.SubElement(table_properties, f"{{{WORD_NAMESPACE}}}jc")
    table_alignment.set(f"{{{WORD_NAMESPACE}}}val", "center")
    table_width = etree.SubElement(table_properties, f"{{{WORD_NAMESPACE}}}tblW")
    table_width.set(f"{{{WORD_NAMESPACE}}}w", "9000")
    table_width.set(f"{{{WORD_NAMESPACE}}}type", "dxa")
    table_borders = etree.SubElement(table_properties, f"{{{WORD_NAMESPACE}}}tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = etree.SubElement(table_borders, f"{{{WORD_NAMESPACE}}}{border_name}")
        border.set(f"{{{WORD_NAMESPACE}}}val", "single")
        border.set(f"{{{WORD_NAMESPACE}}}sz", "4")
        border.set(f"{{{WORD_NAMESPACE}}}space", "0")
        border.set(f"{{{WORD_NAMESPACE}}}color", "000000")

    widths = [1200, 6100, 1700]
    table.append(
        xml_toc_row(
            ["序號", appendix_toc_name_header(appendix_number), "頁碼"],
            widths,
            bold=True,
            shaded=True,
        )
    )

    code_prefix = APPENDIX_CODE_PREFIXES[appendix_number - 1]
    for index, entry_text in enumerate(entry_texts, start=1):
        row = etree.Element(f"{{{WORD_NAMESPACE}}}tr")
        row.append(xml_toc_cell(f"{code_prefix}{index:02d}", widths[0]))
        row.append(xml_toc_cell(strip_toc_row_noise(entry_text), widths[1]))
        if bookmark_names and index <= len(bookmark_names):
            row.append(
                xml_toc_page_reference_cell(
                    bookmark_names[index - 1],
                    appendix_number,
                    index,
                    widths[2],
                )
            )
        else:
            row.append(xml_toc_cell(f"附錄 {appendix_number}-{index}", widths[2]))
        table.append(row)

    return etree.tostring(table)


def xml_set_page_break_before(element_xml: bytes) -> bytes:
    """
    在元素的第一個段落加上 pageBreakBefore，讓這個元素（表格或段落）
    強制從新的一頁開始。

    跟「插入一個只含手動分頁符號的空白段落」不同，pageBreakBefore
    是段落屬性、本身不佔版面：如果前一個工項的表格剛好填滿整頁，
    這個段落本來就會自動換到下一頁，pageBreakBefore 不會再多插入
    一頁空白頁；如果前一個表格沒填滿頁面，則會強制換頁，效果跟手動
    分頁符號相同。藉此避免「跑板」時某些工項之間多出一張空白頁。
    """
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    if element.tag == W_P:
        paragraph = element
    else:
        paragraph = element.find(".//w:p", namespaces=WORD_NAMESPACES)
        if paragraph is None:
            return element_xml

    paragraph_properties = paragraph.find("./w:pPr", namespaces=WORD_NAMESPACES)
    if paragraph_properties is None:
        paragraph_properties = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
        paragraph.insert(0, paragraph_properties)

    if paragraph_properties.find("./w:pageBreakBefore", namespaces=WORD_NAMESPACES) is None:
        page_break_before = etree.Element(f"{{{WORD_NAMESPACE}}}pageBreakBefore")
        paragraph_properties.insert(0, page_break_before)

    return etree.tostring(element)


def xml_section_break_next_page_paragraph() -> bytes:
    paragraph = etree.Element(f"{{{WORD_NAMESPACE}}}p")
    paragraph_properties = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}pPr")
    section_properties = etree.SubElement(
        paragraph_properties,
        f"{{{WORD_NAMESPACE}}}sectPr",
    )
    break_type = etree.SubElement(section_properties, f"{{{WORD_NAMESPACE}}}type")
    break_type.set(f"{{{WORD_NAMESPACE}}}val", "nextPage")
    return etree.tostring(paragraph)


def xml_add_bookmark_to_element(
    element_xml: bytes,
    bookmark_name: str,
    bookmark_id: int,
) -> bytes:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    if element.tag == W_P:
        paragraph = element
    else:
        paragraphs = element.xpath(".//w:p", namespaces=WORD_NAMESPACES)
        if not paragraphs:
            return element_xml
        paragraph = paragraphs[0]

    bookmark_start = etree.Element(f"{{{WORD_NAMESPACE}}}bookmarkStart")
    bookmark_start.set(f"{{{WORD_NAMESPACE}}}id", str(bookmark_id))
    bookmark_start.set(f"{{{WORD_NAMESPACE}}}name", bookmark_name)
    bookmark_end = etree.Element(f"{{{WORD_NAMESPACE}}}bookmarkEnd")
    bookmark_end.set(f"{{{WORD_NAMESPACE}}}id", str(bookmark_id))

    insert_index = 1 if len(paragraph) and paragraph[0].tag.endswith("}pPr") else 0
    paragraph.insert(insert_index, bookmark_start)
    paragraph.insert(insert_index + 1, bookmark_end)
    return etree.tostring(element)


def xml_add_bookmark_to_element_end(
    element_xml: bytes,
    bookmark_name: str,
    bookmark_id: int,
) -> bytes:
    """Add a zero-length bookmark at the end of an element.

    This is used for TOC page ranges.  A large record table can span multiple
    Word pages, so the end bookmark must be placed near the last paragraph of
    the section, not at the start of the table.
    """
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    if element.tag == W_P:
        paragraph = element
    else:
        paragraphs = element.xpath(".//w:p", namespaces=WORD_NAMESPACES)
        if not paragraphs:
            return element_xml
        paragraph = paragraphs[-1]

    bookmark_start = etree.Element(f"{{{WORD_NAMESPACE}}}bookmarkStart")
    bookmark_start.set(f"{{{WORD_NAMESPACE}}}id", str(bookmark_id))
    bookmark_start.set(f"{{{WORD_NAMESPACE}}}name", bookmark_name)
    bookmark_end = etree.Element(f"{{{WORD_NAMESPACE}}}bookmarkEnd")
    bookmark_end.set(f"{{{WORD_NAMESPACE}}}id", str(bookmark_id))

    paragraph.append(bookmark_start)
    paragraph.append(bookmark_end)
    return etree.tostring(element)


def xml_has_visible_content(element) -> bool:
    if xml_element_text(element).strip():
        return True
    return bool(
        element.xpath(
            ".//w:drawing | .//w:pict | .//w:object",
            namespaces=WORD_NAMESPACES,
        )
    )


def xml_has_document_object(element) -> bool:
    if element.tag == W_TBL:
        return True
    return bool(
        element.xpath(
            ".//w:drawing | .//w:pict | .//w:object",
            namespaces=WORD_NAMESPACES,
        )
    )


def xml_has_section_break(element) -> bool:
    return element.find(".//w:sectPr", namespaces=WORD_NAMESPACES) is not None


def xml_contains_page_boundary(element_xml: bytes) -> bool:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return False
    has_page_break = bool(
        element.xpath(".//w:br[@w:type='page']", namespaces=WORD_NAMESPACES)
    )
    has_section_break = element.find(".//w:sectPr", namespaces=WORD_NAMESPACES) is not None
    return has_page_break or has_section_break


def xml_needs_page_break_after_toc(toc_elements: list[bytes]) -> bool:
    return not any(xml_contains_page_boundary(element_xml) for element_xml in toc_elements)


def xml_force_section_break_next_page(element_xml: bytes) -> bytes:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    changed = False
    for section_properties in element.xpath(".//w:sectPr", namespaces=WORD_NAMESPACES):
        break_type = section_properties.find("./w:type", namespaces=WORD_NAMESPACES)
        if break_type is None:
            break_type = etree.Element(f"{{{WORD_NAMESPACE}}}type")
            section_properties.insert(0, break_type)
        break_type.set(f"{{{WORD_NAMESPACE}}}val", "nextPage")
        changed = True

    return etree.tostring(element) if changed else element_xml


def toc_entry_texts(toc_source) -> list[str]:
    if not toc_source:
        return []

    if toc_source["type"] == "table":
        table = etree.fromstring(toc_source["table_xml"])
        return [
            xml_row_text(row)
            for row in table.xpath("./w:tr", namespaces=WORD_NAMESPACES)
            if not is_toc_header_row(xml_row_text(row))
        ]

    return [entry["text"] for entry in toc_source["entries"]]


def toc_source_work_items(toc_source) -> list[str]:
    return unique_work_items(
        [
            to_work_item_name(strip_toc_row_noise(entry_text))
            for entry_text in toc_entry_texts(toc_source)
        ]
    )


def xml_chunk_has_visible_content(elements: list[bytes]) -> bool:
    for element_xml in elements:
        try:
            element = etree.fromstring(element_xml)
        except Exception:
            continue
        if xml_has_visible_content(element):
            return True
    return False


def xml_chunk_has_substantive_content(elements: list[bytes]) -> bool:
    for element_xml in elements:
        try:
            element = etree.fromstring(element_xml)
        except Exception:
            continue
        if xml_has_document_object(element):
            return True
    return False


def assign_chunk_work_items(
    chunks: list[dict],
    toc_work_items: list[str],
) -> list[dict]:
    assigned_chunks = []
    toc_index = 0
    substantive_chunks = [
        chunk
        for chunk in chunks
        if xml_chunk_has_substantive_content(chunk.get("elements", []))
    ]

    for chunk in substantive_chunks:
        work_item = chunk.get("work_item")
        assignment_source = chunk.get("assignment_source")
        if work_item:
            for index in range(toc_index, len(toc_work_items)):
                if same_work_item_name(toc_work_items[index], work_item):
                    toc_index = index + 1
                    break
        else:
            while toc_index < len(toc_work_items):
                work_item = toc_work_items[toc_index]
                toc_index += 1
                if work_item:
                    assignment_source = "toc_order"
                    break

        if not work_item:
            continue

        assigned_chunk = dict(chunk)
        assigned_chunk["work_item"] = work_item
        assigned_chunk["assignment_source"] = assignment_source or "detected"
        assigned_chunks.append(assigned_chunk)

    return assigned_chunks


def xml_is_body_start_candidate(element, boundary_pairs: list[tuple[str, str]]) -> bool:
    if xml_detect_section_work_item(element, boundary_pairs):
        return True
    return xml_has_document_object(element)


def build_appendix_index_from_docx_bytes(
    docx_bytes: bytes,
    boundary_work_items: list[str],
):
    root = read_document_root(docx_bytes)
    body = root.find(".//w:body", namespaces=WORD_NAMESPACES)
    if body is None:
        return {"toc_source": None, "sections": []}

    toc_source = xml_extract_toc_source(root)
    toc_work_items = toc_source_work_items(toc_source)
    boundary_pairs = normalize_work_item_list(
        unique_work_items(toc_work_items + boundary_work_items)
    )
    chunks = []
    current_chunk = None
    toc_phase = toc_source is not None

    for element in body:
        if element.tag == W_SECTPR:
            continue

        if toc_phase:
            if (
                xml_is_toc_block(element)
                or "目錄" in xml_element_text(element)
                or not xml_has_visible_content(element)
            ):
                continue
            if not xml_is_body_start_candidate(element, boundary_pairs):
                continue
            toc_phase = False

        if xml_is_toc_block(element):
            continue

        detected_work_item = xml_detect_section_work_item(element, boundary_pairs)
        if (
            detected_work_item
            and current_chunk is not None
            and xml_chunk_has_visible_content(current_chunk["elements"])
        ):
            chunks.append(current_chunk)
            current_chunk = None

        if current_chunk is None:
            if not xml_has_visible_content(element):
                continue
            current_chunk = {
                "work_item": detected_work_item,
                "assignment_source": "detected" if detected_work_item else None,
                "elements": [],
            }

        current_chunk["elements"].append(etree.tostring(element))

        if current_chunk["work_item"] is None:
            current_chunk["work_item"] = detected_work_item
            if detected_work_item:
                current_chunk["assignment_source"] = "detected"

        if xml_has_section_break(element):
            if xml_chunk_has_visible_content(current_chunk["elements"]):
                chunks.append(current_chunk)
            current_chunk = None

    if current_chunk and xml_chunk_has_visible_content(current_chunk["elements"]):
        chunks.append(current_chunk)

    sections = assign_chunk_work_items(chunks, toc_work_items)

    return {"toc_source": toc_source, "sections": sections}


def section_with_requested_work_item(
    section: dict,
    requested_work_item: str,
) -> dict:
    selected_section = dict(section)
    selected_section["requested_work_item"] = requested_work_item
    return selected_section


def section_title_work_items(section: dict) -> list[str]:
    work_items = []
    for element_xml in section.get("elements", []):
        try:
            element = etree.fromstring(element_xml)
        except Exception:
            continue
        work_items.extend(xml_section_title_work_items(element))
    return unique_work_items(work_items)


def section_has_matching_title(section: dict, work_item: str) -> bool:
    return any(
        same_work_item_name(title_work_item, work_item)
        for title_work_item in section_title_work_items(section)
    )


def section_has_conflicting_title(section: dict, work_item: str) -> bool:
    title_work_items = []
    for element_xml in section.get("elements", []):
        try:
            element = etree.fromstring(element_xml)
        except Exception:
            continue
        if element.tag != W_P:
            continue
        title_work_items.extend(xml_section_title_work_items(element))
    title_work_items = unique_work_items(title_work_items)
    if not title_work_items:
        return False
    return not any(
        same_work_item_name(title_work_item, work_item)
        for title_work_item in title_work_items
    )


def element_title_work_items(element_xml: bytes) -> list[str]:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return []
    return xml_section_title_work_items(element)


def element_title_matches_work_item(element_xml: bytes, work_item: str) -> bool:
    return any(
        same_work_item_name(title_work_item, work_item)
        for title_work_item in element_title_work_items(element_xml)
    )


def element_title_conflicts_with_work_item(element_xml: bytes, work_item: str) -> bool:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return False
    if element.tag != W_P:
        return False

    title_work_items = element_title_work_items(element_xml)
    if not title_work_items:
        return False
    return not any(
        same_work_item_name(title_work_item, work_item)
        for title_work_item in title_work_items
    )


def xml_remove_page_breaks(element_xml: bytes) -> bytes:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    changed = False
    for page_break in element.xpath(".//w:br[@w:type='page']", namespaces=WORD_NAMESPACES):
        parent = page_break.getparent()
        if parent is not None:
            parent.remove(page_break)
            changed = True

    return etree.tostring(element) if changed else element_xml


def xml_element_has_visible_content(element_xml: bytes) -> bool:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return False
    return xml_has_visible_content(element)


def title_work_items_match_work_item(title_work_items: list[str], work_item: str) -> bool:
    return any(
        same_work_item_name(title_work_item, work_item)
        for title_work_item in title_work_items
    )


def xml_table_slice_for_work_item(
    element_xml: bytes,
    work_item: str,
    boundary_work_items: list[str] | None = None,
) -> bytes:
    try:
        element = etree.fromstring(element_xml)
    except Exception:
        return element_xml

    if element.tag != W_TBL:
        return element_xml

    # 只有「列標題」剛好對應到使用者選的工項時，才把它視為另一個工項
    # 表格的起點；否則表格內部的子項目列標（例如「材料進場抽查」）
    # 會被誤判成下一個工項的標題，導致整張表格被攔腰截斷。
    boundary_items = boundary_work_items if boundary_work_items else [work_item]

    rows = element.xpath("./w:tr", namespaces=WORD_NAMESPACES)
    title_rows = [
        (
            row_index,
            [
                title
                for title in xml_section_title_work_items(row)
                if any(same_work_item_name(title, item) for item in boundary_items)
            ],
        )
        for row_index, row in enumerate(rows)
    ]
    title_rows = [(row_index, titles) for row_index, titles in title_rows if titles]
    if not title_rows:
        return element_xml

    matching_title_rows = [
        (row_index, titles)
        for row_index, titles in title_rows
        if title_work_items_match_work_item(titles, work_item)
    ]
    if not matching_title_rows:
        return element_xml

    has_other_work_item = any(
        not title_work_items_match_work_item(titles, work_item)
        for _, titles in title_rows
    )
    if not has_other_work_item:
        return element_xml

    start_row_index = matching_title_rows[0][0]
    end_row_index = len(rows)
    for row_index, titles in title_rows:
        if row_index <= start_row_index:
            continue
        if not title_work_items_match_work_item(titles, work_item):
            end_row_index = row_index
            break

    sliced_element = copy.deepcopy(element)
    sliced_rows = sliced_element.xpath("./w:tr", namespaces=WORD_NAMESPACES)
    for row_index in range(len(sliced_rows) - 1, -1, -1):
        if row_index < start_row_index or row_index >= end_row_index:
            sliced_rows[row_index].getparent().remove(sliced_rows[row_index])

    return etree.tostring(sliced_element)


def selected_section_element_xmls(
    section: dict,
    boundary_work_items: list[str] | None = None,
) -> list[bytes]:
    requested_work_item = section.get("requested_work_item", section["work_item"])
    source_elements = list(section.get("elements", []))
    start_index = 0

    for element_index, element_xml in enumerate(source_elements):
        if element_title_matches_work_item(element_xml, requested_work_item):
            start_index = element_index
            break

    selected_elements = []
    for element_index, element_xml in enumerate(source_elements[start_index:]):
        if (
            element_index > 0
            and element_title_conflicts_with_work_item(element_xml, requested_work_item)
        ):
            break
        selected_elements.append(
            xml_table_slice_for_work_item(
                element_xml,
                requested_work_item,
                boundary_work_items,
            )
        )

    while selected_elements and not xml_element_has_visible_content(selected_elements[-1]):
        selected_elements.pop()

    if selected_elements:
        selected_elements[0] = xml_remove_page_breaks(selected_elements[0])

    return selected_elements


def select_indexed_sections(
    index_data,
    work_items: list[str],
    appendix_number: int | None = None,
) -> list[dict]:
    selected_sections = []
    used_section_indexes = set()
    sections = index_data.get("sections", [])
    for work_item in work_items:
        matched_index = None

        matched_by_title = False
        for section_index, section in enumerate(sections):
            if section_has_matching_title(section, work_item):
                matched_index = section_index
                matched_by_title = True
                break

        if appendix_number in (2, 4, 5, 6) and not matched_by_title:
            continue

        for section_index, section in enumerate(sections):
            if matched_index is not None:
                break
            if section_index in used_section_indexes:
                continue
            if section_has_conflicting_title(section, work_item):
                continue
            if same_work_item_identity(section["work_item"], work_item):
                matched_index = section_index
                break

        if matched_index is None:
            for section_index, section in enumerate(sections):
                if section_index in used_section_indexes:
                    continue
                if section_has_conflicting_title(section, work_item):
                    continue
                if same_work_item_name(section["work_item"], work_item):
                    matched_index = section_index
                    break

        if matched_index is None:
            continue

        selected_section = section_with_requested_work_item(
            sections[matched_index],
            work_item,
        )
        if matched_by_title:
            selected_section["work_item"] = work_item
        else:
            used_section_indexes.add(matched_index)
        selected_sections.append(selected_section)
    return selected_sections


def xml_generated_empty_footer() -> bytes:
    footer = etree.Element(f"{{{WORD_NAMESPACE}}}ftr")
    etree.SubElement(footer, f"{{{WORD_NAMESPACE}}}p")
    return etree.tostring(footer, xml_declaration=True, encoding="UTF-8", standalone=True)


def xml_generated_appendix_footer(appendix_number: int) -> bytes:
    footer = etree.Element(f"{{{WORD_NAMESPACE}}}ftr")
    paragraph = etree.SubElement(footer, f"{{{WORD_NAMESPACE}}}p")
    paragraph_properties = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}pPr")
    justification = etree.SubElement(paragraph_properties, f"{{{WORD_NAMESPACE}}}jc")
    justification.set(f"{{{WORD_NAMESPACE}}}val", "center")

    prefix_run = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}r")
    prefix_text = etree.SubElement(prefix_run, f"{{{WORD_NAMESPACE}}}t")
    prefix_text.text = f"附錄 {appendix_number}-"

    page_field = etree.SubElement(paragraph, f"{{{WORD_NAMESPACE}}}fldSimple")
    page_field.set(f"{{{WORD_NAMESPACE}}}instr", " PAGE \\* MERGEFORMAT ")
    page_run = etree.SubElement(page_field, f"{{{WORD_NAMESPACE}}}r")
    page_text = etree.SubElement(page_run, f"{{{WORD_NAMESPACE}}}t")
    page_text.text = "1"

    return etree.tostring(footer, xml_declaration=True, encoding="UTF-8", standalone=True)


def document_footer_relationships(source_zip: zipfile.ZipFile) -> list[tuple[str, str]]:
    try:
        relationships_xml = source_zip.read("word/_rels/document.xml.rels")
    except KeyError:
        return []

    root = etree.fromstring(relationships_xml)
    footer_relationships = []
    for relationship in root:
        relationship_type = relationship.get("Type", "")
        if not relationship_type.endswith("/footer"):
            continue
        relationship_id = relationship.get("Id")
        target = relationship.get("Target", "")
        if not relationship_id or not target:
            continue
        target_path = target.lstrip("/").replace("\\", "/")
        if not target_path.startswith("word/"):
            target_path = f"word/{target_path}"
        footer_relationships.append((relationship_id, target_path))

    return footer_relationships


def xml_remove_child_elements(parent, local_names: set[str]) -> None:
    for child in list(parent):
        if etree.QName(child).localname in local_names:
            parent.remove(child)


def xml_add_footer_reference(section_properties, footer_type: str, relationship_id: str) -> None:
    footer_reference = etree.Element(f"{{{WORD_NAMESPACE}}}footerReference")
    footer_reference.set(f"{{{WORD_NAMESPACE}}}type", footer_type)
    footer_reference.set(f"{{{RELATIONSHIP_NAMESPACE}}}id", relationship_id)
    section_properties.insert(0, footer_reference)


def xml_set_page_number_start(section_properties, start_number: int | None) -> None:
    xml_remove_child_elements(section_properties, {"pgNumType"})
    if start_number is None:
        return

    page_number_type = etree.Element(f"{{{WORD_NAMESPACE}}}pgNumType")
    page_number_type.set(f"{{{WORD_NAMESPACE}}}start", str(start_number))
    section_properties.insert(0, page_number_type)


def xml_apply_appendix_section_settings(
    root,
    empty_footer_rid: str | None,
    appendix_footer_rid: str | None,
) -> None:
    section_properties_list = root.xpath(".//w:sectPr", namespaces=WORD_NAMESPACES)
    if not section_properties_list:
        return

    for section_index, section_properties in enumerate(section_properties_list):
        xml_remove_child_elements(section_properties, {"footerReference"})
        if section_index == 0:
            xml_set_page_number_start(section_properties, None)
            if empty_footer_rid:
                xml_add_footer_reference(section_properties, "default", empty_footer_rid)
                xml_add_footer_reference(section_properties, "first", empty_footer_rid)
            continue

        xml_set_page_number_start(section_properties, 1 if section_index == 1 else None)
        if appendix_footer_rid:
            xml_add_footer_reference(section_properties, "default", appendix_footer_rid)
            xml_add_footer_reference(section_properties, "first", appendix_footer_rid)


def xml_enable_update_fields(settings_xml: bytes) -> bytes:
    root = etree.fromstring(settings_xml)
    update_fields = root.find("./w:updateFields", namespaces=WORD_NAMESPACES)
    if update_fields is None:
        update_fields = etree.SubElement(root, f"{{{WORD_NAMESPACE}}}updateFields")
    update_fields.set(f"{{{WORD_NAMESPACE}}}val", "true")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def refresh_docx_fields_with_word(
    docx_bytes: bytes,
    appendix_number: int | None = None,
    toc_page_range_bookmarks: list[tuple[str, str]] | None = None,
) -> bytes:
    """Use local Microsoft Word to repaginate and save PAGE/PAGEREF results.

    When toc_page_range_bookmarks is provided, the generated TOC page column is
    rewritten to a static range, for example: 附錄 2-6、附錄 2-7.
    """
    service_url = get_secret_setting("WORD_REFRESH_SERVICE_URL")
    service_token = get_secret_setting("WORD_REFRESH_SERVICE_TOKEN")

    timeout_seconds = 300
    try:
        timeout_seconds = int(
            os.environ.get(
                "WORD_REFRESH_SERVICE_TIMEOUT_SECONDS",
                st.secrets.get("WORD_REFRESH_SERVICE_TIMEOUT_SECONDS", 300),
            )
        )
    except Exception:
        timeout_seconds = 300

    if service_url:
        refreshed_bytes = refresh_docx_fields_via_service(
            docx_bytes,
            service_url,
            appendix_number,
            toc_page_range_bookmarks,
            service_token,
            timeout_seconds,
        )
        if refreshed_bytes is not None:
            return refreshed_bytes

    return refresh_docx_fields_with_local_word(
        docx_bytes,
        appendix_number,
        toc_page_range_bookmarks,
        WORD_COM_LOCK,
    )


def replace_document_body_with_xml(
    docx_bytes: bytes,
    body_element_xmls: list[bytes],
    appendix_number: int,
) -> bytes:
    source_stream = io.BytesIO(docx_bytes)
    output_stream = io.BytesIO()

    with zipfile.ZipFile(source_stream, "r") as source_zip:
        footer_relationships = document_footer_relationships(source_zip)
        empty_footer_rid = footer_relationships[0][0] if len(footer_relationships) > 1 else None
        appendix_footer_rid = (
            footer_relationships[1][0]
            if len(footer_relationships) > 1
            else footer_relationships[0][0]
            if footer_relationships
            else None
        )
        footer_targets = {path: rid for rid, path in footer_relationships}

        with zipfile.ZipFile(output_stream, "w") as output_zip:
            for item in source_zip.infolist():
                data = source_zip.read(item.filename)
                if item.filename == "word/document.xml":
                    root = etree.fromstring(data)
                    body = root.find(".//w:body", namespaces=WORD_NAMESPACES)
                    if body is not None:
                        section_properties = None
                        for child in list(body):
                            if child.tag == W_SECTPR:
                                section_properties = copy.deepcopy(child)
                            body.remove(child)

                        for element_xml in body_element_xmls:
                            body.append(etree.fromstring(element_xml))

                        if section_properties is not None:
                            body.append(section_properties)

                    xml_apply_appendix_section_settings(
                        root,
                        empty_footer_rid,
                        appendix_footer_rid,
                    )
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                elif item.filename == "word/settings.xml":
                    data = xml_enable_update_fields(data)
                elif re.fullmatch(r"word/footer\d+\.xml", item.filename):
                    if empty_footer_rid and footer_targets.get(item.filename) == empty_footer_rid:
                        data = xml_generated_empty_footer()
                    else:
                        data = xml_generated_appendix_footer(appendix_number)

                output_zip.writestr(item, data)

    output_stream.seek(0)
    return output_stream.getvalue()


def build_appendix_docx_fast(
    prepared_upload,
    work_items: list[str],
    project_name: str = "",
) -> tuple[bytes, int, list[str]]:
    index_data = prepared_upload.get("index") or {"toc_source": None, "sections": []}
    appendix_number = int(prepared_upload.get("position", 0)) + 1
    selected_sections = select_indexed_sections(
        index_data,
        work_items,
        appendix_number,
    )
    missing_work_items = [
        work_item
        for work_item in work_items
        if not any(
            same_work_item_name(
                section.get("requested_work_item", section["work_item"]),
                work_item,
            )
            for section in selected_sections
        )
    ]
    toc_work_items = [
        section.get("requested_work_item", section["work_item"])
        for section in selected_sections
    ]
    toc_display_work_items = [
        section.get("requested_work_item", section["work_item"])
        for section in selected_sections
    ]

    start_bookmark_names = [
        f"APPENDIX_{appendix_number}_{section_index + 1}_START"
        for section_index in range(len(selected_sections))
    ]
    end_bookmark_names = [
        f"APPENDIX_{appendix_number}_{section_index + 1}_END"
        for section_index in range(len(selected_sections))
    ]
    toc_page_range_bookmarks = list(zip(start_bookmark_names, end_bookmark_names))

    toc_elements = [
        xml_force_section_break_next_page(xml_set_all_text_font(element_xml, "標楷體"))
        for element_xml in xml_filter_toc_elements(
            index_data.get("toc_source"),
            toc_work_items,
            appendix_number,
            start_bookmark_names,
            toc_display_work_items,
        )
    ]

    body_elements = list(toc_elements)

    for section_index, section in enumerate(selected_sections):
        section_elements = []
        requested_work_item = section.get("requested_work_item", section["work_item"])
        for element_xml in selected_section_element_xmls(section, work_items):
            element_xml = xml_replace_work_item_text(
                element_xml,
                section["work_item"],
                requested_work_item,
            )
            if appendix_number in (2, 5):
                element_xml = xml_fill_project_name_in_record_table(
                    element_xml,
                    project_name,
                )
                element_xml = xml_fill_appendix_code_in_record_table(
                    element_xml,
                    appendix_number,
                    section_index + 1,
                )
            section_elements.append(element_xml)

        if section_elements:
            bookmark_base_id = appendix_number * 100000 + (section_index + 1) * 10
            section_elements[0] = xml_add_bookmark_to_element(
                section_elements[0],
                start_bookmark_names[section_index],
                bookmark_base_id + 1,
            )
            section_elements[-1] = xml_add_bookmark_to_element_end(
                section_elements[-1],
                end_bookmark_names[section_index],
                bookmark_base_id + 2,
            )

        section_elements = [
            xml_force_section_break_next_page(element_xml)
            for element_xml in section_elements
        ]

        if section_index == 0:
            if toc_elements and xml_needs_page_break_after_toc(toc_elements):
                body_elements.append(xml_section_break_next_page_paragraph())
        elif section_elements:
            section_elements[0] = xml_set_page_break_before(section_elements[0])

        body_elements.extend(section_elements)

    output_bytes = replace_document_body_with_xml(
        prepared_upload["docx_bytes"],
        body_elements,
        appendix_number,
    )
    output_bytes = refresh_docx_fields_with_word(
        output_bytes,
        appendix_number,
        toc_page_range_bookmarks,
    )
    return output_bytes, len(selected_sections), missing_work_items


st.set_page_config(page_title="監造附錄六合一產出系統", layout="wide")
require_app_password()
render_page_header()

upload_column, work_item_column = st.columns([1, 2.2], gap="large")

with upload_column:
    st.subheader("Word 檔案匯入")
    uploaded_files = st.file_uploader(
        "一次拖曳 6 個 Word 檔案",
        type=["doc", "docx"],
        accept_multiple_files=True,
        key="appendix_files",
    )
    uploaded_files = uploaded_files or []

    if uploaded_files:
        if len(uploaded_files) != len(APPENDIX_NAMES):
            st.warning("請一次上傳 6 個 Word 檔案，依序對應附錄一到附錄六。")

        with st.container(border=True):
            st.markdown("#### 檔案對應")
            for payload in uploaded_file_payloads(uploaded_files):
                st.write(f"{APPENDIX_NAMES[payload['position']]}：{payload['name']}")

uploaded_payloads = uploaded_file_payloads(uploaded_files)
toc_source_position = toc_source_position_from_payloads(uploaded_payloads)
current_upload_signature = (
    PREPARE_CACHE_VERSION,
    uploaded_payloads_signature(uploaded_payloads),
)
prepare_cache = st.session_state.setdefault("prepare_cache", {})
index_cache = st.session_state.setdefault("index_cache", {})
output_cache = st.session_state.setdefault("output_cache", {})

if not uploaded_files:
    st.session_state["prepared_upload_signature"] = ()
    st.session_state["prepared_uploads"] = []
    st.session_state["toc_work_items"] = []
    st.session_state["toc_errors"] = []
    st.session_state.pop("outputs", None)
    st.session_state.pop("outputs_signature", None)
elif st.session_state.get("prepared_upload_signature") != current_upload_signature:
    st.session_state.pop("outputs", None)
    st.session_state.pop("outputs_signature", None)
    prepared_by_position = {}
    toc_work_items_for_uploads = []
    toc_errors_for_uploads = []
    missing_payloads = []

    for payload in uploaded_payloads:
        cache_key = (PREPARE_CACHE_VERSION, payload["name"], payload["digest"])
        cached_upload = prepare_cache.get(cache_key)
        if cached_upload is None:
            missing_payloads.append(payload)
        else:
            prepared_by_position[payload["position"]] = dict(cached_upload)

    if missing_payloads:
     with st.spinner("正在快速讀取 Word 檔案..."):
        if len(missing_payloads) == 1:
            payload = missing_payloads[0]
            prepared_results = [
                prepare_word_payload(
                    payload,
                    read_toc=(payload["position"] == toc_source_position),
                )
            ]
        else:
            with ThreadPoolExecutor(
                max_workers=worker_count(len(missing_payloads), MAX_PREPARE_WORKERS)
            ) as executor:
                prepared_results = list(
                    executor.map(
                        lambda payload: prepare_word_payload(
                            payload,
                            read_toc=(payload["position"] == toc_source_position),
                        ),
                        missing_payloads,
                    )
                )

        for prepared_upload in prepared_results:
            cache_key = (
                PREPARE_CACHE_VERSION,
                prepared_upload["name"],
                prepared_upload["digest"],
            )
            prepare_cache[cache_key] = dict(prepared_upload)
            prepared_by_position[prepared_upload["position"]] = prepared_upload

    prepared_uploads = [
        prepared_by_position[index]
        for index in sorted(prepared_by_position)
    ]

    toc_source_upload = prepared_by_position.get(toc_source_position)
    if toc_source_upload:
        toc_work_items_for_uploads = list(toc_source_upload["work_items"])

    for prepared_upload in prepared_uploads:
        if prepared_upload["docx_bytes"] is None and prepared_upload["error"]:
            toc_errors_for_uploads.append(
                f"{prepared_upload['name']}：{prepared_upload['error']}"
            )

    st.session_state["prepared_upload_signature"] = current_upload_signature
    st.session_state["prepared_uploads"] = prepared_uploads
    st.session_state["toc_work_items"] = list(dict.fromkeys(toc_work_items_for_uploads))
    st.session_state["toc_errors"] = toc_errors_for_uploads

prepared_uploads = st.session_state.get("prepared_uploads", [])
toc_work_items = st.session_state.get("toc_work_items", [])
toc_errors = st.session_state.get("toc_errors", [])

with work_item_column:
    project_name = st.text_input(
        "工程名稱",
        placeholder="請輸入工程名稱",
        key="project_name",
    )
    st.subheader("工項選擇表單")

    selectable_work_items = [
        work_item
        for work_item in toc_work_items
        if not is_default_work_item_option(work_item)
    ]

    selectable_work_item_labels = {
        work_item: f"{index:02d}. {work_item}"
        for index, work_item in enumerate(selectable_work_items, start=1)
    }

    if toc_errors:
        st.warning("部分檔案無法讀取目錄，請確認檔案格式或是否可正常開啟。")
        with st.expander("查看讀取失敗明細"):
            for error in toc_errors:
                st.write(error)

    if "work_item_select_count" not in st.session_state:
        st.session_state["work_item_select_count"] = 6

    if selectable_work_items:
        st.caption(
            f"已從 Word 目錄抓到 {len(selectable_work_items)} 個可選工項"
        )

        if st.button("新增工項欄位"):
            st.session_state["work_item_select_count"] += 1

        selected_work_items = []
        for row_start in range(0, st.session_state["work_item_select_count"], 2):
            select_columns = st.columns(2)
            for column_index in range(2):
                field_index = row_start + column_index
                if field_index >= st.session_state["work_item_select_count"]:
                    continue

                select_key = f"work_item_select_{field_index + 1}"
                if (
                    select_key in st.session_state
                    and st.session_state[select_key] not in selectable_work_items
                ):
                    st.session_state[select_key] = None

                with select_columns[column_index]:
                    selected_work_item = st.selectbox(
                        f"工項 {field_index + 1}",
                        options=selectable_work_items,
                        format_func=lambda item: selectable_work_item_labels.get(item, item),
                        index=None,
                        placeholder="請選擇工項",
                        filter_mode="contains",
                        key=select_key,
                    )

                if selected_work_item:
                    selected_work_items.append(selected_work_item)

        selected_work_items = list(dict.fromkeys(selected_work_items))

        if selected_work_items:
            for row_start in range(0, len(selected_work_items), 2):
                columns = st.columns(2)
                for column_index, work_item in enumerate(
                    selected_work_items[row_start : row_start + 2]
                ):
                    with columns[column_index].container(border=True):
                        st.markdown(f"#### {work_item}")
                        st.caption("已選取工項")
        else:
            st.info("請從目錄工項選單選擇要納入產出的工項。")
    else:
        selected_work_items = []
        if toc_work_items:
            st.info("請從目錄工項選單選擇要納入產出的工項。")
        else:
            st.info("請先匯入 Word 檔案，系統會自動抓取目錄內容產生工項選單。")

    output_work_items = list(dict.fromkeys(DEFAULT_WORK_ITEMS + selected_work_items))
    st.markdown("#### 本次產出工項")
    st.write("、".join(output_work_items))

current_output_signature = (
    PREPARE_CACHE_VERSION,
    current_upload_signature,
    tuple(output_work_items),
    project_name.strip(),
)
outputs_are_current = (
    st.session_state.get("outputs_signature") == current_output_signature
    and bool(st.session_state.get("outputs"))
)

_, action_column = st.columns([1, 2.2], gap="large")
with action_column:
    _, button_column = st.columns([4, 1])
    with button_column:
        complete_clicked = st.button("完成", type="primary", use_container_width=True)

if complete_clicked:
    if outputs_are_current:
        st.success("已沿用上次產出結果，可直接下載。")
    elif len(uploaded_files) != len(APPENDIX_NAMES):
        st.error("請先一次上傳 6 個 Word 檔案。")
    elif not output_work_items:
        st.error("請至少選擇一個要產出的工項。")
    elif any(upload["docx_bytes"] is None for upload in prepared_uploads):
        st.error("有檔案無法讀取或轉換，請確認 Word 檔案格式後重新上傳。")
        failed_uploads = [
            upload
            for upload in prepared_uploads
            if upload["docx_bytes"] is None
        ]
        with st.expander("查看失敗明細", expanded=True):
            for upload in failed_uploads:
                st.write(f"{upload['name']}：{upload['error']}")
    else:
        outputs = [None] * len(APPENDIX_NAMES)
        empty_outputs = []
        partial_outputs = []

        with st.spinner("處理中..."):
            project_name_for_output = project_name.strip()
            output_file_prefix = sanitize_filename_part(project_name_for_output)
            output_work_items_tuple = tuple(output_work_items)
            indexed_uploads = [None] * len(prepared_uploads[: len(APPENDIX_NAMES)])
            missing_index_jobs = []

            for index, prepared_upload in enumerate(prepared_uploads[: len(APPENDIX_NAMES)]):
                boundary_work_items = appendix_boundary_work_items(
                    prepared_upload,
                    output_work_items_tuple,
                )
                index_key = (
                    PREPARE_CACHE_VERSION,
                    prepared_upload["docx_digest"],
                    boundary_work_items,
                )
                cached_index = index_cache.get(index_key)
                if cached_index is not None:
                    indexed_upload = dict(prepared_upload)
                    indexed_upload["index"] = cached_index
                    indexed_uploads[index] = indexed_upload
                else:
                    missing_index_jobs.append(
                        (index, prepared_upload, index_key, boundary_work_items)
                    )

            if missing_index_jobs:
                if len(missing_index_jobs) == 1:
                    indexed_results = [
                        (
                            missing_index_jobs[0][0],
                            build_prepared_index(
                                missing_index_jobs[0][1],
                                missing_index_jobs[0][3],
                            ),
                            missing_index_jobs[0][2],
                        )
                    ]
                else:
                    with ThreadPoolExecutor(
                        max_workers=worker_count(len(missing_index_jobs), MAX_PREPARE_WORKERS)
                    ) as executor:
                        built_indexes = list(
                            executor.map(
                                lambda job: build_prepared_index(job[1], job[3]),
                                missing_index_jobs,
                            )
                        )
                    indexed_results = [
                        (job[0], indexed_upload, job[2])
                        for indexed_upload, job in zip(built_indexes, missing_index_jobs)
                    ]

                for index, indexed_upload, index_key in indexed_results:
                    if indexed_upload.get("index") is not None:
                        index_cache[index_key] = indexed_upload["index"]
                    indexed_uploads[index] = indexed_upload

            st.session_state["prepared_uploads"] = indexed_uploads
            index_failures = [
                upload
                for upload in indexed_uploads
                if upload is None or upload.get("index") is None
            ]
            if index_failures:
                st.error("有檔案建立快速索引失敗，請確認 Word 檔案格式後重新上傳。")
                with st.expander("查看索引失敗明細", expanded=True):
                    for upload in index_failures:
                        if upload is None:
                            st.write("未知檔案：建立快速索引失敗")
                        else:
                            st.write(f"{upload['name']}：{upload.get('error')}")
                st.stop()

            missing_output_jobs = []

            for index, prepared_upload in enumerate(indexed_uploads):
                output_key = (
                    PREPARE_CACHE_VERSION,
                    prepared_upload["docx_digest"],
                    output_work_items_tuple,
                    project_name_for_output,
                )
                cached_output = output_cache.get(output_key)
                if cached_output is None:
                    missing_output_jobs.append((index, prepared_upload, output_key))
                    continue

                if len(cached_output) == 2:
                    output_bytes, selected_element_count = cached_output
                    missing_work_items = []
                else:
                    output_bytes, selected_element_count, missing_work_items = cached_output
                if selected_element_count == 0:
                    empty_outputs.append(APPENDIX_NAMES[index])
                elif selected_element_count < len(output_work_items_tuple):
                    missing_text = (
                        f"；缺：{', '.join(missing_work_items)}"
                        if missing_work_items
                        else ""
                    )
                    partial_outputs.append(
                        f"{APPENDIX_NAMES[index]}：抓到 {selected_element_count} / 應有 {len(output_work_items_tuple)}{missing_text}"
                    )
                outputs[index] = {
                    "label": APPENDIX_NAMES[index],
                    "file_name": f"{output_file_prefix}_{APPENDIX_NAMES[index]}.docx",
                    "bytes": output_bytes,
                }

            if missing_output_jobs:
                if len(missing_output_jobs) == 1:
                    output_results = [
                        (
                            build_single_output(
                                missing_output_jobs[0][0],
                                missing_output_jobs[0][1],
                                output_work_items_tuple,
                                project_name_for_output,
                            ),
                            missing_output_jobs[0][2],
                        )
                    ]
                else:
                    with ThreadPoolExecutor(
                        max_workers=worker_count(len(missing_output_jobs), MAX_OUTPUT_WORKERS)
                    ) as executor:
                        built_outputs = list(
                            executor.map(
                                lambda job: build_single_output(
                                    job[0],
                                    job[1],
                                    output_work_items_tuple,
                                    project_name_for_output,
                                ),
                                missing_output_jobs,
                            )
                        )
                    output_results = [
                        (built_output, job[2])
                        for built_output, job in zip(built_outputs, missing_output_jobs)
                    ]

                for (
                    index,
                    output_bytes,
                    selected_element_count,
                    missing_work_items,
                ), output_key in output_results:
                    output_cache[output_key] = (
                        output_bytes,
                        selected_element_count,
                        missing_work_items,
                    )
                    if selected_element_count == 0:
                        empty_outputs.append(APPENDIX_NAMES[index])
                    elif selected_element_count < len(output_work_items_tuple):
                        missing_text = (
                            f"；缺：{', '.join(missing_work_items)}"
                            if missing_work_items
                            else ""
                        )
                        partial_outputs.append(
                            f"{APPENDIX_NAMES[index]}：抓到 {selected_element_count} / 應有 {len(output_work_items_tuple)}{missing_text}"
                        )
                    outputs[index] = {
                        "label": APPENDIX_NAMES[index],
                        "file_name": f"{output_file_prefix}_{APPENDIX_NAMES[index]}.docx",
                        "bytes": output_bytes,
                    }

        st.session_state["outputs"] = [output for output in outputs if output is not None]
        st.session_state["outputs_signature"] = current_output_signature
        if empty_outputs:
            st.warning(f"{'、'.join(empty_outputs)} 沒有抓到指定工項內容。")
        if partial_outputs:
            st.warning("部分附錄有少抓工項：" + "；".join(partial_outputs))
        if not empty_outputs and not partial_outputs:
            st.success("處理完成，請分別下載 6 個 Word 檔案。")

if "outputs" in st.session_state:
    outputs_match_current = (
        st.session_state.get("outputs_signature") == current_output_signature
    )
    output_title = "輸出下載" if outputs_match_current else "上次輸出下載"
    title_column, clear_column = st.columns([4, 1])
    with title_column:
        st.subheader(output_title)
    with clear_column:
        if st.button("清除輸出", key="clear_outputs", use_container_width=True):
            st.session_state.pop("outputs", None)
            st.session_state.pop("outputs_signature", None)
            st.rerun()
    download_rows = [st.columns(3), st.columns(3)]
    for index, output in enumerate(st.session_state["outputs"]):
        column = download_rows[index // 3][index % 3]
        with column.container(border=True):
            st.download_button(
                f"下載{output['label']}",
                data=output["bytes"],
                file_name=output["file_name"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"download_{index + 1}",
                on_click="ignore",
            )

from pathlib import Path


source_path = Path(__file__).with_name("監造附錄六合一產出系統.py")
exec(compile(source_path.read_text(encoding="utf-8"), str(source_path), "exec"))

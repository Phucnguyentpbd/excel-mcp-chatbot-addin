import contextlib
import io
import json
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from openpyxl import Workbook, load_workbook


def execute_python_code(
    code: str,
    resolve_path: Callable[[str], str],
    data: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Execute Python code with helpers aligned to the server's file-path policy."""
    if not code or not code.strip():
        raise ValueError("Python code is required")

    stdout_buffer = io.StringIO()
    locals_dict: Dict[str, Any] = {}
    output: Any = None

    def save_workbook(workbook: Workbook, filepath: str) -> str:
        full_path = Path(resolve_path(filepath))
        full_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(full_path)
        return str(full_path)

    def read_json(filepath: str) -> Any:
        full_path = Path(resolve_path(filepath))
        with full_path.open("r", encoding="utf-8") as f:
            return json.load(f)

    def write_json(filepath: str, payload: Any, indent: int = 2) -> str:
        full_path = Path(resolve_path(filepath))
        full_path.parent.mkdir(parents=True, exist_ok=True)
        with full_path.open("w", encoding="utf-8") as f:
            json.dump(payload, f, indent=indent, ensure_ascii=False, default=str)
        return str(full_path)

    globals_dict: Dict[str, Any] = {
        "__builtins__": __builtins__,
        "Workbook": Workbook,
        "Path": Path,
        "data": data,
        "json": json,
        "load_workbook": load_workbook,
        "output": output,
        "print": print,
        "read_json": read_json,
        "resolve_path": resolve_path,
        "save_workbook": save_workbook,
        "write_json": write_json,
    }

    with contextlib.redirect_stdout(stdout_buffer):
        exec(code, globals_dict, locals_dict)

    output = locals_dict.get("output", globals_dict.get("output"))
    stdout_text = stdout_buffer.getvalue()

    return {
        "output": output,
        "stdout": stdout_text,
    }

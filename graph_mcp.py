import ast
import os
import sys
import tempfile
import subprocess
import pandas as pd
from pathlib import Path
import logging

logger = logging.getLogger("SCReportGraph")

# ============================================================
# SECTION 1 — SECURITY POLICY
# ============================================================

# List of allowed root modules
ALLOWED_MODULES = {
    # Data
    'pandas', 'numpy', 'json', 'csv', 'math', 'statistics',
    'collections', 'itertools', 'functools', 'operator',
    # String & Text
    'string', 're', 'textwrap', 'unicodedata',
    # Date & Time
    'datetime', 'calendar', 'time',
    # Visualization
    'matplotlib', 'matplotlib.pyplot', 'matplotlib.ticker',
    'matplotlib.colors', 'matplotlib.patches', 'matplotlib.cm',
    'seaborn',
    # Formatting / printing
    'pprint', 'decimal', 'fractions',
    # Type hints (read-only)
    'typing',
}

# List of blacklisted names/builtins to prevent malicious activities
BANNED_NAMES = {
    'open', 'eval', 'exec', '__import__', 'compile', 'globals', 'locals',
    'input', 'getattr', 'setattr', 'delattr', 'hasattr', 'os', 'sys', 'subprocess',
    'socket', 'shutil', 'builtins', 'sh', 'pty', 'platform', 'importlib',
    '__builtins__', '__class__', '__module__', '__dict__', '__getattribute__'
}


def verify_code_safety(code: str) -> None:
    """
    Parses Python code using AST and validates it against security policies:
    - Only whitelist libraries can be imported.
    - No dangerous builtins or names can be referenced.
    - No magic attribute/property access (__) is allowed.
    """
    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        raise ValueError(f"โค้ดไวยากรณ์ไม่ถูกต้อง (Syntax Error): {e}")

    for node in ast.walk(tree):
        # 1. Check imports
        if isinstance(node, ast.Import):
            for alias in node.names:
                root_module = alias.name.split('.')[0]
                if root_module not in ALLOWED_MODULES:
                    raise PermissionError(f"ไม่อนุญาตให้ใช้โมดูล '{alias.name}'")

        elif isinstance(node, ast.ImportFrom):
            if not node.module:
                raise PermissionError("ไม่อนุญาตให้ใช้ relative import")
            root_module = node.module.split('.')[0]
            if root_module not in ALLOWED_MODULES:
                raise PermissionError(f"ไม่อนุญาตให้ใช้โมดูล '{node.module}'")

        # 2. Check names
        elif isinstance(node, ast.Name):
            if node.id in BANNED_NAMES:
                raise PermissionError(f"ไม่อนุญาตให้ใช้งาน '{node.id}'")

        # 3. Check attributes (prevent __class__, etc.)
        elif isinstance(node, ast.Attribute):
            if node.attr.startswith('__') or node.attr in BANNED_NAMES:
                raise PermissionError(f"ไม่อนุญาตให้เข้าถึง attribute '{node.attr}'")


# ============================================================
# SECTION 2 — GENERAL CODE SANDBOX (Replaces run_visualization_code)
# ============================================================

def run_code(
    code_str: str,
    df: pd.DataFrame | None = None,
    output_png_path: str | None = None,
    column_labels: dict | None = None,   # ← เพิ่ม
    timeout: float = 15.0
) -> dict:
    """
    รันโค้ด Python ทั่วไปในระบบ Sandbox ที่แยก subprocess

    พารามิเตอร์:
        code_str       : โค้ด Python ที่ต้องการรัน (string)
        df             : pandas DataFrame (optional) — ถ้าส่งมา จะสามารถใช้ชื่อ `df` หรือ `data` ในโค้ดได้เลย
        output_png_path: พาธสำหรับเซฟรูปภาพ (optional) — ถ้าไม่ส่งมา จะไม่พยายามเซฟรูป
        column_labels  : dict mapping original column name → Thai description (optional)
                         inject เป็นตัวแปร `column_labels` ใน sandbox ให้ใช้ rename axis label
        timeout        : เวลาสูงสุดในการรัน (วินาที, default 15)

    ผลลัพธ์ที่ return:
        {
            "success": bool,
            "output":  str,
            "message": str,
            "has_image": bool
        }
    """
    # Step 1: Security check
    try:
        verify_code_safety(code_str)
    except Exception as e:
        return {
            "success": False,
            "output": "",
            "message": f"Security Validation Error: {str(e)}",
            "has_image": False,
        }

    # Step 2: Write temporary files
    temp_dir = tempfile.mkdtemp(prefix="sc_sandbox_")
    pkl_path = Path(temp_dir) / "data.pkl"
    script_path = Path(temp_dir) / "runner.py"
    has_png = output_png_path is not None

    try:
        df_load_lines = ""
        if df is not None:
            df.to_pickle(pkl_path)
            posix_pkl_path = pkl_path.as_posix()
            df_load_lines = f"""
import pandas as _pd
df = _pd.read_pickle("{posix_pkl_path}")
data = df
"""

        # ← inject column_labels ถ้ามี
        column_labels_line = ""
        if column_labels:
            column_labels_line = f"column_labels = {repr(column_labels)}\n"
        else:
            column_labels_line = "column_labels = {}\n"

        posix_png = Path(output_png_path).as_posix() if has_png else ""
        save_plot_lines = ""
        if has_png:
            save_plot_lines = f"""
_fig = plt.gcf()
if _fig.get_axes():
    plt.savefig("{posix_png}", bbox_inches='tight')
plt.close('all')
"""

        runner_content = f"""# -*- coding: utf-8 -*-
import sys
import io

# ---- Pre-load libraries available to user code ----
import pandas as pd
import numpy as np
import math
import json
import re
import statistics
import datetime
import collections
import itertools
import functools
import textwrap
import string
import pprint

import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm._load_fontmanager(try_read_cache=False)
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.cm as cm
import seaborn as sns

sns.set_theme(style="ticks")
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = [
    'Garuda', 'Loma', 'Norasi', 'Kinnari', 'Umpush',
    'TlwgMono', 'DejaVu Sans', 'sans-serif'
]
matplotlib.rcParams['axes.unicode_minus'] = False

{df_load_lines}
{column_labels_line}
# ---- Run user code ----
{code_str}

{save_plot_lines}
"""
        with open(script_path, "w", encoding="utf-8") as f:
            f.write(runner_content)

        python_exe = sys.executable
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONPATH"] = ":".join(sys.path)
        for key in ["CLEANUP_SECRET", "GOOGLE_APPLICATION_CREDENTIALS"]:
            env.pop(key, None)

        result = subprocess.run(
            [python_exe, str(script_path)],
            capture_output=True,
            text=True,
            timeout=timeout,
            env=env,
            encoding="utf-8",
            errors="replace",
        )

        stdout_output = result.stdout or ""

        if result.returncode == 0:
            image_created = has_png and Path(output_png_path).exists()
            return {
                "success": True,
                "output": stdout_output,
                "message": "Code executed successfully",
                "has_image": image_created,
            }

        stderr = result.stderr or ""
        error_lines = stderr.strip().split("\n")
        err_msg = error_lines[-1] if error_lines else "Unknown error in code execution."
        logger.error(f"Error in runner execution:\n{stderr}")
        return {
            "success": False,
            "output": stdout_output,
            "message": f"Execution Error: {err_msg}\n{stderr}",
            "has_image": False,
        }

    except subprocess.TimeoutExpired:
        logger.error("Sandbox execution timed out.")
        return {
            "success": False,
            "output": "",
            "message": f"Execution Error: Code timed out after {timeout} seconds.",
            "has_image": False,
        }
    except Exception as e:
        logger.error(f"Unexpected error in sandbox execution: {e}", exc_info=True)
        return {
            "success": False,
            "output": "",
            "message": f"Execution Error: {str(e)}",
            "has_image": False,
        }
    finally:
        try:
            if pkl_path.exists():
                pkl_path.unlink()
            if script_path.exists():
                script_path.unlink()
            os.rmdir(temp_dir)
        except Exception as e:
            logger.warning(f"Failed to clean up temp directory {temp_dir}: {e}")


# ============================================================
# SECTION 3 — BACKWARD COMPAT: run_visualization_code alias
# ============================================================

def run_visualization_code(
    df: pd.DataFrame,
    code_str: str,
    output_png_path: str,
    column_labels: dict | None = None,   # ← เพิ่ม
    timeout: float = 10.0,
) -> dict:
    """
    Backward-compatible wrapper สำหรับโค้ดเก่าที่เรียก run_visualization_code อยู่
    ภายในเรียก run_code ตัวใหม่
    """
    res = run_code(
        code_str=code_str,
        df=df,
        output_png_path=output_png_path,
        column_labels=column_labels,     # ← pass through
        timeout=timeout,
    )
    return {
        "success": res["success"],
        "message": res["message"],
    }
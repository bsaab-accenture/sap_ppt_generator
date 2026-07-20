"""
function_app.py
───────────────
Azure Functions v2 entry point — Agenda DOCX generator.

All trigger definitions live in docx_core/triggers.py (Blueprint).
To add this generator to an existing Function App, register the blueprint:

    from docx_core.triggers import bp as docx_bp
    app.register_functions(docx_bp)
"""

from __future__ import annotations

import logging

import azure.functions as func

from docx_core.triggers import bp as docx_bp

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(name)s]  %(message)s",
    datefmt="%H:%M:%S",
)

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)
app.register_functions(docx_bp)

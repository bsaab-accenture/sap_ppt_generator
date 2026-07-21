"""
ppt_core/generator.py
─────────────────────
Agenda PPT generation logic — PLACEHOLDER.

Update this file once the agenda PPTX template is uploaded to
local_storage/templates-ppt/AgendaTemplate.pptx and the input JSON
schema is finalised.

Public API (keep this signature stable):
    generate_pptx(template_bytes, content_json, cfg) -> bytes
"""

from __future__ import annotations

import logging
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

log = logging.getLogger(__name__)


@dataclass
class GeneratorConfig:
    """Tuneable parameters for one generation run."""
    scripts_dir: Path
    workdir: Path
    max_bullets: int = 4


def generate_pptx(
    template_bytes: bytes,
    content_json: dict,
    cfg: GeneratorConfig,
) -> bytes:
    """
    Generate a populated Agenda PPTX.

    Parameters
    ----------
    template_bytes : raw bytes of the .pptx template file
    content_json   : parsed JSON payload
    cfg            : GeneratorConfig

    Returns
    -------
    bytes  – raw .pptx file content

    TODO: implement agenda-specific slide injection once the template
          and JSON schema are confirmed.
    """
    raise NotImplementedError(
        "Agenda PPT generator is not implemented yet. "
        "Upload AgendaTemplate.pptx and finalise the JSON schema first."
    )

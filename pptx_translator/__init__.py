"""pptx_translator - OOXML-preserving PowerPoint translation patcher.

Two-step, deterministic (no-AI) workflow:

    1. extract  -> JSON of every translatable paragraph
    2. patch    -> new .pptx with your translations, formatting untouched
"""

__version__ = "1.0.0"

from .extract import build_units, extract_to_json  # noqa: F401
from .patch import PatchError, patch_pptx  # noqa: F401

__all__ = ["extract_to_json", "build_units", "patch_pptx", "PatchError", "__version__"]

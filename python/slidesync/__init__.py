"""SlideSync, refresh PowerPoint decks from Excel workbooks."""

from .config import Config, load_config
from .engine import run

__all__ = ["run", "Config", "load_config"]
__version__ = "0.1.0"

from abc import ABC, abstractmethod
from typing import Optional

from src.models.table_model import Sheet

class BaseParser(ABC):
    """Abstract base class for file parsers."""

    @abstractmethod
    def parse(self, file_path: str, sheet_name: Optional[str] = None) -> Sheet:
        """Parses the given file and returns a Sheet object."""
        pass

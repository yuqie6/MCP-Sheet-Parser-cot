from abc import ABC, abstractmethod
from src.models.table_model import Sheet

class BaseParser(ABC):
    @abstractmethod
    def parse(self, file_path: str) -> Sheet:
        """Parses the given file and returns a Sheet object."""
        pass
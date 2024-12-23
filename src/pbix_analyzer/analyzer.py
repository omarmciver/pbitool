"""
PBIX Analyzer main module for analyzing Power BI files.
"""
import zipfile
from pathlib import Path
from typing import List, Dict


class PBIXAnalyzer:
    """
    A class to analyze Power BI (.pbix) files by treating them as ZIP archives.
    """

    def __init__(self, pbix_path: str):
        """
        Initialize the analyzer with a path to a .pbix file.

        Args:
            pbix_path (str): Path to the .pbix file to analyze

        Raises:
            FileNotFoundError: If the specified file doesn't exist
            ValueError: If the file doesn't have a .pbix extension
        """
        self.pbix_path = Path(pbix_path)
        if not self.pbix_path.exists():
            raise FileNotFoundError(f"PBIX file not found: {pbix_path}")

        if self.pbix_path.suffix.lower() != '.pbix':
            raise ValueError(
                f"File must have .pbix extension, got: {self.pbix_path.suffix}"
            )

    def list_contents(self) -> List[Dict[str, str]]:
        """
        List all files contained within the PBIX file.

        Returns:
            List[Dict[str, str]]: List of dictionaries containing file
            information with 'name' and 'size' keys
        Raises:
            ValueError: If the file cannot be opened as a ZIP archive
        """
        contents = []
        try:
            with zipfile.ZipFile(self.pbix_path, 'r') as pbix:
                for file_info in pbix.filelist:
                    contents.append({
                        'name': file_info.filename,
                        'size': f"{file_info.file_size:,} bytes"
                    })
        except zipfile.BadZipFile:
            raise ValueError(
                f"Failed to open {self.pbix_path} as a ZIP file. "
                f"File may be corrupted."
            )

        return contents

    def print_contents(self) -> None:
        """
        Print the contents of the PBIX file in a formatted way.
        """
        contents = self.list_contents()

        print(f"\nContents of {self.pbix_path.name}:")
        print("-" * 80)

        for item in contents:
            print(f"{item['name']:<60} {item['size']:>15}")

        print(f"\nTotal files: {len(contents)}")
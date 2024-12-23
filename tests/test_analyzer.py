"""
Tests for the PBIX Analyzer module.
"""
from pbix_analyzer.analyzer import PBIXAnalyzer
import pytest
from pathlib import Path

# Test data directory
TEST_DATA_DIR = Path(__file__).parent / "data"


def test_analyzer_initialization_with_invalid_extension():
    """Test that analyzer raises error with wrong file extension"""
    with pytest.raises(ValueError) as exc_info:
        PBIXAnalyzer("test.txt")
    assert "must have .pbix extension" in str(exc_info.value)


def test_analyzer_initialization_with_nonexistent_file():
    """Test that analyzer raises error with nonexistent file"""
    with pytest.raises(FileNotFoundError) as exc_info:
        PBIXAnalyzer("nonexistent.pbix")
    assert "PBIX file not found" in str(exc_info.value)

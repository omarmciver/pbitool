# PBIX Analyzer

A Python tool for analyzing Power BI (PBIX) files by treating them as ZIP archives.

## Installation

```bash
pip install -e .
```

## Usage

```python
from pbix_analyzer.analyzer import PBIXAnalyzer

analyzer = PBIXAnalyzer("path/to/your/file.pbix")
analyzer.print_contents()
```

## Development

1. Clone the repository
2. Create a virtual environment: `python -m venv venv`
3. Activate the virtual environment
4. Install development dependencies: `pip install -r requirements.txt`
5. Install the package in editable mode: `pip install -e .`

## Testing

Run tests with pytest:
```bash
pytest
```
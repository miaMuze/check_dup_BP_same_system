# BP Duplicate Checker

A Windows desktop application for detecting potential duplicate Business Partners using fuzzy name matching.

## Installation

### Prerequisites
- Python 3.8 or higher ([Download Python](https://www.python.org/downloads/))
- Windows 10/11

### Build Steps

1. Clone or download this repository
2. Double-click `INSTALL.bat`
3. Wait for the build to complete
4. The executable will be created at `dist\BP_Duplicate_Checker.exe`

## Usage

1. Run `BP_Duplicate_Checker.exe`
2. Click **Browse** to upload your Excel file (must contain columns: `BP_Number`, `Name1`, `Name2`)
3. Configure **Ignore Words** (e.g., Mr, Mrs, Ltd, Company) - these words will be excluded from comparison
4. Set **Minimum Score** threshold (default: 50%)
5. Click **Run Matching**
6. Review results and **Export to Excel** if needed

## How Scoring Works

The application uses a weighted combination of fuzzy matching algorithms:

| Algorithm | Weight | Purpose |
|-----------|--------|---------|
| Token Sort Ratio | 40% | Handles words in different order (e.g., "John Smith" vs "Smith John") |
| Token Set Ratio | 40% | Handles partial matches and extra words |
| Simple Ratio | 20% | Character-level similarity (Levenshtein distance) |

### Text Normalization

Before comparison, all names are normalized:
1. Convert to lowercase
2. Remove punctuation
3. Remove ignore words (configurable)
4. Trim extra whitespace

### Confidence Levels

| Level | Score | Description |
|-------|-------|-------------|
| High | 80-100% | Very likely duplicates - review first |
| Medium | 60-79% | Possible duplicates - needs review |
| Low | 50-59% | Less likely duplicates - lower priority |

## Privacy & Security Disclaimer

**This application runs entirely on your local machine.**

- No data is sent to external servers
- No internet connection is required
- All BP data remains on your computer
- The application does not collect, store, or transmit any information
- Safe for use with sensitive business data

## License

This project is provided as-is for internal use.

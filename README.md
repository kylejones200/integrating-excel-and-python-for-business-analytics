# Integrating Excel and Python for Business Analytics

This project demonstrates integrating Excel and Python for business analytics workflows.

## Business context

Python fixes that. When you combine Excel's accessibility with Python's power, you get the best of both worlds. You can take spreadsheets your colleagues already use, run complex models in Python, and put the results back in Excel without any manual steps. This makes your work more scalable and accurate --- and still accessible to teams that live in Excel.

This post shows how to move data between Excel and Python, automate report generation, and build hybrid analytics workflows that keep everyone productive. You'll learn how to read Excel files, write to them, format them, and even control Excel itself. We'll close with a real-world case study that combines these skills in a financial KPI dashboard.

Python can read Excel files easily using the `pandas` library, which wraps the `openpyxl` and `xlrd` engines under the hood. The simplest version reads the first sheet of an Excel file into a DataFrame.

## Article

Medium article: [Integrating Excel and Python for Business Analytics](https://medium.com/@kylejones_47003/integrating-excel-and-python-for-business-analytics-53281e2985e2)

## Project Structure

```
.
├── README.md           # This file
├── main.py            # Main entry point
├── config.yaml        # Configuration file
├── requirements.txt   # Python dependencies
├── src/               # Core functions
│   ├── core.py        # Excel-Python integration functions
│   └── plotting.py    # Tufte-style plotting utilities
├── tests/             # Unit tests
├── data/              # Data files
└── images/            # Generated plots and figures
```

## Configuration

Edit `config.yaml` to customize:
- Excel file path and sheet name
- Synthetic data generation
- Output settings

## Excel Integration

Features:
- Read Excel: Load data from .xlsx files
- Write Excel: Export processed data
- Multi-sheet: Support for multiple sheets
- Data Analysis: Statistical analysis of Excel data

## Caveats

- By default, generates synthetic data for demonstration.
- Requires openpyxl for Excel file handling.
- Excel file format must be .xlsx.

## Disclaimer

Educational/demo code only. Not financial, safety, or engineering advice. Use at your own risk. Verify results independently before any production or operational use.

## License

MIT — see [LICENSE](LICENSE).
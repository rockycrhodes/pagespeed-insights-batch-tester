# PageSpeed Insights Batch Tester

Runs Google PageSpeed Insights against a list of URLs and generates a formatted Core Web Vitals Excel report.

## Usage

1. Run the batch tester against a list of URLs:

```bash
python psi_batch.py --key YOUR_API_KEY --urls urls.txt --runs 5 --strategy both --output psi_results.json
```

2. Generate the formatted spreadsheet:

```bash
python generate_xlsx.py --input psi_results.json --output CWV_Results.xlsx --label "Pre-launch"
```

## Prerequisites

```bash
pip install openpyxl --break-system-packages
```

`psi_batch.py` uses only the Python standard library.

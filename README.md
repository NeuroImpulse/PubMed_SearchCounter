# PubMed Query Counter 🧠📊

This script takes a list of search queries from an Excel file and uses the PubMed API to count how many articles exist for each one. It adds the results to a new Excel file.

## Features

- Multithreaded PubMed API requests (faster!)
- Reads input from Excel
- Outputs Excel with article counts
- Simple and clean code

## Usage

1. Put your Excel file with a column named `PubMed_Search_Query` into this folder.
2. Run the script:

```bash
python pubmed_query_counter.py

# Excel API Data Extraction Pipeline

A Python-based API data extraction pipeline that reads keywords from an Excel file, searches the MediaWiki API, extracts relevant information, and generates a structured and formatted Excel report.

## Project Overview

This project automates the process of collecting online information for a list of keywords stored in an Excel file.

For example:

- GATE
- IAS
- KAAS
- UPSC

The pipeline reads each keyword, searches Wikipedia through the MediaWiki API, retrieves relevant page information, and stores the results in an Excel file.

## Project Workflow

Excel Input
    ↓
Read Keywords using Pandas
    ↓
Clean and Validate Data
    ↓
MediaWiki API Request
    ↓
Search Relevant Pages
    ↓
Extract Page Information
    ↓
Error Handling & Logging
    ↓
Format Excel Output
    ↓
Output Excel Report

## Features

- Read keywords from Excel
- Remove empty values
- Remove duplicate keywords
- Search the MediaWiki API
- Retrieve multiple search results
- Extract page title
- Extract page description
- Extract detailed page information
- Store source URLs
- Record collection date and status
- Handle API errors and timeouts
- Maintain execution logs
- Generate formatted Excel reports
- Freeze Excel headers
- Add Excel filters
- Create an Excel table automatically

## Technologies Used

- Python
- Pandas
- Requests
- MediaWiki API
- OpenPyXL
- JSON
- Python Logging
- Microsoft Excel
- Git
- GitHub

## Project Structure

```text
excel-api-data-extraction-pipeline/
│
├── input.xlsx
├── output.xlsx
├── main.py
├── .gitignore
└── README.md

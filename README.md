# MAC Address Migration Tool

A Python script that processes and formats MAC address upgrade data for network equipment installations. The tool merges building information with technician assignments and generates organized, color-coded tables for each tech.

## Overview

This tool takes raw data from two Excel worksheets containing:
1. Building codes with old and new MAC addresses
2. Installation schedules with tech assignments

It then combines this data and generates formatted tables grouped by technician and building, making it easy to distribute work assignments.

## Features

- Filters data to include only WWT FS partner installations
- Extracts relevant information (last 4 digits of MAC addresses, bridge identifiers)
- Merges building and tech data based on building codes
- Groups installations by technician
- Creates color-coded, formatted tables with:
  - Red headers showing tech name, building, and installation date
  - Gray column headers
  - Light blue data rows
  - Clear separation between different buildings

## Requirements

- Python 3.9 or higher
- pandas
- openpyxl

## Installation

1. Clone or download this repository

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Prepare your Excel file (`file.xlsx`) with the following structure:
   - **Sheet 1 (Index 0)**: Engineering Prework data
     - Column A: Building Code
     - Column C: Old MAC Address
     - Column D: New MAC Address
   
   - **Sheet 2 (Index 1)**: Tech Information
     - Column B: Building Code
     - Column F: Install Date
     - Column M: Partner (must be "WWT FS")
     - Column S: Bridge Link (last character will be extracted)
     - Column X: Tech Name

2. Run the script:
```bash
python3 main.py
```

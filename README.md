# Python Quotes Web Scraper

## Overview

A Python web scraper that extracts quotes, author names, and author profile links from a quotes website.

## Features

- Scrapes multiple pages
- Extracts quotes and author names
- Extracts author profile links
- Handles request errors
- Removes duplicate quotes
- Exports data to CSV
- Exports formatted data to Excel
- Automatically adjusts Excel column widths

## Technologies

- Python
- Requests
- BeautifulSoup
- pandas
- openpyxl

## Output

The scraper generates:

- `scraped_quotes.csv` — scraped data in CSV format
- `scraped_quotes.xlsx` — formatted Excel version of the data

The data includes:

- Quote
- Author
- Author profile link

## How to Run

Install the required packages:

```bash
pip install -r requirements.txt
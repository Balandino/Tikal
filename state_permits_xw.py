"""Opens latest building permits"""

import sys
from datetime import datetime

import requests
import xlwings


def run():
    """Checks for last month's report"""
    current_month = datetime.now().month
    current_year = datetime.now().year

    # If January, then previous month will be last year
    if current_month == 1:
        current_year = current_year - 1

    while current_month > 0:
        # Ensure leading 0 for dates
        str_month = str(current_month) if current_month >= 10 else f"0{current_month}"
        url = f"https://www.census.gov/construction/bps/xls/statemonthly_{current_year}{str_month}.xls"

        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            print(f"Web site exists: {url}")
            workbook = xlwings.Book.caller()
            macro = workbook.macro("OpenLink")
            macro(url)

            sys.exit(0)
        else:
            print(f"Web site does not exist: {url}")

        current_month = current_month - 1


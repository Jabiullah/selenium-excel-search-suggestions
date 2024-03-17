# Selenium Excel Search Suggestions
Automated tool to scrape Google search suggestions using Selenium WebDriver and Python.
## Installation
Clone the repository:
git clone "https://github.com/Jabiullah/selenium-excel-search-suggestions.git"

## Usage
1. Ensure you have Google Chrome installed on your system.
2. Download the Chrome WebDriver from [here](https://chromedriver.chromium.org/downloads) and place it in your system's PATH.
3. Create an Excel file with search keys in column C starting from row 3.
4. The script will open Google,
   perform searches using the search keys,
   collect suggestions,
   and update the Excel file
   with the longest and shortest suggestions for each search key.

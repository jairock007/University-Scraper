# University Scraper

`UniversityScraper` is a Python script that scrapes course and scholarship information from the University of Oxford's official website. The collected data is then exported to a well-formatted Excel file.

## Features

- Scrapes a complete list of undergraduate courses from Oxford University.
- Extracts scholarship information along with eligibility and amounts.
- Cleans and formats the scraped data.
- Exports the results to an Excel file with multiple sheets for courses, scholarships, and a summary.

## Requirements

- Python 3.x
- Libraries: `requests`, `beautifulsoup4`, `pandas`, `openpyxl`

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/jairock007/universityscraper.git
   cd universityscraper
   ```

2. Install the required libraries:
   ```bash
   pip install requests beautifulsoup4 pandas openpyxl
   ```

## Usage

1. Make sure you have Python installed on your system.
2. Run the script:
   ```bash
   python university_scraper.py
   ```
3. The program will scrape data and save it to an Excel file named `oxford_data_YYYYMMDD_HHMMSS.xlsx` in the current directory.

## Logging

The script uses the `logging` module to provide detailed information about the scraping process. Logs will be output to the console.

## Data Structure

The exported Excel file consists of the following sheets:

- **Courses**: Contains a list of undergraduate courses.
- **Scholarships**: Lists scholarships with details about eligibility and amounts.
- **Summary**: Provides a summary of the total courses and scholarships found, along with the last updated timestamp.

## Contributing

Contributions are welcome! Please feel free to submit a pull request if you want to improve this project.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Author

[Jitendra Singh][https://github.com/jairock007]


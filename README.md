
# APK Review Scraper (Multi-APK)

This Python script is designed to automatically scrape user reviews from multiple Android applications listed in an Excel file. It uses the `google-play-scraper` library to fetch reviews from the Google Play Store and saves the results into a structured Excel file.

## Features
- Reads APK package names from an Excel file (`Updated_List_APK.xlsx`)
- Scrapes reviews for each APK listed
- Skips APKs that are inaccessible or return errors
- Saves all reviews into a single Excel file (`reviews_all_apk.xlsx`)
- Adds an extra column `apk_name` to identify the source application
- Formats review dates as `YYYY-MM-DD`
- Automatically splits output into multiple sheets if the number of rows exceeds 1,000,000
- Deduplicates reviews based on `reviewId` and `apk_name`

## Requirements
- Python 3.x
- Required libraries:
  ```bash
  pip install google-play-scraper pandas openpyxl xlsxwriter
  ```

## Input Format
The input Excel file (`Updated_List_APK.xlsx`) must contain the following columns:
- `Alamat APK  Android (com.xxx.xx)` — the package name of the Android app (required)
- `Nama Platform` — the name of the platform (optional but recommended)

Example:
| Nama Platform | Alamat APK  Android (com.xxx.xx) |
|---------------|----------------------------------|
| Danamas       | com.danamas.mergingapp           |
| Investree     | id.investree                     |

## Output
The script generates an Excel file named `reviews_all_apk.xlsx` containing all the scraped reviews. If the total number of reviews exceeds 1 million, the data is split across multiple sheets named `reviews_1`, `reviews_2`, etc.

Each review includes the following columns:
- `apk_name`
- `reviewId`
- `userName`
- `score`
- `text`
- `at` (review date)
- `replyText`
- `replyAt` (reply date)
- `thumbsUpCount`
- `version`

## How to Run
1. Place `main.py` and `Updated_List_APK.xlsx` in the same folder.
2. Open a terminal and run:
   ```bash
   python main.py
   ```
3. The output file `reviews_all_apk.xlsx` will be created in the same folder.

## Author
Created by Muhammad Miqdad Robbani on November 4, 2025.

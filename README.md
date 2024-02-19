# Transfer the contents of a .srt file into an Excel spreadsheet.
## Overview
It generates an Excel file with three distinct columns: "Entry Time", "Exit Time" and "Subtitle Text".

The script seamlessly populates each column with the relevant information extracted from the .srt file.

See "subs.png". Here you will see headers in Spanish: "Tiempo de entrada" (Entry Time), "Tiempo de salida" (Exit Time) and "Tiempo de subtítulo" (Subtitle Text).

## Requirements:
Python 3

openpyxl library

## Files
subs.py

## Usage
1. Rename your .srt file to "subtitles.srt" and save it in the same directory as "subs.py".
2. Run the "subs.py" script.
3. The script will produce the "subtitles.xlsx" file in the same directory.

## License
This project is governed by the GNU Affero General Public License v3.0. For comprehensive details, kindly refer to the LICENSE file included with this project.

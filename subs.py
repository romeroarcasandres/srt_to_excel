import re
from openpyxl import Workbook

# Function for processing the content of the srt file
def process_srt(content):
    pat = re.compile(r'(\d+)\s+(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})\n(.*?)(?=\n\n\d+|\Z)', re.DOTALL)
    coincidences = pat.findall(content)
    data = []
    for match in coincidences:
        _, entry_time, exit_time, text = match
        data.append({
            'Entry Time': entry_time,
            'Exit Time': exit_time,
            'Subtitle Text': text.strip().replace('\n', '\n')  # Replace \n with manual line breaks
        })
    return data

# Function for writing data to an Excel file
def write_excel(data):
    book = Workbook()
    sheet = book.active

    # Writing headers
    headers = ['Entry Time', 'Exit Time', 'Subtitle Text']
    sheet.append(headers)

    # Writing data to the Excel file
    for dato in data:
        row = [dato[header] for header in headers]
        sheet.append(row)

    # Save the Excel file
    book.save('subtitles.xlsx')

# Read the content of the srt file
with open('subtitles.srt', 'r', encoding='utf-8') as srt_file:
    srt_content = srt_file.read()

# Process the content of the srt file
data_subtitles = process_srt(srt_content)

# Write data to the Excel file
write_excel(data_subtitles)

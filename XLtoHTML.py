from bs4 import BeautifulSoup
import pandas as pd
import openpyxl

# Read your Excel file and specify the column name
df = pd.read_excel('filename.xlsx', sheet_name='Sheet1', header=1)  # Change the filename and sheet name as needed
column_name = 'column name'
# Load your Excel file for writing
wb = openpyxl.load_workbook('filename.xlsx')  # Replace 'Book5.xlsx' with your Excel file name

# Select the appropriate worksheet
sheet = wb['Sheet1']  # Change 'Sheet1' to your sheet name if different

# Specify the correct column reference (e.g., 'N' for column N)
column_reference = 'F'

max_row = df.shape[0]
# Loop through the rows in the specified column, starting from row 6
for row_number in range(3, max_row + 3):  # Assuming you want to process rows 6 to 10
    # Get the plain text from your DataFrame
    plain_text = str(df.loc[row_number - 3, column_name])

    # Initialize BeautifulSoup and create a <body> tag
    soup = BeautifulSoup('', 'html.parser')
    body = soup.new_tag('body')
    soup.append(body)

    # Split the plain text into paragraphs and list items
    sections = plain_text.split('\n')

    # Initialize variables for ordered and unordered lists
    ol_items = []
    ul_items = []

    # Process each section
    for section in sections:
        section_stripped = section.lstrip()
        if section_stripped and len(section_stripped) > 1 and section_stripped[0].isdigit() and section_stripped[1] == '.':
            # If it's an ordered list item, add it to ol_items
            ol_items.append(section.strip())
        elif section.strip() and section.lstrip()[0] == '-':
            # If it's an unordered list item, add it to ul_items
            ul_items.append(section.strip())
        else:
            # Process any previous ordered list items
            if ol_items:
                ol = soup.new_tag('ol')
                for item in ol_items:
                    li = soup.new_tag('li')
                    li.string = item[2:].strip()  # Remove the leading number and period
                    ol.append(li)
                body.append(ol)
                ol_items = []

            # Process any previous unordered list items
            if ul_items:
                ul = soup.new_tag('ul')
                for item in ul_items:
                    li = soup.new_tag('li')
                    li.string = item[1:].strip()  # Remove the leading "-"
                    ul.append(li)
                body.append(ul)
                ul_items = []

            # Create a paragraph for the current section
            paragraph = soup.new_tag('p')
            paragraph.string = section.strip()
            body.append(paragraph)
    # Get the well-formatted HTML
    formatted_html = soup.prettify()

    # Update the Excel cell with the formatted HTML
    cell = sheet[f"{column_reference}{row_number}"]
    cell.value = formatted_html

# Save the modified Excel file
wb.save('filename_converted.xlsx')

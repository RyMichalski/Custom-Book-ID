"""
    This script is used to read an Excel file, add a "Custom ID" column, and then generate the 
    Custom ID based on the values of the "Author" column and "Title" column.
"""

import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from nameparser import HumanName
from nameparser.config import CONSTANTS

FILE_PATH = Path("Dev Testing.xlsx")


# Read the excel file into the primary DataFrame
df = pd.read_excel(
    FILE_PATH, sheet_name="Sheet1", usecols=("Author", "Title", "Custom ID")
)


def extract_first_letters(title):
    """
    This function takes a title string and extracts the first letter of each word,
    while preserving numbers. 'The' is excluded if its the first word.
    Non-alphanumeric characters and spaces are ignored.

    Parameters:
    - title (str): The input title from which to extract first letters.

    Returns:
    str: A string containing the first letters of each word and/or numbers all together.

    Example:
    >>> extract_first_letters("The Lord of Rings")
    'LotR'
    >>> extract_first_letters("101 Dalimations")
    '101D'
    """

    words = re.findall(r"\b[\w\d]+\b", title)

    if words and words[0].lower() == "the":
        words = words[1:]

    return "".join(word[0] if word.isalpha() else word for word in words)


def process_author_initials(row):
    """
    Extracts and formats author initials from the "Author" column by checking
    for the character "&" and either .

    Args:
        row (pd.Series): A single row of the DataFrame.

    Returns:
        pd.Series: The row with the "initials" column added.
    """

    CONSTANTS.initials_format = "{first}{last}"

    # Check if the condition is met for this row
    if "&" not in row["Author"]:
        # Process for rows without "&"
        row["initials"] = (
            HumanName(row["Author"]).initials().replace(".", "").replace(" ", "")
        )
    else:
        # Process for rows with "&"
        authors = row["Author"].split("&")
        row["Author1"] = authors[0].strip()
        row["Author2"] = authors[1].strip() if len(authors) > 1 else ""

        row["initials1"] = (
            HumanName(row["Author1"]).initials().replace(".", "").replace(" ", "")
        )
        row["initials2"] = (
            HumanName(row["Author2"]).initials().replace(".", "").replace(" ", "")
        )
        row["initials"] = row["initials1"] + "&" + row["initials2"]

    return row


# Apply the function to each row
result_df = df.apply(process_author_initials, axis=1)


# Count number of books by author and iterate up by one.
def count_books(result_df):
    """Adds a "Count" column and assigns a unique two-digit count for each author's occurrences.

    Args:
        result_df (pd.DataFrame): DataFrame containing an "Author" column.

    Returns:
        pd.DataFrame: The modified DataFrame with the "Count" column added and values added.
    """
    result_df["Count"] = (
        result_df.groupby("Author").cumcount().add(1).astype(str).str.zfill(2)
    )


count_books(result_df)


# Create new Custom ID column and apply it

result_df["Custom ID"] = result_df.apply(
    lambda row: f"{row['initials']}_{extract_first_letters((row['Title']))}_{row['Count']}",
    axis=1,
)

workbook = load_workbook(FILE_PATH)
sheet = workbook["Sheet1"]


# Writes the "Custom ID" values from the DataFrame to original spreadsheet, starting in cell C2.
for index, r in enumerate(result_df["Custom ID"], start=2):
    sheet.cell(row=index, column=3, value=r)

# Save the file with the Custom ID data populated.
workbook.save(FILE_PATH)

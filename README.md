# Custom Book ID Project

This project is designed to read an Excel file, add a "Custom ID" column, and generate the Custom ID based on the values of the "Author" and "Title" columns. The project utilizes pandas and openpyxl to read and manipulate the Excel file. The library nameparser is used to identify and pull out initials from Author names in various formats.

## custom_id_generator.py

This script is the main part of the project. It handles reading the Excel file, generating the custom IDs, and writing the updated data back to a new Excel file.

The script has the following functions:

- `extract_first_letters(title)`: Extracts the first letters of each word in a title.
- `process_author_initials(author)`: Processes an author's name to extract their initials.
- `generate_ids()`: Generates a custom ID for each book and stores it in the `custom_id` property of the `Book` object.
- `update_excel()`: Updates the 'Custom ID' column in the DataFrame and writes the updated DataFrame back to a new Excel file.

## Usage

1. Install the required dependencies by running `pip install -r requirements.txt`.
2. Place the Excel file you want to process in the same directory as this script. The script expects the Excel file to have the following columns: Author, Title, and a blank Custom ID column.
3. Update the `FILE_PATH` variable in the script to match the name of your Excel file.
4. Run the script using `python custom_id_generator.py`.
5. The script will generate a new Excel file with the "Custom ID" column populated.

## Dependencies

- pandas
- openpyxl
- nameparser

## Author

- Ryan Michalski

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.
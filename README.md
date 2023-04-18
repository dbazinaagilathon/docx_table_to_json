#DOCX_TABLE_TO_JSON

This script is used to create a JSON file from a DOCX table.

The script is written in NodeJS and uses the following libraries:
    "jszip": "^3.10.1",
    "xml2js": "^0.4.23"

The script is used as follows:
    node index.js <docx_file_without_extension>

Example:
    node index.js example

The script will create a JSON file with the same name as the DOCX file.
The docx file must be in the spec directory where the output will be created.
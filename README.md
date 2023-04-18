# DOCX_TABLE_TO_JSON

This script is used to create a JSON file from a DOCX table.

The script is written in NodeJS and uses the following libraries:
  - "jszip": "^3.10.1",
  - "xml2js": "^0.4.23"

The script is used as follows:
    `node index.js <docx_file_without_extension>`

Example:
    `node index.js example`

The script will create a JSON file with the same name as the DOCX file.
The docx file must be in the spec directory where the output will be created.

## Known issues
- Sometimes the script will generate fields with [object][object] as part of the value.
This is due to the fact that the script is not able to parse the XML correctly. This is a known issue and will be fixed in the future. For now check the JSON file and remove the [object][object] part.

- The script will not work if there is more than one table in the DOCX file. Clean the DOCX file and remove all tables except the one you want to convert to JSON.
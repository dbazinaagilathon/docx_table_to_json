const fs = require("fs");
const JSZip = require("jszip");
const { promisify } = require("util");
const xml2js = require("xml2js");
const path = require("path");

const filePath = process.argv[2];

const readFileAsync = promisify(fs.readFile);

const TABLE = "w:tbl";
const TABLE_CELL = "w:tc";
const TABLE_ROW = "w:tr";
const PARAGRAPH = "w:p";
const RUNE = "w:r";
const TEXT = "w:t";

async function convertDocxToJson(docxFilePath) {
  const zip = new JSZip();
  const content = await readFileAsync(docxFilePath);
  const zipEntries = await zip.loadAsync(content);
  const documentXml = await zipEntries.file("word/document.xml").async("text");
  const parser = new xml2js.Parser({ explicitArray: false });
  const result = await parser.parseStringPromise(documentXml);
  return result;
}

const objectMapping = { pages: [] };

const concatTextElements = (textGroup) => {
  let text = "";
  try {
    textGroup.forEach((textGroup) => {
      if (Array.isArray(textGroup)) {
        textGroup.forEach((textGroup) => {
          if (textGroup[RUNE]) {
            if (Array.isArray(textGroup[RUNE])) {
              textGroup[RUNE].forEach((textContent) => {
                if (
                  textContent[TEXT] &&
                  typeof textContent[TEXT] === "string"
                ) {
                  text += textContent[TEXT];
                } else if (
                  textContent[TEXT] &&
                  typeof textContent[TEXT] === "object" &&
                  textContent[TEXT]["_"]
                ) {
                  text += textContent[TEXT]["_"];
                }
              });
            } else {
              const textContent = textGroup[RUNE];
              text += textContent[TEXT];
            }
          }
        });
      } else {
        if (textGroup[RUNE]) {
          if (Array.isArray(textGroup[RUNE])) {
            textGroup[RUNE].forEach((textContent) => {
              if (textContent[TEXT] && typeof textContent[TEXT] === "string") {
                text += textContent[TEXT];
              } else if (
                textContent[TEXT] &&
                typeof textContent[TEXT] === "object" &&
                textContent[TEXT]["_"]
              ) {
                text += textContent[TEXT]["_"];
              }
            });
          } else {
            const textContent = textGroup[RUNE];
            text += textContent[TEXT];
          }
        }
      }
    });
    return text;
  } catch (error) {
    text = "********* FAILED ********";
    return text;
  }
};

const getArrayElements = (arrayOfQuestionsOrAnswers) => {
  let array = [];
  try {
    arrayOfQuestionsOrAnswers.forEach((textGroup) => {
      if (textGroup[RUNE]) {
        if (!Array.isArray(textGroup[RUNE])) {
          array.push(textGroup[RUNE][TEXT]);
        } else {
          let addedText = "";
          textGroup[RUNE].forEach((textContent) => {
            if (textContent[TEXT] && typeof textContent[TEXT] === "string") {
              addedText += textContent[TEXT];
            } else if (
              textContent[TEXT] &&
              typeof textContent[TEXT] === "object" &&
              textContent[TEXT]["_"]
            ) {
              addedText += textContent[TEXT]["_"];
            }
          });
          array.push(addedText);
        }
      }
    });
    return array;
  } catch (error) {
    array = ["********* FAILED ********"];
    return array;
  }
};

(async () => {
  try {
    const json = await convertDocxToJson(
      path.join(__dirname, "./", "spec", `${filePath}.docx`)
    );
    const table = json["w:document"]["w:body"][TABLE][TABLE_ROW];
    table.slice(1).forEach((row) => {
      const rowData = {
        stepName: "",
        shortQuestionText: "",
        title: "",
        screenText: "",
        stepType: "",
        answerValues: "",
        responseValues: "",
        branchingLogic: "",
        additionalDetails: "",
      };

      const textExtractor = (index, field) => {
        if (Array.isArray(row[TABLE_CELL][index][PARAGRAPH])) {
          rowData[field] = concatTextElements(
            row[TABLE_CELL][index][PARAGRAPH]
          );
        } else {
          if (!Array.isArray(row[TABLE_CELL][index][PARAGRAPH][RUNE])) {
            if (
              typeof row[TABLE_CELL][1][PARAGRAPH][RUNE][TEXT] === "object" &&
              row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT]["_"]
            ) {
              rowData[field] =
                row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT]["_"];
            } else {
              if (row[TABLE_CELL][index][PARAGRAPH][RUNE]) {
                if (
                  row[TABLE_CELL][index][PARAGRAPH][RUNE] &&
                  row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT] &&
                  typeof row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT] ===
                    "string"
                ) {
                  rowData[field] =
                    row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT];
                } else {
                  rowData[field] =
                    row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT]["_"];
                }
              }
            }
          } else {
            rowData[field] = concatTextElements([
              row[TABLE_CELL][index][PARAGRAPH],
            ]);
          }
        }
      };

      const arrayExtractor = (index, field) => {
        if (Array.isArray(row[TABLE_CELL][index][PARAGRAPH])) {
          rowData[field] = getArrayElements(row[TABLE_CELL][index][PARAGRAPH]);
        } else {
          if (!Array.isArray(row[TABLE_CELL][index][PARAGRAPH][RUNE])) {
            if (
              typeof row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT] ===
                "object" &&
              row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT]["_"]
            ) {
              rowData[field] =
                row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT]["_"];
            } else {
              if (
                typeof row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT] ===
                "string"
              ) {
                rowData[field] = row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT];
              } else {
                rowData[field] =
                  row[TABLE_CELL][index][PARAGRAPH][RUNE][TEXT]["_"];
              }
            }
          } else {
            rowData[field] = getArrayElements([row[TABLE_CELL][7][PARAGRAPH]]);
          }
        }
      };

      textExtractor(1, "stepName");
      textExtractor(2, "shortQuestionText");
      textExtractor(3, "title");
      textExtractor(4, "screenText");
      textExtractor(5, "stepType");
      textExtractor(8, "branchingLogic");
      textExtractor(9, "additionalDetails");

      if (!rowData.stepType.toLowerCase().includes("choice")) {
        textExtractor(6, "answerValues");
        textExtractor(7, "responseValues");
      } else {
        arrayExtractor(6, "answerValues");
        arrayExtractor(7, "responseValues");
      }

      const {
        stepName,
        shortQuestionText,
        title,
        screenText,
        stepType,
        answerValues,
        responseValues,
        branchingLogic,
        additionalDetails,
      } = rowData;
      objectMapping.pages.push({
        [row[TABLE_CELL][0][PARAGRAPH][RUNE][TEXT]]: {
          stepName,
          shortQuestionText,
          title,
          screenText,
          stepType,
          answerValues,
          responseValues,
          branchingLogic,
          additionalDetails,
        },
      });
    });
    JSON.stringify(objectMapping);
    fs.writeFileSync(
      path.join(__dirname, "./", "spec", `${filePath}.json`),
      JSON.stringify(objectMapping, null, 2)
    );
  } catch (error) {
    console.error(error);
  }
})();

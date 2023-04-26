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
const UNDERLINE = "w:u"
const BOLD = "w:b"
const RUNPROP = "w:rPr"
const VALUE = "w:val"
const DOCUMENT = "w:document"
const BODY = "w:body"
const TEXT_CHOICE = "text_choice"
const NUMERIC = "numeric"
const reviewVerbiageText = "Please review your responses before you submit this form. If you see a response that you wish to change, select 'edit'."
const buttonText = { next: "Next",edit: "edit",back: "Back",start: "Start" }
const objectMapping = { reviewVerbiage: reviewVerbiageText, button: buttonText, pages: {} };
const formStepsMapping = { formSteps: []}
let mappingKey = 0

async function convertDocxToJson(docxFilePath) {
  const zip = new JSZip();
  const content = await readFileAsync(docxFilePath);
  const zipEntries = await zip.loadAsync(content);
  const documentXml = await zipEntries.file("word/document.xml").async("text");
  const parser = new xml2js.Parser({ explicitArray: false });
  const result = await parser.parseStringPromise(documentXml);
  return result;
}
function checkBoldOrUnderline(textContent, textContentBody)
{
  const underline = textContent[RUNPROP] && textContent[RUNPROP][UNDERLINE] && textContent[RUNPROP][UNDERLINE]["$"]&& textContent[RUNPROP][UNDERLINE]["$"][VALUE]=="single";
  const bold = textContent[RUNPROP]&& textContent[RUNPROP].hasOwnProperty(BOLD)
  let text = ""
  if(underline && bold){
    text+=  "<b><u>"+textContentBody+"</u></b>"
  }
  else if(underline){
    text+=  "<u>"+textContentBody+"<u/>"
  }
  else if(bold){
    text+=  "<b>"+textContentBody+"<b/>"
  }
  else{
    text+=  textContentBody;
  }
  return text
}
function checkTextSource(textContent, checkStyle)
{
  let text = ""
  if (textContent[TEXT] && typeof textContent[TEXT] === "string") {
    text += checkStyle ? checkBoldOrUnderline(textContent,textContent[TEXT]) : textContent[TEXT];
  } else if (textContent[TEXT] &&typeof textContent[TEXT] === "object" &&textContent[TEXT]["_"]) {
    text +=  checkStyle ? checkBoldOrUnderline(textContent,textContent[TEXT]["_"]) : textContent[TEXT]["_"];
  }
  else if(textContent[TEXT] && textContent[TEXT]["$"] && textContent[TEXT]["$"]["xml:space"]){
    text += " "
  }
  return text
}
const concatTextElements = (textGroup) => {
  let text = "";
  try {
    textGroup.forEach((textGroup) => {
      if (Array.isArray(textGroup)) {
        textGroup.forEach((textGroup) => {
          if (textGroup[RUNE]) {
            if (Array.isArray(textGroup[RUNE])) {
              textGroup[RUNE].forEach((textContent) => {
                text+=checkTextSource(textContent, false)
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
            if(textGroup[RUNE].every(obj => !obj.hasOwnProperty(TEXT)))
            {
              text+= "<br/>"
            }
            textGroup[RUNE].forEach((textContent) => {
            text+=checkTextSource(textContent, true)
            });
          } else {
            text+=checkTextSource(textGroup[RUNE], false)
          }
        }        
        else{
          text+= "<br/>"
        }
      }
    });
    text = text.replace(/^(<br\/>)+|(<br\/>)+$/g, '') //erases break lines from the beginning or the end
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
            addedText+= checkTextSource(textContent, false)
          });
          if(addedText)
          {
            array.push(addedText);
          }
        }
      }
    });
    return array;
  } catch (error) {
    array = ["********* FAILED ********"];
    return array;
  }
};
function createFormSteps(object)
{
  const step =  {
    mappingKey : object.mappingKey,
    name : object.stepName,
    cDash : object.shortQuestionText,
    type : object.stepType.toLowerCase().replace(/\s/g, "_")
  }
  if(step.type === TEXT_CHOICE)
  {
    step.choices = []
    object.answerValues.forEach(value => {
      step.choices.push({text:value, value: object.responseValues[step.choices.length]})
    });
  }
  else if(step.type === NUMERIC)
  {
    const nums = object.responseValues.split("-").map(Number);
    step.min = nums[0]
    step.max = nums[1]
  }
  return step
}
(async () => {
  try {
    const json = await convertDocxToJson(
      path.join(__dirname, "./", "spec", `${filePath}.docx`)
    );
    const table = json[DOCUMENT][BODY][TABLE][0][TABLE_ROW];
    let count = 1
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
      objectMapping.pages[count]= {
          stepName,
          shortQuestionText,
          title,
          screenText,
          stepType,
          answerValues,
          responseValues,
          branchingLogic,
          additionalDetails,
        }
      if(objectMapping.pages[count].shortQuestionText !== "N/A")
      {
        objectMapping.pages[count].mappingKey = mappingKey
        const step = createFormSteps(objectMapping.pages[count])
        formStepsMapping.formSteps.push(step)
        mappingKey += 1
      }
      count +=1
    });
    JSON.stringify(objectMapping);
    fs.writeFileSync(path.join(__dirname, "./", "spec", `${filePath}.json`),JSON.stringify(objectMapping, null, 2));
    fs.writeFileSync(path.join(__dirname, "./", "spec", `${filePath}FormSteps.json`),JSON.stringify(formStepsMapping, null, 2));
  } catch (error) {
    console.error(error);
  }
})();

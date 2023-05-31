const fs = require("fs");
const JSZip = require("jszip");
const { promisify } = require("util");
const xml2js = require("xml2js");
const path = require("path");
const fsp = fs.promises;

const readFileAsync = promisify(fs.readFile);

//#region constants

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
const buttonText = { next: "Next",edit: "edit",back: "Back",start: "Start" }

//#endregion

async function findDocxFileInDirectory(directoryPath) {
  const files = await fsp.readdir(directoryPath);
  const docxFiles = files.filter(file => path.extname(file) === ".docx");
  if (docxFiles.length > 0) {
    return docxFiles;
  } else {
    throw new Error(`No DOCX files found in directory ${directoryPath}`);
  }
}
async function convertDocxToJson(docxFilePath) {
  const zip = new JSZip();
  const content = await readFileAsync(docxFilePath);
  const zipEntries = await zip.loadAsync(content);
  const documentXml = await zipEntries.file("word/document.xml").async("text");
  const parser = new xml2js.Parser({ explicitArray: false });
  const result = await parser.parseStringPromise(documentXml);
  return result;
}
const checkBoldOrUnderline = (textContent, textContentBody) => {
  const underline = textContent[RUNPROP]?.[UNDERLINE]?.["$"]?.[VALUE] === "single";
  const bold = textContent[RUNPROP]?.[BOLD];
  let text = ""
  if(underline && bold){
    text+=  "<b><u>"+textContentBody+"</u></b>"
  }
  else if(underline){
    text+=  "<u>"+textContentBody+"</u>"
  }
  else if(bold){
    text+=  "<b>"+textContentBody+"</b>"
  }
  else{
    text+=  textContentBody;
  }
  return text
}
const checkTextSource = (textContent, checkStyle)=> {
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
          array.push(checkTextSource(textGroup[RUNE], false));
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
const textExtractor = (index, field, row, rowData) => {
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
const arrayExtractor = (index, field, row, rowData) => {
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
const createFormSteps = (object, mappingKey) => {
  const step =  {
    mappingKey : mappingKey,
    name : object.stepName,
    cDash : object.shortQuestionText,
    type : object.stepType.toLowerCase().replace(/\s/g, "_")
  }
  if(step.type.includes(TEXT_CHOICE))
  {
    step.choices = []
    object.answerValues.forEach(value => {
      step.choices.push({text:value, value: object.responseValues[step.choices.length]})
    });
    step.type = TEXT_CHOICE
  }
  else if(step.type.includes(NUMERIC))
  {
    const nums = object.responseValues.split("-").map(Number);
    step.min = nums[0]
    step.max = nums[1]
  }
  return step
}
const mappingToJson = (rowData, objectMapping) =>
{
  if(rowData.stepName.includes("Copyright")){
    objectMapping.title = rowData.title
    return { title:rowData.title, copyright: rowData.stepName + " " + rowData.screenText }
  }
  else if (rowData.stepName.includes("Instruction")){
    return { instructionalMessages: [rowData.screenText.split("<br/>")] }
  }
  else{
    return { mappingKey: "", question: rowData.screenText, answers: rowData.answerValues }
  }}

(async () => {
  try {
    const directoryPath = path.join(__dirname, "./", "spec")
    const docxFilePath = await findDocxFileInDirectory(directoryPath);
    docxFilePath.forEach(async(doc)=>{
      const objectMapping = { title: "", reviewVerbiage: "", button: buttonText, pages: {} };
      const formStepsMapping = { formSteps: []}
      let mappingKey = 0
      let count = 1
      const json = await convertDocxToJson(path.join(directoryPath, doc));
      const table = json[DOCUMENT][BODY][TABLE].length > 1 ? json[DOCUMENT][BODY][TABLE][0][TABLE_ROW] : json[DOCUMENT][BODY][TABLE][TABLE_ROW];
      table.slice(1).forEach((row) => {
        const rowData = {
          stepName: "",
          shortQuestionText: "",
          title: "",
          screenText: "",
          stepType: "",
          answerValues: "",
          responseValues: "",
          //branchingLogic: "",
          //additionalDetails: "",
        };

        textExtractor(1, "stepName", row, rowData);
        textExtractor(2, "shortQuestionText", row, rowData);
        textExtractor(3, "title", row, rowData);
        textExtractor(4, "screenText", row, rowData);
        textExtractor(5, "stepType", row, rowData);
        // textExtractor(8, "branchingLogic", row);
        // textExtractor(9, "additionalDetails", row);
        if (!rowData.stepType.toLowerCase().includes("choice")) {
          textExtractor(6, "answerValues", row, rowData);
          textExtractor(7, "responseValues", row, rowData);
        } else {
          arrayExtractor(6, "answerValues", row, rowData);
          arrayExtractor(7, "responseValues", row, rowData);
        }
        if(rowData.stepType.toLowerCase().includes("review"))
        {
          objectMapping.reviewVerbiage = rowData.screenText
          return;
        }
        if(rowData.stepType.toLowerCase().includes("completion")){
          return
        }
        objectMapping.pages[count]= mappingToJson(rowData, objectMapping)
        if(rowData.shortQuestionText !== "N/A" && !rowData.stepType.toLowerCase().includes("instruction"))
        {
          objectMapping.pages[count].mappingKey = mappingKey.toString()
          formStepsMapping.formSteps.push(createFormSteps(rowData, mappingKey.toString()))
          mappingKey += 1
        }
        count +=1
      });
      JSON.stringify(objectMapping);
      fs.writeFileSync(path.join(directoryPath,doc.replace(".docx", ".json")),JSON.stringify(objectMapping, null, 2));
      fs.writeFileSync(path.join(directoryPath, "formSteps"+ doc.replace(".docx", ".json")),JSON.stringify(formStepsMapping, null, 2));
      console.log("Finished "+doc)
    })
  } catch (error) {
    console.error(error);
  }
})();

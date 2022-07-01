const axios = require('axios')
const superagent = require('superagent');
//ignore it saying defaults doesnt exist, it does and using default does not work.
//makes sure put requests uses the proper content-type header
axios.defaults.headers.put['Content-Type'] = "application/json"
axios.defaults.headers.put['accept'] = "application/json"

import { Data, tempDataStore, params, templates } from './model.js'
import {
  disableButton,
  retrieveStyles,
  validateHierarchy,
  validateTestSteps
} from './taskpane.js';

/*
  Functions that should be moved over: 
  disableButton,
  retrieveStyles,
  validateHierarchy,
  validateTestSteps
*/

/*Functions that have been integrated into parseArtifacts: pushArtifacts,
 updateSelectionArray, retrieveImages, parseRequirements, retrieveTables,
 parseTestCases, parseTestSteps
 */

/*LINES THAT NEED ATTENTION: 103 165 204 */

var RETRIEVE = "http://localhost:5000/retrieve"
//params:
//ArtifactTypeId: ID based on params.artifactEnums.{artifact-type}

/*This function takes in an ArtifactTypeId and parses the selected text or full body
of a document, returning digestable objects which can easily be used to send artifacts
to spira.*/
const parseArtifacts = async (ArtifactTypeId) => {
  return Word.run(async (context) => {
    disableButton("send-to-spira-button");
    /***************************
     Start range identification
     **************************/
    let projectId = document.getElementById("project-select").value
    let selection = context.document.getSelection();
    let splitSelection = context.document.getSelection().split(['\r']);
    context.load(selection, ['text', 'inlinePictures', 'style', 'styleBuiltIn']);
    context.load(splitSelection, ['text', 'inlinePictures', 'style', 'styleBuiltIn'])
    await context.sync();
    //If there is no selected area, parse the full body of the document instead
    if (!selection.text) {
      selection = context.document.body.getRange();
      splitSelection = context.document.body.getRange().split(['\r']);
      context.load(selection, ['text', 'inlinePictures', 'style', 'styleBuiltIn']);
      context.load(splitSelection, ['text', 'inlinePictures', 'style', 'styleBuiltIn'])
      await context.sync();
    }
    /********************** 
     Start image formatting
    ***********************/
    var imageLines = []
    //i represents each "line" delimited by \r tags
    for (let i = 0; i < splitSelection.items.length; i++) {
      //this checks if there are any images at all on a given line
      if (splitSelection.items[i].inlinePictures.items[0]) {
        //this pushes the 'line number' for each image on a particular line.
        for (let line in splitSelection.items[i].inlinePictures.items) {
          imageLines.push(i)
        }
      }
    }
    var imageObjects = [];
    let images = selection.inlinePictures;
    for (let i = 0; i < images.items.length; i++) {
      let base64 = images.items[i].getBase64ImageSrc();
      await context.sync();
      let imageObj = new templates.Image(base64.m_value, `inline${i}.jpg`, imageLines[0])
      imageObjects.push(imageObj)
      //removes the first item in imageLines as it has been converted to an imageObject;
      imageLines.shift();
    }
    //end of image formatting
    /**********************
     Start Artifact Parsing
    ***********************/
    const bodyRegex = params.regexs.bodyRegex
    const bodyTagRegex = params.regexs.bodyTagRegex
    switch (ArtifactTypeId) {
      //this will parse the data assuming the user wants requirements to be imported.
      case (params.artifactEnums.requirements): {
        /*this returns the styles in an array to be referenced 
        against text in the document*/
        let styles = retrieveStyles('req-');
        let descStart, descEnd;
        //var for later unscoped access to this variable when artifacts are sent to spira
        var requirements = [];
        let body = splitSelection;
        let requirement = new templates.Requirement()
        for (let [i, item] of body.items.entries()) {
          //style stores custom styles while styleBuiltIn only stores default styles
          if (styles.includes(item.styleBuiltIn) || styles.includes(item.style)) {
            /*this wraps up the description of the previous requirement, pushes to
            requirements array, and creates a new requirement object*/
            if (descStart) {
              descEnd = body.items[i - 1]
              /*creates a description range given the beginning 
              and end as delimited by lines with mapped styles*/
              let descRange = descStart.expandToOrNullObject(descEnd)
              await context.sync();
              //descRange is null if the descEnd is not valid to extend the descStart
              if (!descRange) {
                descRange = descStart
              }
              //clears these fields so it is known the description has been added to its requirement
              descStart = undefined
              descEnd = undefined
              let descHtml = descRange.getHtml();
              await context.sync();
              //m_value is the actual string of the html
              /**
               * filter against regex here and remove body tags
               */
              requirement.Description = await filterForLists(descHtml.m_value.replaceAll("\r", ""));
              requirements.push(requirement)
              /*gets the requirement of the to be created requirement based
               on its index in the styles array.*/
              let indent = styles.indexOf(item.styleBuiltIn)
              if (!indent) {
                indent = styles.indexOf(item.style)
              }
              //creates new requirement and populates relevant fields 
              requirement = new templates.Requirement();
              requirement.Name = item.text.replaceAll("\r", "")
              requirement.IndentLevel = indent
            }
            else if (requirement.Name) {
              requirements.push(requirement)
              /*gets the requirement of the to be created requirement based
               on its index in the styles array.*/
              let indent = styles.indexOf(item.styleBuiltIn)
              if (!indent) {
                indent = styles.indexOf(item.style)
              }
              //creates new requirement and populates relevant fields 
              requirement = new templates.Requirement();
              requirement.Name = item.text.replaceAll("\r", "")
              requirement.IndentLevel = indent
            }
            //This should only happen for the first indent level 1 line parsed
            else {
              requirement.Name = item.text.replaceAll("\r", "")
            }
          }
          else {
            if (!descStart) {
              descStart = item
            }
          }
          /*if a line is the last line in the selection/body
          & the current requirement object has least a name,
          & the name is not the current line's text 
          push the last requirement*/
          if (i == body.items.length - 1 && requirement.Name && item.text != requirement.Name) {
            if (item.text) {
              let descRange;
              //if it has a start - expand the range to the end of the description block
              if (descStart) {
                descEnd = item
                descRange = descStart.expandToOrNullObject(descEnd)
              }
              else {
                descRange = item
              }
              await context.sync()
              //this covers if the expandToOrNullObject returns null
              if (!descRange) {
                descRange = descStart
              }
              let descHtml = descRange.getHtml();
              await context.sync();
              /**
               * Filter against regex here as well - also need to remove body tags
               */
              requirement.Description = await filterForLists(descHtml.m_value.replaceAll("\r", ""));
            }
            requirements.push(requirement)
          }
        }
        if (!validateHierarchy(requirements)) {
          requirements = false
          document.getElementById("send-to-spira-button").disable = false
          displayError("heirarchy", true)
          //throw hierarchy error and exit
        }
        //if the heirarchy is invalid, clear requirements and throw error
        sendArtifacts(params.artifactEnums.requirements, imageObjects, requirements, projectId)
        return requirements
      }
      case (params.artifactEnums.testCases): {
        /*when parsing multiple tables tableCounter serves as an index throughout all
        parts of the function*/
        let tableCounter = 0
        let testCases = []
        let styles = retrieveStyles('test-')
        let testCase = new templates.TestCase()
        let descStart, descEnd;
        /********************
         * Parsing out tables
         ********************/
        let selectionTables = selection.tables
        context.load(selectionTables)
        await context.sync();
        //tables = 3d array [table][test step][column], only contains text
        var tables = [];
        for (let i = 0; i < selectionTables.items.length; i++) {
          let table = selectionTables.items[i].values;
          tables.push(table)
        }
        /*May not need this, but it is the images in the FIRST TABLE. If i need it I 
        will make it procedural for each given table when images are placed in spira.*/
        let tableImages = selectionTables.items[0].getRange().inlinePictures;
        context.load(tableImages)
        await context.sync();
        if (!validateTestSteps(tables, styles[2])) {
          //validateTestSteps throws the error
          return false
        }
        /*part of this portion isnt DRY, but due to not being able to pass Word 
        objects between functions cant be made into its own function*/
        for (let [i, item] of splitSelection.items.entries()) {
          //this just removes excess tags
          let itemtext = item.text.replaceAll("\r", "")
          if (styles.includes(item.style) || styles.includes(item.styleBuiltIn)) {
            if (descStart) {
              descEnd = splitSelection.items[i - 1]
              let descRange = descStart.expandToOrNullObject(descEnd);
              context.load(descRange)
              await context.sync();
              /*if the descRange returns null (doesnt populate a range), assume the range
              is only the starting line.*/
              let descHtml = descRange.getHtml();
              await context.sync();
              descStart = undefined; descEnd = undefined;
              if (!testCase.Name) {
                //this doesnt need formatting
                testCase.folderDescription = itemtext
              }
              else {
                //removes tables picked up in the description and adds proper HTML lists
                let filteredDescription = filterForLists(descHtml.m_value.replaceAll("\r", "")); // filter for LISTS!!!
                let tableRegex = params.regexs.tableRegex
                /*descriptionBody parses out the body, Tables
                 matches all tables to be removed as well*/
                let descriptionBody = filteredDescription.match(bodyRegex)[0]
                let descriptionTables = [...descriptionBody.matchAll(tableRegex)]
                for (let j = 0; j < descriptionTables.length; j++) {
                  filteredDescription = filteredDescription.replace(descriptionTables[j][0], "")
                }
                testCase.testCaseDescription = filteredDescription
              }
              if (item.style == styles[0] || item.styleBuiltIn == styles[0]) {
                if (testCase.Name) {
                  testCases.push(testCase)
                }
                testCase = new templates.TestCase();
                testCase.folderName = itemtext
              }
              else {
                if (testCase.Name) {
                  testCases.push(testCase)
                  //makes a new testCase object with the old folderName in case there is not a new one.
                  testCase = new templates.TestCase();
                  testCase.folderName = testCases[testCases.length - 1].folderName
                }
                testCase.Name = itemtext
              }
            }
            /*if there is already a test case name and a a new test case / folder name is detected,
            push the test case to testCases array, then empty the test case object*/
            else if (testCase.Name) {
              testCases.push(testCase)
              //if the currently checked line has styles[0] as its style, put the folder name as the items text
              if (item.style == styles[0] || item.styleBuiltIn == styles[0]) {
                testCase = { folderName: itemtext, folderDescription: "", Name: "", testCaseDescription: "", testSteps: [] }
              }
              //else it must have styles[1] (folder name mapping) as its style, so set as new test case's name
              else {
                testCase = { folderName: testCases[testCases.length - 1].folderName, folderDescription: "", Name: item.text, testCaseDescription: "", testSteps: [] }
              }
            }
            else {
              if (item.style == styles[0] || item.styleBuiltIn == styles[0]) {
                testCase.folderName = itemtext
              }
              //else it must have styles[1] (test case name mapping) as its style, so set as new test case's name
              else {
                testCase.Name = itemtext
              }
              continue
            }
          }
          else if (tables[0] && item.text == tables[0][0][parseInt(styles[2].slice(-1)) - 1]?.concat("\t") && item.text.slice(-1) == "\t") {
            //This procs when there is a table and the first description equals item.text
            //testStepTable = 2d array [row][column]
            /**************************
             *start parseTestSteps Logic*
            ***************************/
            //this is the length of a row, needed for organizing the tables for parsing
            let length = selectionTables.items[tableCounter].values[0].length
            let tableElementRegex = params.regexs.paragraphRegex
            for (let item of splitSelection.items) {
              let html = selectionTables.items[tableCounter].getRange().getHtml();
              await context.sync();
              let elements = [...html.m_value.matchAll(tableElementRegex)]
              var formattedStrings = []
              for (let [i, element] of elements.entries()) {
                //if the row exists, place the current element into its relevant box
                if (formattedStrings[Math.floor(i / length)]) {
                  /*element is a single match from elements - [0] is the full
                   string of the match (as per RegExp String Iterator syntax)*/
                  formattedStrings[Math.floor(i / length)][i % length] = (element[0])
                }
                //if a row doesnt exist, create it, then add the current element
                else {
                  formattedStrings[Math.floor(i / length)] = []
                  formattedStrings[Math.floor(i / length)][i % length] = (element[0])
                }
              }
            }
            let testStepTable = formattedStrings
            /************************
            *end parseTestSteps Logic*
            ************************/
            //table counter lets parseTestSteps know which table is currently being parsed
            tableCounter++
            let testStep = new templates.TestStep();
            let testSteps = []
            //this is true when the "Header rows?" box is checked
            let headerCheck = document.getElementById("header-check").checked
            if (headerCheck) {
              //if the user says there are header rows, remove the first row of the table being parsed.
              testStepTable.shift();
            }
            //take testStepTable and put into test steps
            for (let [i, row] of testStepTable.entries()) {
              //skips lines with empty descriptions to prevent pushing empty steps (returns null if no match)
              let emptyStepRegex = params.regexs.emptyParagraphRegex

              if (row[parseInt(styles[2].slice(-1)) - 1].match(emptyStepRegex)) {
                continue
              }
              testStep = { Description: row[parseInt(styles[2].slice(-1)) - 1], ExpectedResult: row[parseInt(styles[3].slice(-1)) - 1], SampleData: row[parseInt(styles[4].slice(-1)) - 1] }
              testSteps.push(testStep)
            }
            testCase.testSteps = testSteps
            //removes the table that has been processed from this functions local reference
            tables.shift();
          }
          else if (!descStart && (item.text.slice(-1) != "\t")) {
            descStart = item
          }
          if (i == (splitSelection.items.length - 1)) {
            if (descStart) {
              descEnd = splitSelection.items[i]
              let descRange = descStart.expandToOrNullObject(descEnd);
              await context.sync();
              /*if the descRange returns null (doesnt populate a range), assume the range
              is only the starting line.*/
              let descHtml = descRange.getHtml();
              await context.sync();
              let filteredDescription = filterForLists(descHtml.m_value.replaceAll("\r", ""))
              //This removes tables picked up in the description
              let tableRegex = /<table(.|\n|\r)*?\/table>/g
              let descriptionTables = [...filteredDescription.matchAll(tableRegex)]
              for (let j = 0; j < descriptionTables.length; j++) {
                filteredDescription = filteredDescription.replace(descriptionTables[j][0], "")
              }
              testCase.testCaseDescription = filteredDescription
            }
            //dont push a nameless testCase.
            if (testCase.Name) {
              testCases.push(testCase)
            }
          }
        }
        sendArtifacts(params.artifactEnums.testCases, imageObjects, testCases, projectId)
        return testCases
      }
    }
  })
}
//params:
//images = [{base64: b64encoded string, name: "", lineNum: ""}]
//ArtifactTypeId based on params.artifactEnums.{artifact-type}
//Artifacts is the array of parsed artifacts ready for sending

/*This function takes the already parsed artifacts and images and sends
them to spira*/
const sendArtifacts = (ArtifactTypeId, images, Artifacts, projectId) => {
  //this checks if an empty artifact array is passed in (should never happen)
  if (Artifacts.length == 0) {
    //empty is the error message key for the model object.
    displayError("empty", true);
    document.getElementById("send-to-spira-button").disabled = false;
    return
  }
}


/* Filters a string and changes any word-outputted lists to properly formatted html lists. INDENTING IS NOT YET IMPLEMENTED*/
const filterForLists = (description) => {
  let startRegEx = params.regexs.paragraphRegex;
  let elemList = [...description.matchAll(startRegEx)];
  description = convertToIndentedList(description, elemList);
  return description
}

/* Scans each element in an array of 'strings' for "style='margin-left:#.0in" where the # is the indent level 
   Then it keeps track of the current indent level as it loops through the array, processing the elements
   through convertToListElem and adding an extra <ul> or <ol> as necessary to properly turn them into html 
   lists. */
const convertToIndentedList = (description, elemList) => {
  let indentLevel = 0;
  let lastOrdered = false;
  for (let i = 0; i < elemList.length; i++) {
    /* Use elemList[i][0] in order to reach the matched strings. */
    let elem = elemList[i][0];
    let match = elem.match(params.regexs.marginRegEx);
    let result = convertToListElem(elem);
    let alteredElem = result.elem;
    let ordered = result.ordered;
    if (match) {
      let curIndentLevel = (parseInt(match[1]) * 2 + parseInt(match[2]) * 0.2) - 1
      while (curIndentLevel > indentLevel) {
        alteredElem = listDelimiter(alteredElem, true, false, ordered);
        indentLevel++;
      }
      while (curIndentLevel < indentLevel) {
        alteredElem = listDelimiter(alteredElem, false, true, lastOrdered);
        indentLevel--;
      }
      lastOrdered = ordered;
    }
    else {
      let curIndentLevel = 0;
      while (curIndentLevel < indentLevel) {
        alteredElem = listDelimiter(alteredElem, false, true, lastOrdered);
        indentLevel--;
      }
    }
    description = description.replace(elem, alteredElem);
  }
  return description;
}

/* Takes a single <p> to </p> element and turns it into a list element if it has the necessary class*/
const convertToListElem = (pElem) => {
  let listElem = pElem + "";
  let ordered = !(listElem.includes(">·<span") || listElem.includes(">o<span") || listElem.includes(">§<span"));
  let orderedRegEx = params.regexs.orderedRegEx;
  if (listElem.includes("class=MsoListParagraphCxSpFirst")) { //Case for if the element is the first element in a list
    //Must add extra html element codes at the beginning and end of the list to wrap the list elements together.
    listElem = listDelimiter(listElem, true, false, ordered); // starts a list
  }
  else if (listElem.includes("class=MsoListParagraphCxSpLast")) { //Case for if the element is the last element in a list.
    listElem = listDelimiter(listElem, false, false, ordered); // ends a list
  }
  else if (listElem.includes("class=MsoListParagraph ")) { //Case for if the element is the only element in the list
    listElem = listDelimiter(listElem, true, false, ordered); // starts a list
    listElem = listDelimiter(listElem, false, false, ordered); // ends a list
  }
  if (listElem.includes("class=MsoListParagraph")) { // This will happen for every element that is part of a list
    listElem = listElem.replace("<p ", "<li ").replace("</p>", "</li>").replaceAll(orderedRegEx, "><span");
    listElem = listElem.replaceAll("&nbsp;", "");
  }
  //Case for if the element is not part of a list is handled by just returning it back.
  return { elem: listElem, ordered: ordered };
}

/* Adds a <ul> or <ol> element based on the parameters and if the element is an unordered or ordered list. */
/* endPrefix = true means that it should put the </ul> or </ol> BEFORE the element instead of after. */
const listDelimiter = (elem, start, endPrefix, ordered) => {
  // Checks for if is affecting an ordered or unordered list.
  let orderedMarker = ordered ? "ol" : "ul";
  // Checks if it is starting or ending a list
  let listMarker = `<${start ? "" : "/"}${orderedMarker}>`
  // Checks where it should position the list marker
  elem = endPrefix || start ? `${listMarker}${elem}` : `${elem}${listMarker}`;
  return elem;
}

export {
  parseArtifacts
}
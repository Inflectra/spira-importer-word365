const axios = require('axios')
const superagent = require('superagent');
//ignore it saying defaults doesnt exist, it does and using default does not work.
//makes sure put requests uses the proper content-type header
axios.defaults.headers.put['Content-Type'] = "application/json"
axios.defaults.headers.put['accept'] = "application/json"

export {
  parseArtifacts
}

import { Data, tempDataStore, params, templates } from './model.js'
import {
  disableButton,
  retrieveStyles,
  validateHierarchy,
  validateTestSteps,
  displayError,
  pushImage,
  showProgressBar,
  updateProgressBar,
  indentRequirement,
  hideProgressBar,
  enableButton
} from './taskpane.js';

/*
  Functions that should be moved over: 
  retrieveStyles,
  validateHierarchy,
  validateTestSteps
*/

/*Functions that have been integrated into parseArtifacts: pushArtifacts,
 updateSelectionArray, retrieveImages, parseRequirements, retrieveTables,
 parseTestCases, parseTestSteps
 */

/*LINES THAT NEED ATTENTION: 121 175 214 */

var RETRIEVE = "http://localhost:5000/retrieve"

//params:
//ArtifactTypeId: ID based on params.artifactEnums.{artifact-type}
//model: user model object based on Data() model. (contains user credentials)

/*This function takes in an ArtifactTypeId and parses the selected text or full body
of a document, returning digestable objects which can easily be used to send artifacts
to spira.*/
const parseArtifacts = async (ArtifactTypeId, model) => {
  return Word.run(async (context) => {
    disableButton(params.buttons.sendToSpira);
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
              let descRange = descStart.expandTo(descEnd)
              context.load(descRange)
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
              //this filters out unwanted html clutter and formats lists
              let filteredDescription = filterDescription(descHtml.m_value)
              requirement.Description = filteredDescription
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
              let filteredDescription = filterDescription(descHtml.m_value)
              requirement.Description = filteredDescription
            }
            requirements.push(requirement)
          }
        }
        if (!validateHierarchy(requirements)) {
          requirements = false
          enableButton(params.buttons.sendToSpira)
          //throw hierarchy error and exit function
          displayError("heirarchy", true)
          return
        }
        //if the heirarchy is invalid, clear requirements and throw error
        sendArtifacts(params.artifactEnums.requirements, imageObjects, requirements, projectId, model)
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
        // let tableImages = selectionTables.items[0].getRange().inlinePictures;
        // context.load(tableImages)
        // await context.sync();
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
                //this whole block will be replaced with filterDescription when that is done
                let filteredDescription = filterDescription(descHtml.m_value.replaceAll("\r", ""), true);
                testCase.folderDescription = filteredDescription
              }
              else {
                //removes tables picked up in the description and adds proper HTML lists
                let filteredDescription = filterDescription(descHtml.m_value.replaceAll("\r", ""), true)
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
              let filteredDescription = filterDescription(descHtml.m_value.replaceAll("\r", ""), true)
              testCase.testCaseDescription = filteredDescription
            }
            //dont push a nameless testCase.
            if (testCase.Name) {
              testCases.push(testCase)
            }
          }
        }
        sendArtifacts(params.artifactEnums.testCases, imageObjects, testCases, projectId, model)
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
const sendArtifacts = async (ArtifactTypeId, images, Artifacts, projectId, model) => {
  //this checks if an empty artifact array is passed in (should never happen)
  if (Artifacts.length == 0) {
    //empty is the error message key for the model object.
    displayError("empty", true);
    enableButton(params.buttons.sendToSpira)
    return
  }
  let imgRegex = params.regexs.imageRegex
  showProgressBar();
  switch (ArtifactTypeId) {
    //this is the logic for sending requirements to spira
    case (params.artifactEnums.requirements): {
      //to save refactoring time
      let requirements = Artifacts
      let lastIndent = 0
      const outdentCall = model.user.url + params.apiComponents.apiBase + projectId +
        params.apiComponents.postOrPutRequirement + params.apiComponents.initialOutdent +
        model.user.userCredentials

      try {
        //this call is separate from the rest due to being specially outdented
        let firstCall = await axios.post(outdentCall, requirements[0])

        let placeholders = [...requirements[0].Description.matchAll(imgRegex)]
        for (let i = 0; i < placeholders.length; i++) {
          pushImage(firstCall.data, images[0])
          images.shift();
        }
        updateProgressBar(1, requirements.length)
        //removes the first requirement to save on checking in the for..of function after this
        requirements.shift()
      }
      catch (err) {
        await axios.post(RETRIEVE, { failed1: "req" })
        displayError("failedReq", false, item);
      }
      const apiUrl = model.user.url + params.apiComponents.apiBase + projectId +
        params.apiComponents.postOrPutRequirement + model.user.userCredentials
      for (let [i, req] of requirements.entries()) {
        try {
          let call = await axios.post(apiUrl, req)
          await axios.post(RETRIEVE, call.body)
          await indentRequirement(apiUrl, call.data.RequirementId, req.IndentLevel - lastIndent)
          lastIndent = req.IndentLevel;
          let placeholders = [...req.Description.matchAll(imgRegex)]
          //the 'p' itemization of placeholders isnt needed - just needs to happen once per placeholder
          for (let p of placeholders) {
            await pushImage(call.data, images[0])
            images.shift();
          }
          updateProgressBar(i + 1, requirements.length);
        }
        catch (err) {
          displayError("failedReq", false, req)
        }
      }
      hideProgressBar();
      document.getElementById("send-to-spira-button").disabled = false;
      return
    }
    case (params.artifactEnums.testCases): {
      let testCaseFolders = await retrieveTestCaseFolders(projectId, model);
      let testCases = Artifacts
      for (let [i, testCase] of testCases.entries()) {
        let folder = testCaseFolders.find(folder => folder.Name == testCase.folderName)
        //if the folder doesnt exist, create a new folder on spira
        if (!folder) {
          let newFolder = {}
          newFolder.TestCaseFolderId = await createTestCaseFolder(testCase.folderName,
            testCase.folderDescription, projectId, model);
          newFolder.Name = testCase.folderName;
          folder = newFolder
          testCaseFolders.push(newFolder);
        }
        //this returns the full test case object response
        let testCaseArtifact = await sendTestCase(testCase.Name, testCase.testCaseDescription,
          folder.TestCaseFolderId, projectId, model)
        //imgRegex defined at the top of the sendArtifacts function
        let placeholders = [...testCase.testCaseDescription.matchAll(imgRegex)]
        //p isnt needed but I do need to iterate through the placeholders(this is shorter syntax)
        for (let p of placeholders) {
          await pushImage(testCaseArtifact, images[0])
          images.shift();
        }
        for (let testStep of testCase.testSteps) {
          let step = await pushTestStep(testCaseArtifact.TestCaseId, testStep, model, projectId);
          if (images[0]) {
            await pushImage(step, images[0], testCaseArtifact.TestCaseId);
            images.shift();
          }
        }
        updateProgressBar(i + 1, testCases.length);
      }
      hideProgressBar();
      enableButton(params.buttons.sendToSpira)
    }
  }
}


/* Filters a string and changes any word-outputted lists to properly formatted html lists. */
const filterForLists = (description) => {
  let startRegex = params.regexs.paragraphRegex;
  let elemList = [...description.matchAll(startRegex)];
  description = convertToIndentedList(description, elemList);
  return description
}

/* Scans each element in an array of 'strings' for "style='margin-left:#.0in" where the # is the indent level 
   Then it keeps track of the current indent level as it loops through the array, processing the elements
   through convertToListElem and adding an extra <ul> or <ol> as necessary to properly turn them into html 
   lists. Implement exception for 1.1.1.1 lists */
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
  let exceptedList = false;
  if (exceptedList = params.regexs.exceptedListRegEx.test(listElem)) { //THIS MAY CAUSE BUGS WHEN using numbers 
    return { elem: listElem, ordered: ordered, exceptedList: exceptedList };
  }
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
  return { elem: listElem, ordered: ordered, exceptedList };
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

//this function retrieves an array of all testCaseFolders in the selected project
const retrieveTestCaseFolders = async (projectId, model) => {
  let apiCall = model.user.url + params.apiComponents.apiBase + projectId +
    params.apiComponents.postOrGetTestFolders + model.user.userCredentials;
  try {
    let callResponse = await superagent.get(apiCall).set('accept', "application/json").set('Content-Type', "application/json")
    return callResponse.body
  }
  catch (err) {
    //throw error or do nothing (would be network error maybe?)
  }
}

//This function creates a new test case folder in spira and returns its TestCaseFolderId
const createTestCaseFolder = async (folderName, description, projectId, model) => {
  let apiCall = model.user.url + params.apiComponents.apiBase + projectId +
    params.apiComponents.postOrGetTestFolders + model.user.userCredentials
  try {
    let folderCall = await axios.post(apiCall, {
      Name: folderName,
      Description: description
    })
    return folderCall.data.TestCaseFolderId;
  }
  catch (err) {
    console.log(err);
    return null;
  }
}
//sends a test case to spira
const sendTestCase = async (testCaseName, testCaseDescription, testFolderId, projectId, model) => {
  try {
    let apiCall = model.user.url + params.apiComponents.apiBase + projectId + params.apiComponents.postOrPutTestCase +
      model.user.userCredentials
    var testCaseResponse = await axios.post(apiCall, {
      Name: testCaseName,
      Description: testCaseDescription,
      TestCaseFolderId: testFolderId
    })
    return testCaseResponse.data;
  }
  catch (err) {
    console.log(err);
    return null;
  }
}

//sends a test step to spira
const pushTestStep = async (testCaseId, testStep, model, projectId) => {
  /*pushTestCase should call this passing in the created testCaseId and iterate through passing
  in that test cases test steps.*/
  let apiCall = model.user.url + params.apiComponents.apiBase + projectId +
    params.apiComponents.getTestCase + testCaseId + params.apiComponents.postOrPutTestStep +
    model.user.userCredentials;
  try {
    //testStep = {Description: "", SampleData: "", ExpectedResult: ""}
    //we dont need the response from this - so no assigning to variable.
    let stepCall = await axios.post(apiCall, testStep)
    return stepCall.data
  }
  catch (err) {
    console.log(err)
  }
}
/*filters tables (for test cases/test steps) and body tags out of the description, and
converts lists into a readable format*/
//params:
//description: HTML string that signifies the description block
//isTestCase: 
const filterDescription = (description, isTestCase, lists) => {
  //this function will filter out tables + excess html tags and info from all html based fields
  //this gets the index of the 2 body tags in the HTML string
  let bodyTagmatches = [...description.matchAll(params.regexs.bodyTagRegex)]
  /*this slices the body to be the beginning of the first
   body tag to the beginning of the closing body tag*/
  let htmlBody = description.slice(bodyTagmatches[0].index, bodyTagmatches[1].index)
  //this removes the first body tag, and then trims whitespace
  htmlBody = htmlBody.replace(bodyTagmatches[0][0], "").trim()
  let whitespaceMatch = [...htmlBody.matchAll(params.regexs.whitespaceRegex)]
  for (let match of whitespaceMatch) {
    //this takes any double or more spaces, and turns them into single spaces
    htmlBody = htmlBody.replace(match, " ")
  }
  htmlBody = filterForLists(htmlBody.replaceAll("\r", "").replaceAll("\t", ""))
  //if the description is for a test case, remove tables (They will be parsed as steps 
  //separately)
  if (isTestCase) {
    let tableMatches = [...htmlBody.matchAll(params.regexs.tableRegex)]
    for (let match of tableMatches) {
      /*tableMatches[i][0] is the full match - the second array is the matched groups but
      in this case I do not need the groups, only the full match*/
      htmlBody.replace(match[0], "")
    }
  }
  axios.post(RETRIEVE, { html: htmlBody })
  return htmlBody
}

const filterForListsNew = (description, lists) => {
  let matches = [];
  for (let i = 0; i < lists.length; i++) {
    let list = lists[i];
    let newListElement = "";
    for (let j = 0; j < list.length; j++) {
      let listItem = list[j];
      match = listItem.html.m_value.match(params.regexs.listReplacementRegEx);
      let newItemElement = "";
      
    }
  }
}
export {
  parseArtifacts
}

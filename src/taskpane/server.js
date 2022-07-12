const axios = require('axios')
const superagent = require('superagent');
//ignore it saying defaults doesnt exist, it does and using default does not work.
//makes sure put requests uses the proper content-type header
axios.defaults.headers.put['Content-Type'] = "application/json"
axios.defaults.headers.put['accept'] = "application/json"

export {
  parseArtifacts,
  loginCall,
  retrieveStyles,
  updateSelectionArray
}

import { Data, ERROR_MESSAGES, params, templates } from './model.js'
import {
  disableButton,
  displayError,
  clearErrors,
  showProgressBar,
  updateProgressBar,
  hideProgressBar,
  enableButton,
  enableMainButtons,
  disableMainButtons
} from './taskpane.js';

/*
  Functions that should be moved over: 
  pushImage (i think),
  indentRequirement
*/

/*Functions that have been integrated into parseArtifacts: pushArtifacts,
 updateSelectionArray, retrieveImages, parseRequirements, retrieveTables,
 parseTestCases, parseTestSteps
 */

/*LINES THAT NEED ATTENTION: 121 175 214 */

var RETRIEVE = "http://localhost:5000/retrieve"

//params:
//ArtifactTypeId: Int - ID based on params.artifactEnums.{artifact-type}
//model: Object - user model object based on Data() model. (contains user credentials)
//versionSupport: boolean - expresses if the user supports version 1.3. If not, skips 1.3 functionality

/*This function takes in an ArtifactTypeId and parses the selected text or full body
of a document, returning digestable objects which can easily be used to send artifacts
to spira.*/
const parseArtifacts = async (ArtifactTypeId, model, versionSupport) => {
  return Word.run(async (context) => {
    disableMainButtons();
    /***************************
     Start range identification
     **************************/
    let projectId = document.getElementById("project-select").value
    let selection = context.document.getSelection();
    let splitSelection = context.document.getSelection().paragraphs;
    if (versionSupport) {
      splitSelection = context.document.getSelection().split(['\r'])
    }
    context.load(splitSelection, ['text', 'inlinePictures', 'style', 'styleBuiltIn'])
    context.load(selection, ['text', 'inlinePictures', 'style', 'styleBuiltIn']);
    try {
      await context.sync();
    }
    catch (err) {
      /*there is a crashing error with context.sync(), 
      if your cursor is on an empty line .getSelection() throws error and crashes
      so this handles the fallback and forces full body parsing in this senario.
      May want to add a confirmation for full body parsing and explain
      what can cause this rather than sending the full body without
      user knowledge*/
      selection = {}
      //we need an error for explaining this (and giving the user the option to confirm or cancel)?
      // displayError()
    }
    //If there is no selected area, parse the full body of the document instead
    if (!selection.text) {
      //.getRange is an API v1.3 function. It is only needed for .expandTo calls. paragraphs can handle the rest.
      if (versionSupport) {
        selection = context.document.body.getRange();
        splitSelection = context.document.body.getRange().split(['\r'])
      }
      else {
        selection = context.document.body
        splitSelection = context.document.body.paragraphs;
      }

      context.load(selection, ['text', 'inlinePictures', 'style', 'styleBuiltIn']);
      context.load(splitSelection, ['text', 'inlinePictures', 'style', 'styleBuiltIn'])
      await context.sync();
    }
    /*this verifies that the body has been detected successfully / that it exists. 
    This should never be true here unless a user tries to import a literally empty document
    but I figure the application should at least just tell them thats what it looks like theyre
    trying to do and handle it without crashing the add-in.*/
    //these replaceAll's serve to remove formatting characters that are not actual text
    if (!selection.text.replaceAll("\r", "").replaceAll("&nbsp", "").replaceAll("\t", "").replaceAll("\n", "")) {
      enableMainButtons();
      displayError(ERROR_MESSAGES.empty, false)
      return false
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
    console.log(imageObjects)
    //end of image formatting
    /*************************
     Start list parsing
     *************************/
    //lists are not supported in API version 1.1
    if (versionSupport) {
      let lists = selection.lists
      context.load(lists)
      await context.sync();
      var formattedLists = []
      var formattedSingleItemLists = []
      for (let list of lists.items) {
        context.load(list)
        context.sync();
        // Inner array to store the elements of each list
        let newList = [];
        //set up accessing the paragraphs of a list
        context.load(list, ['paragraphs'])
        await context.sync();
        let paragraphs = list.paragraphs;
        context.load(paragraphs);
        await context.sync();
        for (let paragraph of paragraphs.items) {
          context.load(paragraph)
          await context.sync();
          let listItem = paragraph.listItemOrNullObject
          context.load(listItem, ['level', 'listString'])
          await context.sync();
          /*this covers odd formats with long listStrings (hover listString for details).
          The "false" newList will get pushed, and serve as a "skip this one" flag for 
          the later function which places formatted lists into the description.*/
          if (paragraph.listItemOrNullObject.listString.length > 5 || (newList == false && typeof newList == 'boolean')) {
            newList = false
            continue
          }
          let html = paragraph.getHtml();
          await context.sync();
          let pRegex = params.regexs.paragraphRegex
          let match = html.m_value.match(pRegex)
          html = match[0]
          let spanMatch = html.match(params.regexs.listSpanRegex)
          if (spanMatch) {
            html = html.replace(spanMatch[0], "")
          }
          let pTagMatches = [...html.matchAll(params.regexs.pTagRegex)]
          if (pTagMatches) {
            for (let matchList of pTagMatches) {
              html = html.replace(matchList[0], "")
            }
          }
          if (listItem.listString.match(params.regexs.orderedListRegex)) {
            /*the second level unordered list listString is literally an o (letter O)
            So this prevents that from unintentionally removing an o from the users sentence*/
            html = html.replace(listItem.listString, "")
          }
          newList.push(new templates.ListItem(html, listItem.level))
        }
        //single item lists need to be parsed differently than multi item ones
        if (paragraphs.items.length == 1) {
          formattedSingleItemLists.push(newList)
        }
        else {
          formattedLists.push(newList)
        }
        newList = []
      }
    }
    //end of list parsing
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
              let descRange;
              if (versionSupport) {
                descRange = descStart.expandTo(descEnd)
              }
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
              let filteredDescription = filterDescription(descHtml.m_value, false, formattedLists, formattedSingleItemLists)
              requirement.Description = filteredDescription
              requirements.push(requirement)
              /*gets the requirement of the to be created requirement based
               on its index in the styles array.*/
              let indent = styles.indexOf(item.styleBuiltIn)
              if (indent == -1) {
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
              if (indent == -1) {
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
              let filteredDescription = filterDescription(descHtml.m_value, false, formattedLists, formattedSingleItemLists)
              requirement.Description = filteredDescription
            }
            requirements.push(requirement)
          }
        }
        //if the heirarchy is invalid, clear requirements and throw error
        if (!validateHierarchy(requirements)) {
          requirements = false
          enableMainButtons();
          //throw hierarchy error and exit function
          displayError(ERROR_MESSAGES.hierarchy, false)
          return
        }
        clearErrors();
        await sendArtifacts(params.artifactEnums.requirements, imageObjects, requirements, projectId, model, styles)
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
        /*This is used for "detecting" a table when parsing test cases, while the actual
        text sent to spira is based on the API.*/
        var tables = [];
        for (let i = 0; i < selectionTables.items.length; i++) {
          let table = selectionTables.items[i].values;
          tables.push(table)
        }
        if (!validateTestSteps(tables, styles[2])) {
          //validateTestSteps throws the error, this just fails quietly
          return false
        }
        /*part of this portion isnt DRY, but due to not being able to pass Word 
        objects between functions cant be made into its own function*/
        for (let [i, item] of splitSelection.items.entries()) {
          //this checks if a line is a part of the next table to be parsed.
          //amInTable serves as a flag to parse a table text.
          let amInTable = false;
          try {
            var nextTable = selectionTables.items[tableCounter].getRange();
            context.load(nextTable)
            await context.sync();
            //checks if the item is within the next table
            let inTableCheck = nextTable.intersectWith(item)
            context.load(inTableCheck)
            //this context.sync() is the part that throws the error quietly, without it it crashes
            await context.sync();
            if (inTableCheck && nextTable.text.includes(item.text.trim())) {
              amInTable = true
            }
          }
          catch (err) {
            //do nothing, amInTable should remain false
          }
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
                let filteredDescription = filterDescription(descHtml.m_value.replaceAll("\r", ""), true, formattedLists, formattedSingleItemLists);
                testCase.folderDescription = filteredDescription
              }
              else {
                //removes tables picked up in the description and adds proper HTML lists
                let filteredDescription = filterDescription(descHtml.m_value.replaceAll("\r", ""), true, formattedLists, formattedSingleItemLists)
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
          //tables = [][][] string where [table][row][column]
          //amInTable is a flag set at the beginning of each loop.
          else if (amInTable) {
            //This procs when there is a table and the first description equals item.text
            //testStepTable = 2d array [row][column]
            /**************************
             *start parseTestSteps Logic*
            ***************************/
            let table = selectionTables.items[tableCounter].getRange().getHtml()
            await context.sync()
            /*this checks if the item is within the table being parsed as the .includes method doesnt do a good job of 
            excluding things due to things like single letters existing. A single letter line is very likely
            to be included in a table, but this guarantees that wont disrupt / damage the parsing of test steps
            while also allowing multiple tables to be parsed for 1 test case. */
            let rows = [...table.m_value.matchAll(params.regexs.tableRowRegex)]
            let formattedTable = []
            for (let row of rows) {
              let tableRow = []
              let cells = [...row[0].matchAll(params.regexs.tableDataRegex)]
              for (let cell of cells) {
                tableRow.push(cell[0])
              }
              formattedTable.push(tableRow)
            }
            //formattedStrings should be the outputted 2d array
            let testStepTable = formattedTable
            //table counter lets parseTestSteps know which table is currently being parsed
            let testStep = new templates.TestStep();
            let testSteps = []
            //this is true when the "Header rows?" box is checked
            let headerCheck = document.getElementById("header-check").checked
            if (headerCheck) {
              //if the user says there are header rows, remove the first row of the table being parsed.
              testStepTable.shift();
            }
            //take testStepTable and put into test steps
            for (let row of testStepTable) {
              //skips lines with empty descriptions to prevent pushing empty steps (returns null if no match)
              let emptyStepRegex = params.regexs.emptyParagraphRegex

              if (row[parseInt(styles[2].slice(-1)) - 1].match(emptyStepRegex)) {
                continue
              }
              testStep = { Description: row[parseInt(styles[2].slice(-1)) - 1], ExpectedResult: row[parseInt(styles[3].slice(-1)) - 1], SampleData: row[parseInt(styles[4].slice(-1)) - 1] }
              testSteps.push(testStep)
            }
            testCase.testSteps = [...testCase.testSteps, ...testSteps]
            //removes the table that has been processed from this functions local reference
            tableCounter++
            tables.shift();
            /************************
            *end parseTestSteps Logic*
            ************************/
          }
          else if (!descStart) {
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
              let filteredDescription = filterDescription(descHtml.m_value.replaceAll("\r", ""), true, formattedLists, formattedSingleItemLists)
              testCase.testCaseDescription = filteredDescription
            }
            if (testCase.Name) {
              testCases.push(testCase)
            }
            else if ((styles.includes(item.style) || styles.includes(item.styleBuiltIn)) && itemtext) {
              testCase.Name = itemtext
              testCases.push(testCase)
            }
          }
        }
        await sendArtifacts(params.artifactEnums.testCases, imageObjects, testCases, projectId, model, styles)
        return testCases
      }
    }
  })
}
//params:
//images = [{base64: b64encoded string, name: "", lineNum: ""}]
//ArtifactTypeId based on params.artifactEnums.{artifact-type}
//Artifacts is the array of parsed artifacts ready for sending
//projectId: Int; ID of selected project at time of "send to spira" button press
//model: Data() object; used for user object containing user credentials
/*styles: []string; The user selected styles, needed for determining order of test step
properties inside a table.*/

/*This function takes the already parsed artifacts and images and sends
them to spira*/
const sendArtifacts = async (ArtifactTypeId, images, Artifacts, projectId, model, styles) => {
  //this checks if an empty artifact array is passed in (should never happen)
  if (Artifacts.length == 0) {
    //empty is the error message key for the model object.
    displayError(ERROR_MESSAGES.empty, false);
    enableButton(params.buttonIds.sendToSpira)
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
          await pushImage(firstCall.data, images[0], projectId, model)
          images.shift();
        }
        updateProgressBar(1, requirements.length)
        //removes the first requirement to save on checking in the for..of function after this
        requirements.shift()
      }
      catch (err) {
        displayError(ERROR_MESSAGES.failedReq, false, item);
      }
      const apiUrl = model.user.url + params.apiComponents.apiBase + projectId +
        params.apiComponents.postOrPutRequirement + model.user.userCredentials
      for (let [i, req] of requirements.entries()) {
        try {
          let call = await axios.post(apiUrl, req)
          await indentRequirement(apiUrl, call.data.RequirementId, req.IndentLevel - lastIndent)
          lastIndent = req.IndentLevel;
          let placeholders = [...req.Description.matchAll(imgRegex)]
          //the 'p' itemization of placeholders isnt needed - just needs to happen once per placeholder
          for (let p of placeholders) {
            await pushImage(call.data, images[0], projectId, model)
            images.shift();
          }
          updateProgressBar(i + 2, requirements.length + 1);
        }
        catch (err) {
          displayError(ERROR_MESSAGES.failedReq, false, req)
        }
      }
      updateProgressBar(requirements.length + 1, requirements.length + 1);
      // hideProgressBar();
      document.getElementById("send-to-spira-button").disabled = false;
      return
    }
    case (params.artifactEnums.testCases): {
      let testCaseFolders = await retrieveTestCaseFolders(projectId, model);
      let testCases = Artifacts
      for (let [i, testCase] of testCases.entries()) {
        let folder = testCaseFolders.find(folder => folder.Name == testCase.folderName)
        //if the folder doesnt exist, create a new folder on spira
        if (!folder && testCase.folderName) {
          let newFolder = {}
          newFolder.TestCaseFolderId = await createTestCaseFolder(testCase.folderName,
            testCase.folderDescription, projectId, model);
          if (!newFolder.TestCaseFolderId) {
            displayError(ERROR_MESSAGES.failedReq, false, testCase)
            return
          }
          newFolder.Name = testCase.folderName;
          folder = newFolder
          testCaseFolders.push(newFolder);
        }
        //if the name is empty for the folder, set it as null (root directory)
        else if (!testCase.folderName){
          folder.TestCAseFolderId = null
        }
        //this returns the full test case object response
        let testCaseArtifact = await sendTestCase(testCase.Name, testCase.testCaseDescription,
          folder.TestCaseFolderId, projectId, model)
        if (!testCaseArtifact) {
          displayError(ERROR_MESSAGES.failedReq, false, testCase)
          return
        }
        //imgRegex defined at the top of the sendArtifacts function
        let placeholders = [...testCase.testCaseDescription.matchAll(imgRegex)]
        //p isnt needed but I do need to iterate through the placeholders(this is shorter syntax)
        for (let p of placeholders) {
          try {
            await pushImage(testCaseArtifact, images[0], projectId, model)
          }
          catch (err) {
            console.log(err)
            displayError(ERROR_MESSAGES.failedReq, false, testCase)
          }
          images.shift();
        }
        for (let testStep of testCase.testSteps) {
          let step = await sendTestStep(testCaseArtifact.TestCaseId, testStep, model, projectId);
          if (!step) {
            displayError(ERROR_MESSAGES.failedReq, false, testCase)
            return
          }
          //these are the <img> 'placeholders' for all 3 testStep fields
          let placeholders = [...testStep.Description.matchAll(imgRegex),
          ...testStep.SampleData.matchAll(imgRegex),
          ...testStep.ExpectedResult.matchAll(imgRegex)]
          for (let p of placeholders) {
            if (images[0]) {
              //this handles images for all 3 test step fields. 
              try {
                await pushImage(step, images[0], projectId, model, testCaseArtifact.TestCaseId, styles);
              }
              catch (err) {
                displayError(ERROR_MESSAGES.failedReq, false, testCase)
              }
              images.shift();
            }
          }

        }
        updateProgressBar(i + 1, testCases.length);
      }
      updateProgressBar(testCases.length + 1, testCases.length + 1);
      // hideProgressBar();
      enableMainButtons();
    }
  }
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
    displayError(ERROR_MESSAGES.testCaseFolders, false)
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
const sendTestStep = async (testCaseId, testStep, model, projectId) => {
  /*sendTestCase should call this passing in the created testCaseId and iterate through passing
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
//description: HTML string that signifies the description blockg
//isTestCase: {boolean} says whether it is a test case (for removing tables)
//lists: [][] ListItem (see model)
//singleItemLists: same but for lists that are only 1 item (they are handled differently)
const filterDescription = (description, isTestCase, lists, singleItemLists) => {
  //this function will filter out tables + excess html tags and info from all html based fields
  //this gets the index of the 2 body tags in the HTML string
  let bodyTagMatches = [...description.matchAll(params.regexs.bodyTagRegex)]
  /*this slices the body to be the beginning of the first
   body tag to the beginning of the closing body tag*/
  let htmlBody = description.slice(bodyTagMatches[0].index, bodyTagMatches[1].index)
  //this removes the first body tag, and then trims whitespace
  htmlBody = htmlBody.replace(bodyTagMatches[0][0], "").trim()
  htmlBody = formatDescriptionLists(htmlBody.replaceAll("\r", "").replaceAll("\t", ""), lists, singleItemLists)
  //this trims excess spaces sometimes brought in by lists
  let whitespaceMatch = [...htmlBody.matchAll(params.regexs.whitespaceRegex)]

  for (let match of whitespaceMatch) {
    //this takes any double or more spaces, and turns them into single spaces
    htmlBody = htmlBody.replace(match[0], " ")
  }
  //if the description is for a test case, remove tables (They will be parsed as steps 
  //separately)
  if (isTestCase) {
    let tableMatches = [...htmlBody.matchAll(params.regexs.tableRegex)]
    for (let match of tableMatches) {
      /*tableMatches[i][0] is the full match - the second array is the matched groups but
      in this case I do not need the groups, only the full match*/
      htmlBody = htmlBody.replace(match[0], "")
    }
  }
  return htmlBody
}

/*This function formats description lists to follow HTML structuring instead of 
what the Word API spits out by default*/
//description: string of outputted and mostly formatted HTML
//lists: formatted lists with objects of {text: string, indentLevel: int}
//singleItemLists: the same as lists but for single items (needs to be parsed differently)
const formatDescriptionLists = (description, lists, singleItemLists) => {
  //if there aren't any lists in this description, return the description unaltered
  if (!description.includes("MsoListParagraph")) {
    /*There is sometimes inconsistent spacing between list symbols and the text of a list item,
    so this removes excess whitespace there. Can still be inconsistent after this, but only
    by 1 space instead of by 1-10*/

    return description
  }

  //listStarts is the starting element of every list, single list starts is needed to
  //know how many single lists are parsed here.
  let singleItemListStarts = [...description.matchAll(params.regexs.singleListItemRegex)]
  let multiItemListStarts = [...description.matchAll(params.regexs.firstListItemRegex)]
  let listStarts = [...multiItemListStarts, ...singleItemListStarts]
  let listEnds = [...description.matchAll(params.regexs.lastListItemRegex)]
  for (let [i, start] of listStarts.entries()) {
    /*these have to be exact - lists[i] being undefined means something different than it 
    existing and being defined as false. False serves as the flag to skip parsing this section*/
    if ((lists[i] === false && lists[i] != undefined) || (singleItemLists[i - lists.length] === false && singleItemLists[i - lists.length] != undefined)) {
      let nbspMatches = [...description.matchAll(params.regexs.nonBreakingWhitespaceRegex)]
      for (let match of nbspMatches) {
        description = description.replace(match[0], " ")
      }
      continue
    }
    let replacementArea;
    /*this will handle all the multi item lists, single item ones will be handled slightly
      differently */
    if (i < multiItemListStarts.length) {
      //the opening of the first tag that represents a list
      let startIndex = description.indexOf(start[0]);
      /*this is the index needed to cut off the entire <p></p> grouping for the last item
      in a list */
      let endIndex = description.indexOf(listEnds[i][0]) + listEnds[i][0].length;
      //this is the "list" in HTML returned by the Word javascript API
      replacementArea = description.slice(startIndex, endIndex)
      //if this index exists, it is (most likely) an ordered list
      var targetList = lists[i]
    }
    //this handles single item lists
    else {
      /*i is based on the index within listStarts which is all of the opening elements
      for multi item lists followed by all the single item list elements*/
      var targetList = singleItemLists[i - lists.length]
      //since this is only a single line list, the regex match IS the entire list
      replacementArea = start[0]
    }
    //this determines if a list is ordered or not (passed to listConstructor)
    let orderTest = replacementArea.match(params.regexs.orderedListRegex)?.index
    //(orderTest) here is a conditional as it determines whether a list is considered a ol or ul
    let listHtml = listConstructor((orderTest), targetList)
    description = description.replace(replacementArea, listHtml)
  }
  //Deletes the relevant lists and singleItemLists from their arrays.
  /*easier to do this tacked on than to re-write the core of this function to 
  account for this.*/
  for (let i = 0; i < multiItemListStarts.length; i++) {
    lists.shift();
  }
  for (let i = 0; i < singleItemListStarts.length; i++) {
    singleItemLists.shift();
  }
  return description
}

//constructs lists in proper HTML format and returns them
//List: [] ListItem (see model)
//isOrdered: {boolean} represents whether it is an ordered list or not
const listConstructor = (isOrdered, list) => {
  //provides a boolean value for if a list is an item with a single lists
  let isSingle = (list.length == 1)
  /*this logic determines what tags/regexs to use depending on whether the regex ive used to 
  populate the "isOrdered" parameter thinks it looks like an ordered list 
  (worst case senario, it puts the wrong kind of list)*/
  let openingTag = '<ol>'
  let closingTag = '</ol>'
  let openingRegex = params.regexs.olTagRegex
  let closingRegex = params.regexs.olClosingTagRegex
  if (!isOrdered) {
    openingTag = '<ul>'
    closingTag = '</ul>'
    openingRegex = params.regexs.ulTagRegex
    closingRegex = params.regexs.ulClosingTagRegex
  }
  let listHtml = openingTag
  if (!isSingle) {
    var previousIndent = 0
    //we assume the first item is unindented as spira doesnt support it any other way anyways
    list[0].indentLevel = 0
    for (let listItem of list) {
      if (listItem.indentLevel > previousIndent) {
        //this will only indent at max 1 time as spira doesnt support anything beyond that.
        listHtml = listHtml.concat(openingTag)
      }
      else if (listItem.indentLevel < previousIndent) {
        //this handles any number of levels of outdenting
        let difference = previousIndent - listItem.indentLevel
        for (let i = 0; i < difference; i++) {
          //closes difference number of <ol> nestings 
          listHtml = listHtml.concat(closingTag)
        }
      }
      //this checks if the listItem.text still contains a list symbol (ie 1., a., A.) and removes it
      let symbolTest = listItem.text.match(params.regexs.orderedListSymbolRegex)
      //when no match, symbolTest = null which crashes when accessing [0]
      if (symbolTest) {
        listItem.text = listItem.text.replace(symbolTest[0], "")
      }
      //adds the li element to the "DOM"
      listHtml = listHtml.concat(`<li>${listItem.text}</li>`)
      previousIndent = listItem.indentLevel
    }
    /*after indent levels have been set, this finds how many "open" <ol> tags are left
      and closes them at the end.*/
    let openings = [...listHtml.matchAll(openingRegex)].length
    let closings = [...listHtml.matchAll(closingRegex)].length
    for (let i = 0; i < (openings - closings); i++) {
      listHtml = listHtml.concat(closingTag)
    }
    return listHtml
  }
  //this is the handling of single item lists, multi item lists return above
  let symbolTest = list[0].text.match(params.regexs.orderedListSymbolRegex)
  //symbolTest is null if no match
  if (symbolTest) {
    list[0].text = list[0].text.replace(symbolTest[0], "")
  }
  listHtml = openingTag + `<li>${list[0].text}</li>` + closingTag
  return listHtml
}

const retrieveStyles = (pageTag) => {
  let styles = []
  for (let i = 1; i <= 5; i++) {
    let style = Office.context.document.settings.get(pageTag + 'style' + i.toString());
    //if this is for one of the last 3 test style selectors, choose column1-3 as auto populate settings
    if (!style && pageTag == "test-" && i >= 3) {
      Office.context.document.settings.set(pageTag + 'style' + i.toString(), 'column' + (i - 2).toString())
      style = 'column' + (i - 2).toString()
    }
    //if there isnt an existing setting, populate with headings
    else if (!style) {
      Office.context.document.settings.set(pageTag + 'style' + i.toString(), 'heading' + i.toString())
      style = 'heading' + i.toString();
    }
    styles.push(style)
  }
  return styles
}

const validateHierarchy = (requirements) => {
  //requirements = [{Name: str, Description: str, IndentLevel: int}, ...]
  //the first requirement must always be indent level 0 (level 1 in UI terms)
  if (requirements[0].IndentLevel != 0) {
    return false
  }
  let prevIndent = 0
  for (let i = 0; i < requirements.length; i++) {
    //if there is a jump in indent levels greater than 1, fails validation
    if (requirements[i].IndentLevel > prevIndent + 1) {
      return false
    }
    prevIndent = requirements[i].IndentLevel
  }
  //if the loop goes through without returning false, the hierarchy is valid so returns true
  return true
}

//passes in all relevant tables and description style for test steps (only required field).
const validateTestSteps = (tables, descStyle) => {
  //the column of descriptions according to the style mappings. -1 for indexing against the arrays
  let column = parseInt(descStyle.slice(-1)) - 1
  //tables = [[[], [], []], [[], []], ...] array of 2d arrays(3d array). tables[tableNum][row][column]
  for (let i = 0; i < tables.length; i++) {
    //holder for at least one description for a test step being in a given table
    let atLeastOneDesc = false;
    for (let j = 0; j < tables[i].length; j++) {
      /*if a description for any row of a table within the style mapped column exists, the table
      is considered valid*/
      if (tables[i][j][column]) {
        atLeastOneDesc = true
      }
    }
    //if there is a table containing no test step descriptions, throws an error and stops execution
    if (!atLeastOneDesc) {
      displayError(ERROR_MESSAGES.table, false);
      return false
    }
  }
  clearErrors();
  return true
}

const pushImage = async (Artifact, image, projectId, model, testCaseId, styles) => {
  //image = {base64: "", name: "", lineNum: int}
  /*upload images and build link of image location in spira 
  ({model.user.url}/{projectID}/Attachment/{AttachmentID}.aspx)*/
  //Add AttachmentURL to each imageObject after they are uploaded
  let pid = projectId
  let imgLink;
  let imageApiCall = model.user.url + "/services/v6_0/RestService.svc/projects/"
    + pid + `/documents/file?username=${model.user.username}&api-key=${model.user.api_key}`
  try {
    if (Artifact.RequirementId) {
      let imageCall = await axios.post(imageApiCall, {
        FilenameOrUrl: image.name, BinaryData: image.base64,
        AttachedArtifacts: [{ ArtifactId: Artifact.RequirementId, ArtifactTypeId: 1 }]
      })
      imgLink = model.user.url + `/${pid}/Attachment/${imageCall.data.AttachmentId}.aspx`
    }
    //checks if the artifact is a test step
    else if (Artifact.TestStepId) {
      let imageCall = await axios.post(imageApiCall, {
        FilenameOrUrl: image.name, BinaryData: image.base64,
        AttachedArtifacts: [{ ArtifactId: Artifact.TestStepId, ArtifactTypeId: 7 }]
      })
      imgLink = model.user.url + `/${pid}/Attachment/${imageCall.data.AttachmentId}.aspx`
    }
    //test steps have TestCaseId's, so checks for TestStepId first.
    else if (Artifact.TestCaseId) {
      let imageCall = await axios.post(imageApiCall, {
        FilenameOrUrl: image.name, BinaryData: image.base64,
        AttachedArtifacts: [{ ArtifactId: Artifact.TestCaseId, ArtifactTypeId: 2 }]
      })
      imgLink = model.user.url + `/${pid}/Attachment/${imageCall.data.AttachmentId}.aspx`
    }
  }
  catch (err) {
    console.log(err)
    return
  }
  let fullArtifactObj;
  //checks if the artifact is a requirement (if not it is a test case)
  if (Artifact.RequirementId) {
    try {
      //makes a get request for the target artifact which will be updated to contain an image
      let getArtifact = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
        `/requirements/${Artifact.RequirementId}?username=${model.user.username}&api-key=${model.user.api_key}`;
      let getArtifactCall = await superagent.get(getArtifact).set('accept', 'application/json').set('Content-Type', 'application/json');
      //This is the body of the get response in its entirety.
      fullArtifactObj = getArtifactCall.body
    }
    catch (err) {
      console.log(err)
      return
    }
    // now replace the placeholder in the description with img tags
    let placeholderRegex = params.regexs.imageRegex
    //gets an array of all the placeholders for images. 
    let placeholders = [...fullArtifactObj.Description.matchAll(placeholderRegex)]
    /*placeholders[0][0] is the first matched instance - because you need to GET for each change
    this should work each time - each placeholder should have 1 equivalent image in the same
    order they appear throughout the document.*/
    fullArtifactObj.Description = fullArtifactObj.Description.replace(placeholders[0][0], `<img alt=${image?.name} src=${imgLink}><br />`)
    //PUT artifact with new description (including img tags now)
    let putArtifact = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
      `/requirements?username=${model.user.username}&api-key=${model.user.api_key}`;
    try {
      await axios.put(putArtifact, fullArtifactObj)
    }
    catch (err) {
      //do nothing
      console.log(err)
      return
    }
  }
  else if (Artifact.TestStepId) {
    try {
      //this needs to have a different way of getting TestCaseId
      //makes a get request for the target artifact which will be updated to contain an image
      let getArtifact = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
        `/test-cases/${testCaseId}/test-steps/${Artifact.TestStepId}?username=${model.user.username}&api-key=${model.user.api_key}`;
      let getArtifactCall = await superagent.get(getArtifact).set('accept', 'application/json').set('Content-Type', 'application/json');
      //This is the body of the get response in its entirety.
      fullArtifactObj = getArtifactCall.body
    }
    catch (err) {
      //do nothing
      console.log(err)
      return
    }
    // now replace the placeholder in the description with img tags
    let placeholderRegex = params.regexs.imageRegex

    //This is needed to determine the order of how images will appear throughout
    //the document relative to the TestStep field they represent.
    let styleOrganizer = [{ column: parseInt(styles[2].substring(styles[2].length - 1)), for: "descPlaceholders" },
    { column: parseInt(styles[3].substring(styles[2].length - 1)), for: "expectedPlaceholders" },
    { column: parseInt(styles[4].substring(styles[2].length - 1)), for: "samplePlaceholders" }]

    //this sorts the styles by increasing order of column number (lowest column first)
    styleOrganizer.sort((a, b) => (a.column > b.column) ? 1 : -1)
    let descPlaceholders = [...fullArtifactObj.Description.matchAll(placeholderRegex)]
    let samplePlaceholders = [...fullArtifactObj.SampleData.matchAll(placeholderRegex)]
    let expectedPlaceholders = [...fullArtifactObj.ExpectedResult.matchAll(placeholderRegex)]

    /*this allows us to directly reference any placeholder array with the
     styleOrganizer "for" property*/
    let placeholderReference = {
      descPlaceholders: descPlaceholders,
      samplePlaceholders: samplePlaceholders, expectedPlaceholders: expectedPlaceholders
    }

    /*placeholders[0][0] is the first matched instance - because you need to GET for 
    each changethis should work each time - each placeholder should have 1 equivalent
    image in the same order they appear throughout the document.*/
    for (let placeholder of styleOrganizer) {
      //placeholder = object from styleOrganizer
      if (placeholderReference[placeholder.for].length != 0) {
        /*this changes the relevant behavior based on whichever property is both the
        most to the left in the testSteps tables and also has an image placeholder */
        switch (placeholder.for) {
          case ("descPlaceholders"): {
            fullArtifactObj.Description = fullArtifactObj.Description.replace(descPlaceholders[0][0], `<img src=${imgLink} alt=${image.name} />`)
            break
          }
          case ("samplePlaceholders"): {
            fullArtifactObj.SampleData = fullArtifactObj.SampleData.replace(samplePlaceholders[0][0], `<img src=${imgLink} alt=${image.name} />`)
            break
          }
          case ("expectedPlaceholders"): {
            fullArtifactObj.ExpectedResult = fullArtifactObj.ExpectedResult.replace(expectedPlaceholders[0][0], `<img src=${imgLink} alt=${image.name} />`)
            break
          }
        }
        break
      }
      continue
    }
    //PUT artifact with new description (including img tags now)
    let putArtifact = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
      `/test-cases/${fullArtifactObj.TestCaseId}/test-steps?username=${model.user.username}&api-key=${model.user.api_key}`;
    try {
      await axios.put(putArtifact, fullArtifactObj)
    }
    catch (err) {
      //do nothing
      console.log(err)
      return
    }
  }
  else if (Artifact.TestCaseId) {
    try {
      //handle test cases
      //makes a get request for the target artifact which will be updated to contain an image
      let getArtifact = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
        `/test-cases/${Artifact.TestCaseId}?username=${model.user.username}&api-key=${model.user.api_key}`;
      let getArtifactCall = await superagent.get(getArtifact).set('accept', 'application/json').set('Content-Type', 'application/json');
      //This is the body of the get response in its entirety.
      fullArtifactObj = getArtifactCall.body
    }
    catch (err) {
      //do nothing
      console.log(err)
      return
    }
    //crashing HERE
    //dsajdhsjkda
    //dhasdjkl
    // now replace the placeholder in the description with img tags
    let placeholderRegex = /<img(.|\n|\r)*("|\s)\>/g
    //gets an array of all the placeholders for images. 
    let placeholders = [...fullArtifactObj.Description.matchAll(placeholderRegex)]
    /*placeholders[0][0] is the first matched instance - because you need to GET for each change
    this should work each time - each placeholder should have 1 equivalent image in the same
    order they appear throughout the document.*/
    fullArtifactObj.Description = fullArtifactObj.Description.replace(placeholders[0][0], `<img alt=${image?.name} src=${imgLink}><br />`)
    //PUT artifact with new description (including img tags now)
    let putArtifact = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
      `/test-cases?username=${model.user.username}&api-key=${model.user.api_key}`;
    try {
      await axios.put(putArtifact, fullArtifactObj)
    }
    catch (err) {
      //do nothing
      console.log(err)
      return false
    }
  }
  return true
}

/*indents requirements to the appropriate level, relative to the last requirement in the product
before this add-on begins to add more. (No way to find out indent level of the last requirement
  in a product from the Spira API (i think))*/
const indentRequirement = async (apiCall, id, indent) => {
  apiCall = apiCall.replace("requirements", `requirements/${id}/${indent > 0 ? "indent" : "outdent"}`)
  indent = Math.abs(indent);
  //loop for indenting/outdenting requirement
  for (let i = 0; i < indent; i++) {
    try {
      await axios.post(apiCall, {});
    }
    catch (err) {
      //do nothing - error displayed from catch block this function is called from
      console.log(err)
      return false
    }
  }
  return true
}

const loginCall = async (apiUrl) => {
  var response = await superagent.get(apiUrl).set('accept', 'application/json').set("Content-Type", "application/json")
  return response
}

async function updateSelectionArray() {
  return Word.run(async (context) => {
    //check for highlighted text  
    //splits the selected areas by enter-based indentation. 
    let selection = context.document.body
    context.load(selection, 'text');
    await context.sync();
    selection = context.document.body.paragraphs
    //loads the text, style elements, and any images from a given line
    context.load(selection, ['text', 'styleBuiltIn', 'style'])
    await context.sync();
    // Testing parsing lines of text from the selection array and logging it
    let lines = []
    selection.items.forEach((item) => {
      lines.push({
        text: item.text, style: (item.styleBuiltIn == "Other" ? item.style : item.styleBuiltIn),
        custom: (item.styleBuiltIn == "Other")
      })
    })
    return lines
  })
}

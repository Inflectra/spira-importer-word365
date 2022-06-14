/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

/***********************
Initialization Functions
***********************/

const axios = require('axios')
const superagent = require('superagent');
//adding content-type header as it is required for spira api v6_0
//ignore it saying defaults doesnt exist, it does and using default does not work.

// Global selection array, used throughout
/*This is a global variable because the word API call functions are unable to return
values from within due to the required syntax of returning a Word.run((callback) =>{}) 
function. */
var SELECTION = [];
//setting a user object to maintain credentials when using other parts of the add-in
var USER_OBJ = { url: "", username: "", password: "" }
var RETRIEVE = "http://localhost:5000/retrieve"

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    setDefaultDisplay();
    setEventListeners();
    document.body.classList.add('ms-office');
    // this element doesn't currently exist
    // document.getElementById("help-connection-google").style.display = "none";
  }
});

const setDefaultDisplay = () => {
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("req-style-mappings").style.display = 'none';
  document.getElementById("test-style-mappings").style.display = 'none';
  document.getElementById("empty-err").style.display = 'none';
  document.getElementById("failed-req-err").style.display = 'none'
}

const setEventListeners = () => {
  // document.getElementById('test').onclick = test;
  document.getElementById('btn-login').onclick = () => loginAttempt();
  document.getElementById('dev-mode').onclick = () => devmode();
  // document.getElementById('send-artifacts').onclick = () => pushArtifacts();
  document.getElementById('log-out').onclick = () => logout();
  //I think theres a way to use classes to reduce this to 2 but unsure
  document.getElementById("confirm-req-style-mappings").onclick = () => confirmStyleMappings(true, 'req-');
  document.getElementById("confirm-test-style-mappings").onclick = () => confirmStyleMappings(true, 'test-');
  document.getElementById("select-requirements").onclick = () => openStyleMappings("req-");
  document.getElementById("select-test-cases").onclick = () => openStyleMappings("test-")
}

const devmode = () => {
  //moves us to the main interface without manually entering credentials
  document.getElementById('panel-auth').classList.add('hidden');
  document.getElementById('main-screen').classList.remove('hidden');
  document.getElementById("main-screen").style.display = "flex"
}


/****************
Testing Functions 
*****************/
//basic testing function for validating code snippet behaviour.
export async function test() {

  return Word.run(async (context) => {
    /*this is the syntax for accessing all tables. table text ends with \t when retrieved from
     the existing fucntion (load(context.document.getSelection(), 'text'))and this can be 
     utilized in order to identify when you have entered a table, at which point we parse
     information out of the table using that method to know the structure (returns 2d array) */
    await context.sync();
    let tables = await retrieveTables()
    await axios.post("http://localhost:5000/retrieve", { Tables: tables })
    //try catch block for backend node call to prevent errors crashing the application
    // try {
    //   let call1 = await axios.post("http://localhost:5000/retrieve", { lines: lines })
    // }
    // catch (err) {
    //   console.log(err)
    // }
    // Tests the parseRequirements Function
    // let requirements = parseRequirements(lines);

    //try catch block for backend node call to prevent errors crashing the application
    // try {
    //   let call1 = await axios.post("http://localhost:5000/retrieve", { lines: lines, headings: requirements })
    // }
    // catch (err) {
    //   console.log(err)
    // }
  })
}

/**************
Spira API calls
**************/

const loginAttempt = async () => {
  /*disable the login button to prevent someone from pressing it multiple times, this can
  overpopulate the products selector with duplicate sets.*/
  document.getElementById("btn-login").disabled = true
  //retrieves form data from input elements
  let url = document.getElementById("input-url").value
  let username = document.getElementById("input-username").value
  let rssToken = document.getElementById("input-password").value
  //allows user to enter URL with trailing slash or not.
  let apiBase = "/services/v6_0/RestService.svc/projects"
  if (url[url.length - 1] == "/") {
    //url cannot be changed as it is tied to the HTML DOM input object, so creates a new variable
    var finalUrl = url.substring(0, url.length - 1)
  }
  //formatting the URL as it should be to populate products / validate user credentials
  let validatingURL = finalUrl || url + apiBase + `?username=${username}&api-key=${rssToken}`;
  try {
    //call the products API to populate relevant products
    var response = await superagent.get(validatingURL).set('accept', 'application/json').set("Content-Type", "application/json")
    if (response.body) {
      //if successful response, move user to main screen
      document.getElementById('panel-auth').classList.add('hidden');
      document.getElementById('main-screen').classList.remove('hidden');
      document.getElementById('main-screen').style.display = 'flex';
      document.getElementById("btn-login").disabled = false
      //save user credentials in global object to use in future requests
      USER_OBJ = {
        url: finalUrl || url, username: username, password: rssToken
      }
      //populate the products dropdown with the response body.
      populateProjects(response.body)
      //On successful login, hide error message if its visible
      document.getElementById("login-err-message").classList.add('hidden')
      return
    }
  }
  catch (err) {
    //if the response throws an error, show an error message
    document.getElementById("login-err-message").classList.remove('hidden');
    document.getElementById("btn-login").disabled = false;
    return
  }
}

// Send a requirement to Spira using the requirements API
const pushRequirements = async () => {
  await updateSelectionArray();
  // Tests the parseRequirements Function
  let requirements = parseRequirements(SELECTION);
  let lastIndent = 0;
  /*if someone has selected an area with no properly formatted text, show an error explaining
  that and then return this function to prevent sending an empty request.*/
  if (requirements.length == 0) {
    document.getElementById("empty-err").style.display = 'flex';
    setTimeout(() => {
      document.getElementById('empty-err').style.display = 'none';
    }, 8000)
    return
  }
  let id = document.getElementById('project-select').value;
  let item = requirements[0]
  const outdentCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + id +
    `/requirements/indent/-20?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
  try {
    let call = await axios.post(outdentCall, { Name: item.Name, Description: item.Description, RequirementTypeId: 2 });
    await axios.post(RETRIEVE, { call: call.data })
  }
  catch (err) {
    await axios.post(RETRIEVE, { err: err });
    /*shows the requirement which failed to add. This should work if it fails in the middle of 
    sending a set of requirements*/
    document.getElementById("failed-req-error").textContent = `The request to the API has failed on requirement: '${item.Name}'. All, if any previous requirements should be in Spira.`
    document.getElementById("failed-req-error").style.display = "flex";
    setTimeout(() => {
      document.getElementById('failed-req-error').style.display = 'none';
    }, 8000)
  }
  for (let i = 1; i < requirements.length; i++) {
    let item = requirements[i];
    const apiCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + id +
      `/requirements?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
    // try catch block to stop application crashing and show error message if call fails
    try {
      let call = await axios.post(apiCall, { Name: item.Name, Description: item.Description, RequirementTypeId: 2 });
      await indentRequirement(apiCall, call.data.RequirementId, item.IndentLevel - lastIndent)
      lastIndent = item.IndentLevel;
    }
    catch (err) {
      /*shows the requirement which failed to add. This should work if it fails in the middle of 
      sending a set of requirements*/
      document.getElementById("failed-req-error").textContent = `The request to the API has failed on requirement: '${item.Name}'. All, if any previous requirements should be in Spira.`
      document.getElementById("failed-req-error").style.display = "flex";
      setTimeout(() => {
        document.getElementById('failed-req-error').style.display = 'none';
      }, 8000)
    }
  }
  return
}

/*indents requirements to the appropriate level, relative to the last requirement in the product
before this add-on begins to add more. (No way to find out indent level of the last requirement
  in a product from the Spira API (i think))*/
const indentRequirement = async (apiCall, id, indent) => {
  if (indent > 0) {
    //loop for indenting requirement
    for (let i = 0; i < indent; i++) {
      try {
        let call2 = await axios.post(apiCall.replace("requirements", `requirements/${id}/indent`), {});
      }
      catch (err) {
        console.log(err)
      }
    }
  }
  else {
    //loop for outdenting requirement
    for (let i = 0; i > indent; i--) {
      try {
        let call2 = await axios.post(apiCall.replace("requirements", `requirements/${id}/outdent`), {});
      }
      catch (err) {
        console.log(err)
      }
    }
  }
}

/* 
  Sends all of the test case folders and test cases found in the selection to the Spira instance
  WIP Until parseTestSteps is fully implemented
*/
const pushTestCases = async () => {
  await updateSelectionArray();
  // testCases = {folderName: "", Name: "", testSteps: [{Description, expected result, sample data}, ...]}
  let testCases = await parseTestCases(SELECTION);
  // testCaseFolders = [{Name: "", TestCaseFolderId: int}, ...]
  let testCaseFolders = await retrieveTestCaseFolders();
  for (let i = 0; i < testCases.length; i++) {
    let testCase = testCases[i];
    await axios.post(RETRIEVE, { here: "heer" })
    // First check if it's in an existing folder
    let folder = testCaseFolders.find(folder => folder.Name == testCase.folderName)

    if (!folder) { // If the folder doesn't exist yet, make it and then make the 
      let newFolder = {}
      newFolder.TestCaseFolderId = await pushTestCaseFolder(testCase.folderName, testCase.folderDescription);
      newFolder.folderName = testCase.folderName;
      await axios.post(RETRIEVE, { newfolder: newFolder })
      folder = newFolder
      testCaseFolders.push(newFolder);
    }
    // make the testCase and keep the Id for later
    let testCaseId = await pushTestCase(testCase.Name, testCase.testCaseDescription, folder.TestCaseFolderId);
    // now make the testSteps
    for (let j = 0; j < testCase.testSteps.length; j++) {
      await pushTestStep(testCaseId, testCase.testSteps[j]);
    }
  }


  // CURRENTLY USED FOR TESTING
  // let folderResponse = await pushTestCaseFolder("Test Folder", "First Functional Folder Test");
  // let testCaseResponse = await pushTestCase("test case", folderResponse)
  // try {
  //   let call1 = await axios.post("http://localhost:5000/retrieve", { Folder: folderResponse, TestCase: testCaseResponse })
  // }
  // catch (err) {
  //   console.log(err);
  // }
}
const retrieveTestCaseFolders = async () => {
  let projectId = document.getElementById('project-select').value;
  let apiCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + projectId +
    `/test-folders?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
  let callResponse = await superagent.get(apiCall).set('accept', "application/json").set('Content-Type', "application/json")
  return callResponse.body
}

const pushTestStep = async (testCaseId, testStep) => {
  /*pushTestCase should call this passing in the created testCaseId and iterate through passing
  in that test cases test steps.*/
  let projectId = document.getElementById('project-select').value;
  await axios.post(RETRIEVE)
  let apiCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + projectId +
    `/test-cases/${testCaseId}/test-steps?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
  try {
    //testStep = {Description: "", SampleData: "", ExpectedResult: ""}
    //we dont need the response from this - so no assigning to variable.
    await axios.post(RETRIEVE, { api: apiCall, description: testStep.Description, smpl: testStep.SampleData, exp: testStep.ExpectedResult })
    await axios.post(apiCall, {
      Description: testStep.Description,
      SampleData: testStep.SampleData,
      ExpectedResult: testStep.ExpectedResult
    })
  }
  catch (err) {
    console.log(err)
    await axios.post(RETRIEVE, { err: err })
  }
}
/* 
  Creates a test case using the information given and sends it to the Spira instance. Returns the Id of the created test case
*/
const pushTestCase = async (testCaseName, testCaseDescription, testFolderId) => {
  try {
    var testCaseResponse = await axios.post(`${USER_OBJ.url}/services/v6_0/RestService.svc/projects/24/test-cases?username=${USER_OBJ.username}
      &api-key=${USER_OBJ.password}`, {
      Name: testCaseName,
      Description: testCaseDescription,
      TestCaseFolderId: testFolderId
    })
    return testCaseResponse.data.TestCaseId;
  }
  catch (err) {
    console.log(err);
    return null;
  }
}
/*  
  Creates a test folder and returns the Test Folder Id
*/
const pushTestCaseFolder = async (folderName, description) => {
  let projectId = document.getElementById('project-select').value;
  let apiCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + projectId +
    `/test-folders?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
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


/******************** 
HTML DOM Manipulation
********************/

const populateProjects = (projects) => {
  let dropdown = document.getElementById('project-select')
  projects.forEach((project) => {
    /*creates an option for each product which displays the name
     and has a value of its ProjectId for use in API calls*/
    let option = document.createElement("option");
    option.text = project.Name
    option.value = project.ProjectId
    dropdown.add(option)
  })
  return
}

const logout = () => {
  USER_OBJ = { url: "", username: "", password: "" };
  document.getElementById('main-screen').classList.add('hidden');
  //display: flex is set after hidden is removed, may want to make this only use style.display
  document.getElementById('main-screen').style.display = "none";
  //removes currently entered RSS token to prevent a user from leaving their login credentials
  //populated after logging out and leaving their computer.
  document.getElementById("input-password").value = ""
  document.getElementById('panel-auth').classList.remove('hidden');
  clearDropdownElement('project-select');
}

const openStyleMappings = async (pageTag) => {
  //opens the requirements style mappings if requirements is the selected artifact type
  /*all id's and internal word settings are now set using a "pageTag". This allows code 
  to be re-used between testing and requirement style settings. The tags are req- for
  requirements and test- for test cases.*/
  //checks the current selected artifact type then loads the appropriate menu
  if (pageTag == "req-") {
    document.getElementById("req-style-mappings").classList.remove("hidden")
    document.getElementById("test-style-mappings").style.display = 'none'
    document.getElementById("req-style-mappings").style.display = 'flex'
    //populates all 5 style mapping boxes
  }
  //opens the test cases style mappings if test mappings is the selected artifact type
  else {
    document.getElementById("req-style-mappings").style.display = 'none'
    document.getElementById("test-style-mappings").style.display = 'flex'
  }
  //wont populate styles for requirements if it is already populated
  if(document.getElementById("req-style-select1").childElementCount && pageTag == "req-"){
    return
  }
  //wont populate styles for test-cases if it is already populated
  else if(document.getElementById("test-style-select1").childElementCount && pageTag == "test-"){
    return
  }
  //doesnt populate styles if both test & req style selectors are populated
  else if (document.getElementById("test-style-select1").childElementCount && document.getElementById("req-style-select1").childElementCount){
    return;
  }
  //retrieveStyles gets the document's settings for the style mappings. Also auto sets default values
  let settings = retrieveStyles(pageTag)
  //Goes line by line and retrieves any custom styles the user may have used.
  let styles = await getStyles();
  //only the top 2 select objects should have all styles. bottom 3 are table based (at least for now).
  if (pageTag == "test-") {
    for (let i = 1; i <= 2; i++) {
      populateStyles(styles, pageTag + 'style-select' + i.toString());
    }
    //bottom 3 selectors will be related to tables
    for (let i = 3; i <= 5; i++) {
      let tableStyles = ["column1", "column2", "column3", "column4", "column5"]
      populateStyles(tableStyles, pageTag + 'style-select' + i.toString())
    }
  }
  else {
    for (let i = 1; i <= 5; i++) {
      populateStyles(styles, pageTag + 'style-select' + i.toString());
    }
  }
  //move selectors to the relevant option
  settings.forEach((setting, i) => {
    document.getElementById(pageTag + "style-select" + (i + 1).toString()).value = setting
  })
}

//closes the style mapping page taking in a boolean 'result'
//pageTag is req or test depending on which page is currently open

const confirmStyleMappings = (result, pageTag) => {
  //result = true when a user selects confirm to exit a style mappings page
  if (result) {
    //saves the users style preferences. this is document bound
    for (let i = 1; i <= 5; i++) {
      let setting = document.getElementById(pageTag + "style-select" + i.toString()).value
      Office.context.document.settings.set(pageTag + 'style' + i.toString(), setting);
    }
    //this saves the settings
    Office.context.document.settings.saveAsync()
  }
  //clears dropdowns to prevent being populated with duplicate options upon re-opening
  for (let i = 1; i <= 5; i++) {
    clearDropdownElement('req-style-select' + i.toString());
    clearDropdownElement('test-style-select' + i.toString());
  }
}

//Populates a passed in style-selector with the avaiable word styles
const populateStyles = (styles, element_id) => {
  let dropdown = document.getElementById(element_id)
  styles.forEach((style) => {
    /* Creates an option for each style available */
    let option = document.createElement("option");
    option.text = style
    option.value = style
    dropdown.add(option);
  })
}

const clearDropdownElement = (element_id) => {
  let dropdown = document.getElementById(element_id);
  while (dropdown.length > 0) {
    dropdown.remove(0);
  }
}

/********************
Word/Office API calls
********************/

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

/* Get an Array of {text, style} objects from the user's selected text, delimited by /r
 (/r is the plaintext version of a new line started by enter)*/
export async function updateSelectionArray() {
  return Word.run(async (context) => {
    //check for highlighted text
    //splits the selected areas by enter-based indentation. 
    let selection = context.document.getSelection();
    context.load(selection, 'text');
    await context.sync();
    if (selection.text) {
      selection = context.document.getSelection().split(['/r'])
      context.load(selection, ['text', 'styleBuiltIn', 'style'])
      await context.sync();
    }

    // if nothing is selected, select the entire body of the document
    else {
      selection = context.document.body.getRange().split(['/r']);
      context.load(selection, ['text', 'styleBuiltIn', 'style'])
      await context.sync();
    }
    // Testing parsing lines of text from the selection array and logging it
    let lines = []
    selection.items.forEach((item) => {
      lines.push({
        text: item.text, style: (item.styleBuiltIn == "Other" ? item.style : item.styleBuiltIn),
        custom: (item.styleBuiltIn == "Other")
      })
    })
    SELECTION = lines;
  })
}

/* Gets an array of all the tables from the Word document and returns it. */
const retrieveTables = async () => {
  return Word.run(async (context) => {
    let selection = context.document.getSelection().tables;
    context.load(selection);
    await context.sync();
    let tables = [];
    for (let i = 0; i < selection.items.length; i++) {
      let table = selection.items[i].values;
      tables.push(table);
    }
    return tables;
  })
}
/*********************
Pure data manipulation
**********************/

const pushArtifacts = async () => {
  let artifacts = document.getElementById("artifact-select").value;
  if (artifacts == "requirements") {
    await pushRequirements();
  }
  else {
    await pushTestCases();
  }
}

// Parses an array of range objects based on style and turns them into requirement objects
const parseRequirements = (lines) => {
  let requirements = []
  let styles = retrieveStyles('req-')
  lines.forEach((line) => {
    //removes the indentation tags from the text
    line.text = line.text.replaceAll("\t", "").replaceAll("\r", "")
    let requirement = {};
    // TODO: refactor to use for loop where IndentLevel = styles index rather than a switch statement.
    switch (line.style) {
      case "normal":
        //only executes if there is a requirement to add the description to.
        if (requirements.length > 0) {
          //if it is description text, add it to Description of the previously added item in requirements. This allows multi line descriptions
          requirements[requirements.length - 1].Description = requirements[requirements.length - 1].Description + ' ' + line.text
        }
        break
      //Uses the file styles settings to populate into this function. If none set, uses heading1-5
      case styles[0]: {
        requirement = { Name: line.text, IndentLevel: 0, Description: "" }
        requirements.push(requirement)
        break
      }
      case styles[1]: {
        requirement = { Name: line.text, IndentLevel: 1, Description: "" }
        requirements.push(requirement)
        break
      }
      case styles[2]: {
        requirement = { Name: line.text, IndentLevel: 2, Description: "" }
        requirements.push(requirement)
        break
      }
      case styles[3]: {
        requirement = { Name: line.text, IndentLevel: 3, Description: "" }
        requirements.push(requirement)
        break
      }
      case styles[4]:
        requirement = { Name: line.text, IndentLevel: 4, Description: "" }
        requirements.push(requirement)
        break
      //lines not stylized normal or concurrent with style mappings are discarded.
      default: break
    }
    /*if a requirement is populated with an empty name (happens when a line has a style but 
    no text), remove it from the requirements before moving to the next line*/
    if (requirement.Name == "") {
      requirements.pop();
    }
  })
  return requirements
}

const parseTestCases = async (lines) => {
  let testCases = []
  //styles = ['style1', 'style2', columnStyle, columnStyle, columnStyle]
  let styles = retrieveStyles("test-")

  let testCase = { folderName: "", folderDescription: "", Name: "", testCaseDescription: "", testSteps: [] }
  //tables = [[test case 1 steps], [test case 2 steps], ...]
  //word API functions cannot return normally.
  let tables = await retrieveTables()
  await axios.post(RETRIEVE, { talbes: tables })
  for (let i = 0; i < lines.length; i++) {
    //removes enter indent tags
    lines[i].text = lines[i].text.replaceAll("\r", "")
    /*line text ends with a \t for each field contained in a table. This can also be done
    accidentally by the user so that is why the second conditional checks if the text equals
    the first line of the table. If it doesnt, we will just remove the \t and look for 
    relevant styles.*/

    /*tables[0][0] is the first table, first row, and then the last indexing element
    takes the column number out of the style mapping for the description of a test step.
    This is then tested against the line text to see if it is in fact the first element in a 
    table. Then, the line from the table adds back in the \t tag as it is removed when you
    retrieve tables from the word API but is still in the line.text from retrieving all lines.
    */

    //this checks if a line is the first box in the next table
    if (tables[0] && lines[i].text.slice(-1) == "\t" && lines[i].text == tables[0][0][parseInt(styles[2].slice(-1))].concat("\t")) {
      let testSteps = parseTable(tables[0])
      //allows multiple tables to populate the same test case in the case it is multiple tables
      testCase = { ...testCase, testSteps: [...testSteps] }
      /*removes the table which has just been parsed so we dont need to iterate through tables
      in the conditional. */
      tables.shift();
    }
    //this handles whether a line is a folder name or test case name
    switch (lines[i].style) {
      case styles[0]:
        testCase.folderName = lines[i].text
        break
      case styles[1]:
        testCase.Name = lines[i].text
        break
      case "Normal" || "normal":
        if (lines[i - 1].style == styles[0]) {
          testCase.folderDescription = lines[i].text
          break
        }
        else if (lines[i - 1].style == styles[1]) {
          testCase.testCaseDescription = lines[i].text
        }
        break
      default:
        //do nothing
        break

    }
    //if the relevant fields are populated, push the test case and reset the testCase variable
    //3rd conditional checks that the next element is not (likely) a table
    // await axios.post(RETRIEVE, {conditions: testCase.folderName, condition2: testCase.Name, condition3: tables.length})
    if (testCase.folderName && testCase.Name && !tables.length) {
      await axios.post(RETRIEVE, { testCase: testCase })
      testCases.push(testCase)
      testCase = { Name: "", testCaseDescription: "", folderDescription: "", folderName: "", testSteps: [] }
    }
  }
  await axios.post(RETRIEVE, { testCases: testCases })
  return testCases
}

// Updates selection array and then loops through it and adds any
// user-created styles found to its array and returns it. WIP

const parseTable = (table) => {
  let styles = retrieveStyles('test-')
  let testSteps = []
  //columnNums = column numbers for [description, expected result, sample data]
  let columnNums = [parseInt(styles[2].slice(-1)) - 1, parseInt(styles[3].slice(-1)) - 1,
  parseInt(styles[4].slice(-1)) - 1]
  //row = [column1, column 2, column3, ...]
  table.forEach((row) => {
    let testStep = { Description: "", ExpectedResult: "", SampleData: "" }
    //populates fields based on styles, doesnt populate if no description
    if (row[columnNums[0]] != "") {
      testStep.Description = row[columnNums[0]]
      testStep.ExpectedResult = row[columnNums[1]]
      testStep.SampleData = row[columnNums[2]]
      testSteps.push(testStep)
    }
    //pushes it to the testSteps array
  })
  return testSteps
  //return an array of testStep objects
}

const scanForCustomStyles = async () => {
  let customStyles = [];
  await updateSelectionArray();
  for (let i = 0; i < SELECTION.length; i++) {
    if (SELECTION[i].custom && !customStyles.includes(SELECTION[i].style)) {
      customStyles.push(SELECTION[i].style);
    }
  }
  return customStyles;
}

/* Returns an array with all of the styles intended to be used for the 
  Style dropdown menus */
const getStyles = async () => {
  let userStyles = await usedStyles();
  let allStyles = await trimStyles(Object.values(Word.Style), userStyles)
  return allStyles;
}

/* Returns a new array by filtering out styles from the first set and adding them to the second set*/
const trimStyles = async (styles, prevStyles) => {
  let newStyles = prevStyles;
  await axios.post(RETRIEVE, { trimming: "styles" });
  for (let i = 0; i < styles.length; i++) {
    let style = styles[i];
    // only add the style to the new set if none of these are in the string
    if (style.indexOf("Toc") < 0 && style.indexOf("Table") < 0 && style.indexOf("Other") < 0 && style.indexOf("Normal") < 0) {
      //make sure there are no repeats
      if (!newStyles.includes(style)) {
        newStyles.push(style)
      }
    }
  }
  return newStyles;
}

/* Returns the array of all used styles in the selection */
const usedStyles = async () => {
  let styles = [];
  await updateSelectionArray();
  for (let i = 0; i < SELECTION.length; i++) {
    //normal is a reserved style for descriptions of requirements
    if (!styles.includes(SELECTION[i].style) && SELECTION[i].style != "Normal") {
      styles.push(SELECTION[i].style);
    }
  }
  return styles;
}
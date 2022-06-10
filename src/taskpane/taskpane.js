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
  document.getElementById('send-artifacts').onclick = () => test();
  document.getElementById('log-out').onclick = () => logout();
  document.getElementById("style-mappings-button").onclick = () => openStyleMappings();
  //I think theres a way to use classes to reduce this to 2 but unsure
  document.getElementById("confirm-req-style-mappings").onclick = () => closeStyleMappings(true, 'req-');
  document.getElementById("cancel-req-style-mappings").onclick = () => closeStyleMappings(false, 'req-');
  document.getElementById("confirm-test-style-mappings").onclick = () => closeStyleMappings(true, 'test-');
  document.getElementById("cancel-test-style-mappings").onclick = () => closeStyleMappings(false, 'test-');
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
    let selection = context.document.getSelection().tables;
    context.load(selection);
    await context.sync();
    await axios.post("http://localhost:5000/retrieve", {se: selection.items[0].values})
    let lines = SELECTION;
    //try catch block for backend node call to prevent errors crashing the application
    try {
      // let call1 = await axios.post("http://localhost:5000/retrieve", { lines: lines })
    }
    catch (err) {
      console.log(err)
    }
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
  overpopulate the projects selector with duplicate sets.*/
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
  //formatting the URL as it should be to populate projects / validate user credentials
  let validatingURL = finalUrl || url + apiBase + `?username=${username}&api-key=${rssToken}`;
  try {
    //call the projects API to populate relevant projects
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
      //populate the projects dropdown with the response body.
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
  // Tests the pushRequirements Function
  let id = document.getElementById('project-select').value;
  for (let i = 0; i < requirements.length; i++) {
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

/*indents requirements to the appropriate level, relative to the last requirement in the project
before this add-on begins to add more. (No way to find out indent level of the last requirement
  in a project from the Spira API (i think))*/
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
*/
const pushTestCases = async () => {
  // CURRENTLY USED FOR TESTING
  let folderResponse = await pushTestCaseFolder("Test Folder", "First Functional Folder Test");
  let testCaseResponse = await pushTestCase("test case", folderResponse)
  try {
    let call1 = await axios.post("http://localhost:5000/retrieve", { Folder: folderResponse, TestCase: testCaseResponse })
  }
  catch (err) {
    console.log(err);
  }
}
/* 
  Creates a test case using the information given and sends it to the Spira instance. Returns the Id of the created test case
*/
const pushTestCase = async (testCaseName, testFolderId) => {
  try {
    let testCaseResponse = await axios.post(`${USER_OBJ.url}/services/v6_0/RestService.svc/projects/24/test-cases?username=${USER_OBJ.username}
      &api-key=${USER_OBJ.password}`, {
      Name: testCaseName,
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
  let folderCall = "err";
  try {
    folderCall = await axios.post(apiCall, {
      Name: folderName,
      Description: description
    })
    return folderCall.data.TestCaseFolderId;
  }
  catch (err) {
    console.log(err);
  }
}

const pushTestStep = async (testCaseId, testStep) => {
  /*pushTestCase should call this passing in the created testCaseId and iterate through passing
  in that test cases test steps.*/
  let projectId = document.getElementById('project-select').value;
  let apiCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + projectId +
    `/test-cases/${testCaseId}/test-steps?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
  try {
    //testStep = {Description: "", SampleData: "", ExpectedResult: ""}
    //we dont need the response from this - so no assigning to variable.
    await axios.post(apiCall, {
      Description: testStep.Description,
      SampleData: testStep.SampleData,
      ExpectedResult: testStep.ExpectedResult
    })
  }
  catch (err) {
    console.log(err)
  }
}
/******************** 
HTML DOM Manipulation
********************/

const populateProjects = (projects) => {
  let dropdown = document.getElementById('project-select')
  projects.forEach((project) => {
    /*creates an option for each project which displays the name
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

const openStyleMappings = async () => {
  //opens the requirements style mappings if requirements is the selected artifact type
  /*all id's and internal word settings are now set using a "pageTag". This allows code 
  to be re-used between testing and requirement style settings. The tags are req- for
  requirements and test- for test cases.*/
  let pageTag;
  document.getElementById("main-screen").classList.add("hidden")
  //checks the current selected artifact type then loads the appropriate menu
  if (document.getElementById("artifact-select").value == "requirements") {
    pageTag = "req-"
    document.getElementById("req-style-mappings").style.display = 'flex'
    //populates all 5 style mapping boxes
  }
  //opens the test cases style mappings if test mappings is the selected artifact type
  else {
    pageTag = "test-"
    document.getElementById("test-style-mappings").style.display = 'flex'
  }
  //retrieveStyles gets the document's settings for the style mappings. Also auto sets default values
  let settings = retrieveStyles(pageTag)
  //Goes line by line and retrieves any custom styles the user may have used.
  let customStyles = await scanForCustomStyles();
  //only the top 2 select objects should have all styles. bottom 3 are table based (at least for now).
  if (pageTag == "test-") {
    for (let i = 1; i <= 2; i++) {
      populateStyles(customStyles.concat(Object.keys(Word.Style)), pageTag + 'style-select' + i.toString());
    }
    //bottom 3 selectors will be related to tables
    for (let i = 3; i <= 5; i++) {
      let tableStyles = ["column1", "column2", "column3", "column4", "column5"]
      populateStyles(tableStyles, pageTag + 'style-select' + i.toString())
    }
  }
  else {
    for (let i = 1; i <= 5; i++) {
      populateStyles(customStyles.concat(Object.keys(Word.Style)), pageTag + 'style-select' + i.toString());
    }
  }
  //move selectors to the relevant option
  settings.forEach((setting, i) => {
    document.getElementById(pageTag + "style-select" + (i + 1).toString()).value = setting
  })
}

//closes the style mapping page taking in a boolean 'result'
//pageTag is req or test depending on which page is currently open

const closeStyleMappings = (result, pageTag) => {
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
  //returns user to main screen
  document.getElementById("main-screen").classList.remove("hidden")
  document.getElementById("req-style-mappings").style.display = 'none'
  document.getElementById("test-style-mappings").style.display = 'none'
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

const handleErrors = (error) => {

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
    switch (line.style.toLowerCase()) {
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

const parseTestCases = (lines) =>{
  let testCases = []
  //styles = ['style1', 'style2', ...]
  let styles = retrieveStyles("test-")
  let testStep = {Description: "", ExpectedResult: "", SampleData: ""}
  lines.forEach((line) =>{
    let testCase = {folderName: "", Name: "", testSteps: []}
    //this handles whether a line is a folder name or test case name
    switch (line.style.toLowerCase()){
      case styles[0]: 
        testCase.folderName = line.text
      case styles[1]:
        testCase.Name = line.text
      case styles[2]:

    }

  })
}

// Updates selection array and then loops through it and adds any
// user-created styles found to its array and returns it. WIP
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


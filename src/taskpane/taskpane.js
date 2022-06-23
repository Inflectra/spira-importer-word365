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
//ignore it saying defaults doesnt exist, it does and using default does not work.
//makes sure put requests uses the proper content-type header
axios.defaults.headers.put['Content-Type'] = "application/json"
axios.defaults.headers.put['accept'] = "application/json"

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
  document.getElementById("send-to-spira").style.display = 'none'
  document.getElementById("hierarchy-err").style.display = "none"
  document.getElementById('table-err').style.display = 'none'
}

const setEventListeners = () => {
  // document.getElementById('test').onclick = () => test();
  document.getElementById('btn-login').onclick = () => loginAttempt();
  document.getElementById('dev-mode').onclick = () => devmode();
  document.getElementById('send-to-spira').onclick = () => pushArtifacts();
  document.getElementById('log-out').onclick = () => logout();
  document.getElementById("select-requirements").onclick = () => openStyleMappings("req-");
  document.getElementById("select-test-cases").onclick = () => openStyleMappings("test-");
  document.getElementById("confirm-req-style-mappings").onclick = () => confirmStyleMappings('req-');
  document.getElementById("confirm-test-style-mappings").onclick = () => confirmStyleMappings('test-');

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
  await axios.post(RETRIEVE, { "uh": "oh" })
  let newReqs = await parseRequirements();
  await axios.post(RETRIEVE, { newReqs: newReqs })
  // return Word.run(async (context) => {
  //   let body = context.document.getSelection().split(['\r']);
  //   let styles = retrieveStyles('req-');
  //   context.load(body, ['text', 'style', 'styleBuiltIn'])
  //   await context.sync();
  //   //body = {items: [{text, style, styleBuiltIn}, ...]}
  //   //check for next line with a style within the styles
  //   let descStart, descEnd;
  //   let requirements = []
  //   let requirement = { Name: "", Description: "", RequirementTypeId: 1 }
  //   //need to use .entries to get index within a for .. of loop.
  //   for (let [i, item] of body.items.entries()) {
  //     if (styles.includes(item.styleBuiltIn) || styles.includes(item.style)) {
  //       //if a description has a starting range already, assign descEnd to the previous line
  //       if (descStart) {
  //         descEnd = body.items[i - 1]
  //         //creates an entre "description range" to be gotten as html.
  //         let descRange = descStart.expandToOrNullObject(descEnd)
  //         await context.sync();
  //         //descRange is null if the descEnd is not valid to extend the descStart
  //         if (!descRange) {
  //           descRange = descStart
  //         }
  //         //clears these fields so it is known the description has been added to its requirement
  //         descStart = undefined
  //         descEnd = undefined
  //         let descHtml = descRange.getHtml();
  //         await context.sync();
  //         //m_value is the actual string of the html
  //         requirement.Description = descHtml.m_value
  //         try {
  //           requirements.push(requirement)
  //           requirement = { Name: item.text, Description: "", RequirementTypeId: 1 }
  //         }
  //         catch (err) {
  //           await axios.post(RETRIEVE, { err: err })
  //         }
  //       }
  //       /*if there isnt a description section between the last name and this name, push to spira
  //       and start a new requirement*/
  //       else if (requirement.Name){
  //         requirements.push(requirement)
  //         requirement = { Name: item.text, Description: "", RequirementTypeId: 1 }
  //       }
  //       else {
  //         requirement.Name = item.text
  //       }
  //       continue
  //     }
  //     else {
  //       if (!descStart) {
  //         descStart = item
  //       }
  //     }
  //     /*if a line is the last line in the selection/body
  //      & the current requirement object has least a name, push the last requirement*/
  //     if (i == body.items.length -1 && requirement.Name){
  //       if(item.text && descStart){
  //         descEnd = item
  //         let descRange = descStart.expandToOrNullObject(descEnd)
  //         await context.sync();
  //         //descRange is null if the descEnd is not valid to extend the descStart
  //         if (!descRange) {
  //           descRange = descStart
  //         }
  //         let descHtml = descRange.getHtml();
  //         await context.sync();
  //         requirement.Description = descHtml.m_value
  //       }
  //       requirements.push(requirement)
  //     }
  //   }
  //   await axios.post(RETRIEVE, [...requirements, descStart])
  //   return requirements
  // }
  // )
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
  let requirements = await parseRequirements();
  //requirements = [{Name, Description, RequirementTypeId, IndentLevel}]
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
  let pid = document.getElementById('project-select').value;
  let firstReq = requirements[0];
  //this call is for the purpose of resetting the indent level each time a set of reqs are sent
  const outdentCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + pid +
    `/requirements/indent/-20?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
  try {
    //image parsing needs to be re-integrated here and in the try/catch block in the for loop.
    let firstCall = await axios.post(outdentCall, firstReq);
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
  //this handles the rest of the requirements calls which are indented relative to the first.
  for (let i = 1; i < requirements.length; i++) {
    let item = requirements[i];
    const apiCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + pid +
      `/requirements?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
    // try catch block to stop application crashing and show error message if call fails
    try {
      //indent requirement working perfectly
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
*/
const pushTestCases = async () => {
  await updateSelectionArray();
  // testCases = {folderName: "", Name: "", testSteps: [{Description, expected result, sample data}, ...]}
  let testCases = await parseTestCases(SELECTION);
  //if parseTestCases fails due to bad table styles, stops execution of the function
  if (!testCases) {
    return
  }
  // testCaseFolders = [{Name: "", TestCaseFolderId: int}, ...]
  let testCaseFolders = await retrieveTestCaseFolders();
  for (let i = 0; i < testCases.length; i++) {
    let testCase = testCases[i];
    // First check if it's in an existing folder
    let folder = testCaseFolders.find(folder => folder.Name == testCase.folderName)
    if (!folder) { // If the folder doesn't exist yet, make it and then make the 
      let newFolder = {}
      newFolder.TestCaseFolderId = await pushTestCaseFolder(testCase.folderName, testCase.folderDescription);
      newFolder.folderName = testCase.folderName;
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
/* 
  Creates a test case using the information given and sends it to the Spira instance. Returns the Id of the created test case
*/
const pushTestCase = async (testCaseName, testCaseDescription, testFolderId) => {
  let projectId = document.getElementById("project-select").value
  try {
    var testCaseResponse = await axios.post(`${USER_OBJ.url}/services/v6_0/RestService.svc/projects/${projectId}/test-cases?username=${USER_OBJ.username}
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
/*image should be the image object with base64 string and relevant metadata, Artifact is the 
post response from creation, you need to GET the artifact from the spira API each time you 
PUT due to the Concurrency date being checked by API.*/
//this function can be optimized to make 1 put request per artifact, rather than 1 per image
const pushImage = async (Artifact, image) => {
  let pid = Artifact.ProjectId
  //image = {base64: "", name: "", lineNum: int}
  /*upload images and build link of image location in spira 
  ({USER_OBJ.url}/{projectID}/Attachment/{AttachmentID}.aspx)*/
  //Add AttachmentURL to each imageObject after they are uploaded
  let imgLink;
  let imageApiCall = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/"
    + pid + `/documents/file?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`
  try {
    if (Artifact.RequirementId) {
      let imageCall = await axios.post(imageApiCall, {
        FilenameOrUrl: image.name, BinaryData: image.base64,
        AttachedArtifacts: [{ ArtifactId: Artifact.RequirementId, ArtifactTypeId: 1 }]
      })
      imgLink = USER_OBJ.url + `/${pid}/Attachment/${imageCall.data.AttachmentId}.aspx`
    }
    else if (Artifact.TestCaseId) {
      let imageCall = await axios.post(imageApiCall, {
        FilenameOrUrl: image.name, BinaryData: image.base64,
        AttachedArtifacts: [{ ArtifactId: Artifact.TestCaseId, ArtifactTypeId: 2 }]
      })
    }
  }
  catch (err) {
    console.log(err)
  }
  let fullArtifactObj;
  //checks if the artifact is a requirement (if not it is a test case)
  if (Artifact.RequirementId) {
    try {
      //makes a get request for the target artifact which will be updated to contain an image
      let getArtifact = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + pid +
        `/requirements/${Artifact.RequirementId}?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
      let getArtifactCall = await superagent.get(getArtifact).set('accept', 'application/json').set('Content-Type', 'application/json');
      //This is the body of the get response in its entirety.
      fullArtifactObj = getArtifactCall.body
    }
    catch (err) {
      //do nothing
      console.log(err)
    }
    // now replace the placeholder in the description with img tags
    let placeholderRegex = /\[inline[\d]*\.jpg\]/g
    //gets an array of all the placeholders for images. 
    let placeholders = [...fullArtifactObj.Description.matchAll(placeholderRegex)]
    /*placeholders[0][0] is the first matched instance - because you need to GET for each change
    this should work each time - each placeholder should have 1 equivalent image in the same
    order they appear throughout the document.*/
    fullArtifactObj.Description = fullArtifactObj.Description.replace(placeholders[0][0], `<img alt=${image.name} src=${imgLink}><br />`)
    //PUT artifact with new description (including img tags now)
    let putArtifact = USER_OBJ.url + "/services/v6_0/RestService.svc/projects/" + pid +
      `/requirements?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
    try {
      await axios.put(putArtifact, fullArtifactObj)
    }
    catch (err) {
      //do nothing
      console.log(err)
    }
  }
  else if (Artifact.TestCaseId) {
    //handle test cases
  }
  else {
    //handle error (should never reach here, but if it does it should be handled)
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
  /*hides the send button when they select an artifact type, so if they had completed
  style mappings for one or the other then change the artifact type they are required to enter
  their style mappings again.*/
  document.getElementById("send-to-spira").style.display = "none"
  //checks the current selected artifact type then loads the appropriate menu
  if (pageTag == "req-") {
    document.getElementById("select-requirements").style['background-color'] = "#022360"
    document.getElementById("select-test-cases").style['background-color'] = "#0078d7"
    document.getElementById("req-style-mappings").classList.remove("hidden")
    document.getElementById("test-style-mappings").style.display = 'none'
    document.getElementById("req-style-mappings").style.display = 'flex'
    //populates all 5 style mapping boxes
  }
  //opens the test cases style mappings if test mappings is the selected artifact type
  else {
    document.getElementById("select-test-cases").style['background-color'] = "#022360"
    document.getElementById("select-requirements").style['background-color'] = "#0078d7"
    document.getElementById("req-style-mappings").style.display = 'none'
    document.getElementById("test-style-mappings").style.display = 'flex'
  }
  //wont populate styles for requirements if it is already populated
  if (document.getElementById("req-style-select1").childElementCount && pageTag == "req-") {
    return
  }
  //wont populate styles for test-cases if it is already populated
  else if (document.getElementById("test-style-select1").childElementCount && pageTag == "test-") {
    return
  }
  //doesnt populate styles if both test & req style selectors are populated
  else if (document.getElementById("test-style-select1").childElementCount && document.getElementById("req-style-select1").childElementCount) {
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

const confirmStyleMappings = async (pageTag) => {
  //saves the users style preferences. this is document bound
  let styles = []
  for (let i = 1; i <= 5; i++) {
    if (document.getElementById(pageTag + "style-select" + i.toString()).value) {
      let setting = document.getElementById(pageTag + "style-select" + i.toString()).value
      //checks if a setting is used multiple times
      if (!styles.includes(setting)) {
        styles.push(setting)
      }
      //gives an error explaining they have empty style mappings and cannot proceed.
      else {
        document.getElementById("duplicate-styles-err").classList.remove("hidden")
        setTimeout(() => {
          document.getElementById("duplicate-styles-err").classList.add("hidden")
        }, 8000)
        //hides the final button if it is already displayed when a user inputs invalid styles.
        document.getElementById("send-to-spira").style.display = "none"
        return
      }
      Office.context.document.settings.set(pageTag + 'style' + i.toString(), setting);
    }
    //gives an error explaining they have duplicate style mappings which conflict.
    else {
      document.getElementById("empty-styles-err").classList.remove("hidden")
      setTimeout(() => {
        document.getElementById("empty-styles-err").classList.add("hidden")
      }, 8000)
      document.getElementById("send-to-spira").style.display = "none"
      return
    }
  }
  //hides error on successful confirm
  document.getElementById("duplicate-styles-err").classList.add("hidden")
  document.getElementById("empty-styles-err").classList.add("hidden")
  //this saves the settings
  Office.context.document.settings.saveAsync()
  //show the send to spira button after this is clicked and all style selectors are populated.
  document.getElementById("send-to-spira").style.display = "inline-block"
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
      //loads the text, style elements, and any images from a given line
      context.load(selection, ['text', 'styleBuiltIn', 'style', 'inlinePictures'])
      await context.sync();
    }

    // if nothing is selected, select the entire body of the document
    else {
      selection = context.document.body.getRange().split(['/r']);
      context.load(selection, ['text', 'styleBuiltIn', 'style', 'inlinePictures'])
      await context.sync();
    }
    // Testing parsing lines of text from the selection array and logging it
    let lines = []
    selection.items.forEach((item) => {
      lines.push({
        text: item.text, style: (item.styleBuiltIn == "Other" ? item.style : item.styleBuiltIn),
        custom: (item.styleBuiltIn == "Other"), images: item.inlinePictures
      })
    })
    SELECTION = lines;
    return
  })
}

const parseRequirements = async () => {
  return Word.run(async (context) => {
    let body = context.document.getSelection();
    let styles = retrieveStyles('req-');
    context.load(body, ['text', 'style', 'styleBuiltIn'])
    await context.sync();
    //if there is no selection, select the whole body
    if (!body.text) {
      body = context.document.body.getRange();
      context.load(body, ['text', 'style', 'styleBuiltIn'])
      await context.sync();
      await axios.post(RETRIEVE, {in: 'here'})
    }
    body = body.split(['\r'])
    context.load(body)
    await context.sync();
    //body = {items: [{text, style, styleBuiltIn}, ...]}
    //check for next line with a style within the styles
    await axios.post(RETRIEVE, {items: body.items})
    let descStart, descEnd;
    let requirements = []
    let requirement = { Name: "", Description: "", RequirementTypeId: 2, IndentLevel: 0 }
    //need to use .entries to get index within a for .. of loop.
    for (let [i, item] of body.items.entries()) {
      if (styles.includes(item.styleBuiltIn) || styles.includes(item.style)) {
        //if a description has a starting range already, assign descEnd to the previous line
        if (descStart) {
          descEnd = body.items[i - 1]
          //creates an entre "description range" to be gotten as html.
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
          requirement.Description = descHtml.m_value.replaceAll("\r", "")
          try {
            requirements.push(requirement)
            let indent = styles.indexOf(item.styleBuiltIn)
            if (!indent) {
              indent = styles.indexOf(item.style)
            }
            requirement = {
              Name: item.text.replaceAll("\r", ""), Description: "",
              RequirementTypeId: 2, IndentLevel: indent
            }
          }
          catch (err) {
            await axios.post(RETRIEVE, { err: err })
          }
        }
        /*if there isnt a description section between the last name and this name, push to spira
        and start a new requirement*/
        else if (requirement.Name) {
          requirements.push(requirement)
          let indent = styles.indexOf(item.styleBuiltIn)
          if (!indent) {
            indent = styles.indexOf(item.style)
          }
          requirement = {
            Name: item.text.replaceAll("\r", ""), Description: "",
            RequirementTypeId: 2, IndentLevel: indent
          }
        }
        //this should only happen on the first (valid) line parsed
        else {
          requirement.Name = item.text.replaceAll("\r", "")
        }
        continue
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
        //
        if (item.text) {
          let descRange;
          if (descStart) {
            descEnd = item

            descRange = descStart.expandToOrNullObject(descEnd)
          }
          else {
            descRange = descStart
          }
          await context.sync();
          //descRange is null if the descEnd is not valid to extend the descStart
          if (!descRange) {
            descRange = descStart
          }
          let descHtml = descRange.getHtml();
          await context.sync();
          requirement.Description = descHtml.m_value.replaceAll("\r", "")
        }
        requirements.push(requirement)
      }
    }
    if (validateHierarchy(requirements)) {
      await axios.post(RETRIEVE, [...requirements, descStart])
      return requirements
    }
    return []
  }
  )
}

/* Gets an array of all the tables from the Word document and returns it. */
const retrieveTables = async () => {
  return Word.run(async (context) => {
    let check = context.document.getSelection();
    context.load(check, 'text');
    await context.sync();
    //checks if there is a selection
    if (check.text) {
      let selection = context.document.getSelection().tables;
      context.load(selection);
      await context.sync();
      let tables = [];
      for (let i = 0; i < selection.items.length; i++) {
        let table = selection.items[i].values;
        tables.push(table);
      }
      return tables;
    }
    //If there is not a selection, retrieve tables from the entire body.
    else {
      let selection = context.document.body.tables;
      context.load(selection);
      await context.sync();
      let tables = [];
      for (let i = 0; i < selection.items.length; i++) {
        let table = selection.items[i].values;
        tables.push(table);
      }
      return tables;
    }
  })
}

const retrieveImages = async () => {
  //spira API call requires FilenameOrUrl and binary base 64 encoded data
  //word api has context.document.body.inlinePictures.items[i].getBase64ImageSrc()
  return Word.run(async (context) => {
    await updateSelectionArray();
    let imageLines = []
    for (let i = 0; i < SELECTION.length; i++) {
      //go through each line, if it has images extract them with the line number
      if (SELECTION[i].images.items[0]) {
        for (let j = 0; j < SELECTION[i].images.items.length; j++) {
          //pushes a line number for each image so its known where they get placed
          imageLines.push(i)
        }
      }
    }
    let imageObjects = []
    let imageObj;
    let images = context.document.getSelection();
    context.load(images, 'text');
    await context.sync();
    //if there is a selection, gets all images in that selection
    if (images.text) {
      images = context.document.getSelection().inlinePictures;
      images.load();
      await context.sync();
    }
    // if nothing is selected, select the entire body of the document
    else {
      images = context.document.body.inlinePictures;
      images.load()
      await context.sync();
    }
    //creates image objects with base64 encoded string, name, and the line the image is on. 
    for (let i = 0; i < images.items.length; i++) {
      let base64 = images.items[i].getBase64ImageSrc();
      await context.sync();
      //lineNum starts counting at 0, so first line is line 0. 
      imageObj = { base64: base64.m_value, name: `inline${i}.jpg`, lineNum: imageLines[0] }
      imageObjects.push(imageObj)
      imageObj = {}
      imageLines.shift();
    }
    return imageObjects
  })
}
/*********************
Pure data manipulation
**********************/

const pushArtifacts = async () => {
  let active = document.getElementById("req-style-mappings").style.display
  //if the requirements style mappings are visible that is the selected artifact.
  if (active == "flex") {
    await pushRequirements();
  }
  else {
    await pushTestCases();
  }
}

const parseTestCases = async (lines) => {
  let testCases = []
  //styles = ['style1', 'style2', columnStyle, columnStyle, columnStyle]
  let styles = retrieveStyles("test-")
  let testCase = { folderName: "", folderDescription: "", Name: "", testCaseDescription: "", testSteps: [] }
  //tables = [[test case 1 steps], [test case 2 steps], ...]
  //word API functions cannot return normally.
  let tables = await retrieveTables()
  //makes sure each table has at least 1 description based on chosen table mappings
  if (!validateTestSteps(tables, styles[2])) {
    return false
  }
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

    /*this checks if a line is the first description in the next table (a table that has not been parsed
       still exists, the line ends in \t character, and the text matches the tables first description)*/
    if (tables[0] && lines[i].text == tables[0][0][parseInt(styles[2].slice(-1)) - 1]?.concat("\t") && lines[i].text.slice(-1) == "\t") {
      let testSteps = parseTestSteps(tables[0])
      //allows multiple tables to populate the same test case in the case it is multiple tables
      testCase = { ...testCase, testSteps: [...testSteps] }
      /*removes the table which has just been parsed so we dont need to iterate through tables
      in the conditional. */
      tables.shift();
    }
    /*if the relevant fields are populated and the current line would replace the 
    folderName/Name or is the last line, push the test case and reset the testCase variable */
    if (testCase.folderName && testCase.Name && (lines[i].style == styles[0] || lines[i].style == styles[1] || i == lines.length - 1)) {
      testCases.push(testCase)
      /*clears testCase besdies folder name. If there is a new folder name it gets replaced
      in the following switch statement.*/
      testCase = { Name: "", testCaseDescription: "", folderDescription: "", folderName: testCase.folderName, testSteps: [] }
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
        //this only works for 1 line / 1 paragraph descriptions. 
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
  }
  return testCases
}

const parseTestSteps = (table) => {
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
      document.getElementById('table-err').style.display = "flex"
      return false
    }
  }
  document.getElementById('table-err').style.display = "none"
  return true
}
/*first row is the first row of the table being checked - line row is the accompanying
x number of 'lines' from the global SELECTION object.*/

const rowCheck = (firstRow, lineRow) => {
  //for loop, compare inner items at i
}

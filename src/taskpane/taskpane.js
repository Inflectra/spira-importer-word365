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
import * as components from './model'

// Global selection array, used throughout
/*This is a global variable because the word API call functions are unable to return
values from within due to the required syntax of returning a Word.run((callback) =>{}) 
function. */
var model = new components.Data();
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
  document.getElementById('test').onclick = () => test();
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
  let bruh = await newParseTestCases();
}
/**************
Spira API calls
**************/

// REFACTOR IN PROGRESS 
// Replaced USER_OBJ with the new model.user.(field) format
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
      model.user.url = finalUrl || url;
      model.user.username = username;
      model.user.api_key = rssToken;
      model.user.userCredentials = `?username=${username}&api-key=${rssToken}`;
      //populate the products dropdown & model object with the response body.
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
  // Disable the Send to Spira button until the requirements have sent
  document.getElementById("send-to-spira-button").disabled = true;
  await updateSelectionArray();
  let images = await retrieveImages();
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
  const outdentCall = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
    `/requirements/indent/-20?username=${model.user.username}&api-key=${model.user.api_key}`;
  // Make progress bar appear
  document.getElementById("progress-bar-progress").style.width = "0%";
  document.getElementById("progress-bar").classList.remove("hidden");
  try {
    //image parsing needs to be re-integrated here and in the try/catch block in the for loop.
    let firstCall = await axios.post(outdentCall, firstReq);
    //picks out <img> tags from html string (requirement description)
    let imgRegex = /<img(.|\n)*("|\s)>/g
    let placeholders = [...firstReq.Description.matchAll(imgRegex)]
    for (let i = 0; i < placeholders.length; i++) {
      pushImage(firstCall.data, images[0])
      images.shift();
    }
    updateProgressBar(1, requirements.length);
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
    const apiCall = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
      `/requirements?username=${model.user.username}&api-key=${model.user.api_key}`;
    // try catch block to stop application crashing and show error message if call fails
    try {
      //indent requirement working perfectly
      let call = await axios.post(apiCall, { Name: item.Name, Description: item.Description, RequirementTypeId: 2 });
      await indentRequirement(apiCall, call.data.RequirementId, item.IndentLevel - lastIndent)
      lastIndent = item.IndentLevel;
      let imgRegex = /<img(.|\n)*("|\s)>/g
      let placeholders = [...item.Description.matchAll(imgRegex)]
      for (let j = 0; j < placeholders.length; j++) {
        pushImage(call.data, images[0])
        images.shift();
      }
      updateProgressBar(i + 1, requirements.length);
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
  // Hide progress bar again now that it's finished, and enable Send to Spira button
  document.getElementById("progress-bar").classList.add("hidden");
  document.getElementById("send-to-spira-button").disabled = false;
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
  let images = await retrieveImages();
  document.getElementById("send-to-spira-button").disabled = true;
  await updateSelectionArray();
  // testCases = [{folderName: "", Name: "", testSteps: [{Description, expected result, sample data}, ...]}]
  let testCases = await newParseTestCases();
  document.getElementById("progress-bar-progress").style.width = "0%";
  document.getElementById("progress-bar").classList.remove("hidden");
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
      newFolder.Name = testCase.folderName;
      folder = newFolder
      testCaseFolders.push(newFolder);
    }
    // make the testCase and keep the Id for later
    let testCaseArtifact = await pushTestCase(testCase.Name, testCase.testCaseDescription, folder.TestCaseFolderId);
    //Uses this to determine how many images need to be placed within the document.
    let placeholderRegex = /<img(.|\n|\r)*("|\s)\>/g
    //gets an array of all the placeholders for images. 
    let placeholders = [...testCase.testCaseDescription.matchAll(placeholderRegex)]
    for (let j = 0; j < placeholders.length; j++) {
      await pushImage(testCaseArtifact, images[0])
      images.shift();
    }
    // now make the testSteps
    for (let j = 0; j < testCase.testSteps.length; j++) {
      let step = await pushTestStep(testCaseArtifact.TestCaseId, testCase.testSteps[j]);
      if (images[0]) {
        await pushImage(step, images[0], testCaseArtifact.TestCaseId);
        images.shift();
      }
    }
    updateProgressBar(i + 1, testCases.length);
  }
  document.getElementById("progress-bar").classList.add("hidden");
  document.getElementById("send-to-spira-button").disabled = false;
}

const retrieveTestCaseFolders = async () => {
  let projectId = document.getElementById('project-select').value;
  let apiCall = model.user.url + "/services/v6_0/RestService.svc/projects/" + projectId +
    `/test-folders?username=${model.user.username}&api-key=${model.user.api_key}`;
  let callResponse = await superagent.get(apiCall).set('accept', "application/json").set('Content-Type', "application/json")
  return callResponse.body
}

const pushTestStep = async (testCaseId, testStep) => {
  /*pushTestCase should call this passing in the created testCaseId and iterate through passing
  in that test cases test steps.*/
  let projectId = document.getElementById('project-select').value;
  let apiCall = model.user.url + "/services/v6_0/RestService.svc/projects/" + projectId +
    `/test-cases/${testCaseId}/test-steps?username=${model.user.username}&api-key=${model.user.api_key}`;
  try {
    //testStep = {Description: "", SampleData: "", ExpectedResult: ""}
    //we dont need the response from this - so no assigning to variable.
    let stepCall = await axios.post(apiCall, {
      Description: testStep.Description,
      SampleData: testStep.SampleData,
      ExpectedResult: testStep.ExpectedResult
    })
    return stepCall.data
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
    var testCaseResponse = await axios.post(`${model.user.url}/services/v6_0/RestService.svc/projects/${projectId}/test-cases?username=${model.user.username}
      &api-key=${model.user.api_key}`, {
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
/*  
  Creates a test folder and returns the Test Folder Id
*/
const pushTestCaseFolder = async (folderName, description) => {
  let projectId = document.getElementById('project-select').value;
  let apiCall = model.user.url + "/services/v6_0/RestService.svc/projects/" + projectId +
    `/test-folders?username=${model.user.username}&api-key=${model.user.api_key}`;
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
const pushImage = async (Artifact, image, testCaseId) => {
  let pid
  if (Artifact.ProjectId != 0) {
    pid = Artifact.ProjectId
  }
  //Test steps do not populate ProjectId on POST while other artifact types do. 
  else {
    pid = document.getElementById("project-select").value
  }
  //image = {base64: "", name: "", lineNum: int}
  /*upload images and build link of image location in spira 
  ({model.user.url}/{projectID}/Attachment/{AttachmentID}.aspx)*/
  //Add AttachmentURL to each imageObject after they are uploaded
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
      //do nothing
      console.log(err)
    }
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
      `/requirements?username=${model.user.username}&api-key=${model.user.api_key}`;
    try {
      await axios.put(putArtifact, fullArtifactObj)
    }
    catch (err) {
      //do nothing
      console.log(err)
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
    }
    // now replace the placeholder in the description with img tags
    let placeholderRegex = /<img(.|\n|\r)*?("|\s)>/g
    //gets an array of all the placeholders for images. 
    let placeholders = [...fullArtifactObj.Description.matchAll(placeholderRegex)]
    /*placeholders[0][0] is the first matched instance - because you need to GET for each change
    this should work each time - each placeholder should have 1 equivalent image in the same
    order they appear throughout the document.*/
    fullArtifactObj.Description = fullArtifactObj.Description.replace(placeholders[0][0], `<img alt=${image?.name} src=${imgLink}><br />`)
    //PUT artifact with new description (including img tags now)
    let putArtifact = model.user.url + "/services/v6_0/RestService.svc/projects/" + pid +
      `/test-cases/${fullArtifactObj.TestCaseId}/test-steps?username=${model.user.username}&api-key=${model.user.api_key}`;
    try {
      await axios.put(putArtifact, fullArtifactObj)
    }
    catch (err) {
      //do nothing
      console.log(err)
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
    }
  }
  else {
    //handle error (should never reach here, but if it does it should be handled)
  }
  return
}

/********************
HTML DOM Manipulation
********************/

const populateProjects = (projects) => {
  model.projects = projects
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
  model.clearUser();
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

const updateProgressBar = (current, total) => {
  const MAX_WIDTH = 80; //CURRENTLY HARDCODED TO BE the same as the width of the progress-bar CSS element
  let width = current / total * MAX_WIDTH;
  let bar = document.getElementById("progress-bar-progress");
  bar.style.width = width + "%";
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
    }
    body = body.split(['\r'])
    context.load(body)
    await context.sync();
    //body = {items: [{text, style, styleBuiltIn}, ...]}
    //check for next line with a style within the styles
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
          requirement.Description = await filterForLists(descHtml.m_value.replaceAll("\r", ""));
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
          requirement.Description = await filterForLists(descHtml.m_value.replaceAll("\r", ""));
        }
        requirements.push(requirement)
      }
    }
    if (validateHierarchy(requirements)) {
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

const newParseTestCases = async () => {
  /*Currently I am returning empty test cases with no names / testCaseDescriptions. Test steps partly work - 
  there is an issue if the table cell has multiple lines. Its iterating through to the end at least so 
  that is progress. */
  return Word.run(async (context) => {
    let tableCounter = 0
    let testCases = []
    let styles = retrieveStyles('test-')
    let testCase = { folderName: "", folderDescription: "", Name: "", testCaseDescription: "", testSteps: [] }
    //tables = 3d array [table][test step][column]
    let tables = await retrieveTables();
    //this checks to make sure tables selected have at least 1 description within them. 
    if (!validateTestSteps(tables, styles[2])) {
      return false
    }
    let body = context.document.getSelection();
    context.load(body, ['text', 'style', 'styleBuiltIn'])
    await context.sync();
    //if there is no selection, select the whole body
    if (!body.text) {
      body = context.document.body.getRange();
      context.load(body, ['text', 'style', 'styleBuiltIn'])
      await context.sync();
    }
    body = body.split(['\r'])
    context.load(body, ['text', 'style', 'styleBuiltIn'])
    await context.sync();
    let descStart, descEnd;
    for (let [i, item] of body.items.entries()) {
      let itemtext = item.text.replaceAll("\r", "")
      //checks if the line is a style which is mapped to the style mappings
      if (styles.includes(item.style) || styles.includes(item.styleBuiltIn)) {
        if (descStart) {
          descEnd = body.items[i - 1]
          let descRange = descStart.expandToOrNullObject(descEnd);
          context.load(descRange)
          await context.sync();
          /*if the descRange returns null (doesnt populate a range), assume the range
          is only the starting line.*/
          let descHtml = descRange.getHtml();
          await context.sync()
          descStart = undefined; descEnd = undefined;
          //if the current item is a folder name, update folder name and initialize new testCase
          if (!testCase.Name) {
            testCase.folderDescription = descHtml.m_value.replaceAll("\r", "")
          }
          else {
            //removes tables picked up in the description and adds proper HTML lists
            let filteredDescription = await filterForLists(descHtml.m_value.replaceAll("\r", "")); // filter for LISTS!!!
            let tableRegex = /<table(.|\n|\r)*?\/table>/g
            let descriptionTables = [...filteredDescription.matchAll(tableRegex)]
            for (let j = 0; j < descriptionTables.length; j++) {
              filteredDescription = filteredDescription.replace(descriptionTables[j][0], "")
            }
            testCase.testCaseDescription = filteredDescription
          }
          if (item.style == styles[0] || item.styleBuiltIn == styles[0]) {
            if (testCase.Name) {
              testCases.push(testCase)
            }
            testCase = { folderName: itemtext, folderDescription: "", Name: "", testCaseDescription: "", testSteps: [] }
          }
          else {
            if (testCase.Name) {
              testCases.push(testCase)
              //makes a new testCase object with the old folderName in case there is not a new one.
              testCase = { folderName: testCases[testCases.length - 1].folderName, folderDescription: "", Name: "", testCaseDescription: "", testSteps: [] }
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
        }
      }
      //this conditional verifies that the text that was found is the next table.
      else if (tables[0] && item.text == tables[0][0][parseInt(styles[2].slice(-1)) - 1]?.concat("\t") && item.text.slice(-1) == "\t") {
        //This procs when there is a table and the first description equals item.text
        //testStepTable = 2d array [row][column]
        let testStepTable = await newParseTestSteps(tableCounter);
        //table counter lets parseTestSteps know which table is currently being parsed
        tableCounter++
        let testStep = { Description: "", ExpectedResult: "", SampleData: "" }
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
          let emptyStepRegex = /<p(.)*?>\&nbsp\;<\/p>/g

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
      //second part of this conditional gates tables from becoming the description start
      else if (!descStart && item.text.slice(-1) != "\t") {
        descStart = item
      }
      //if it is the last line, add description if relevant, then push to testCases
      if (i == (body.items.length - 1)) {
        if (descStart) {
          descEnd = body.items[i]
          let descRange = descStart.expandToOrNullObject(descEnd);
          await context.sync();
          /*if the descRange returns null (doesnt populate a range), assume the range
          is only the starting line.*/
          let descHtml = descRange.getHtml();
          await context.sync();
          let filteredDescription = await filterForLists(descHtml.m_value.replaceAll("\r", ""))
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
    return testCases
  })
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

const newParseTestSteps = async (tableIndex) => {
  return Word.run(async (context) => {
    //validates a selection or full document
    let selection = context.document.getSelection();
    context.load(selection, ['text'])
    await context.sync();
    let tables, body;
    if (!selection.text) {
      //if there isnt a selection, insteasd parses the entire body
      tables = context.document.body.tables
      body = context.document.body.getRange().split(["\r"]);
      context.load(tables);
      context.load(body, ['text', 'inlinePictures']);
      await context.sync();
    }
    else {
      tables = context.document.getSelection().tables;
      body = context.document.getSelection().split(['\r']);
      context.load(tables);
      context.load(body, ['text', 'inlinePictures']);
      await context.sync();
    }

    //tables.items[0].values[0].length is the length of each row.
    let length = tables.items[tableIndex].values[0].length;
    for (let item of body.items) {
      if (item.text.slice(-1) == "\t") {
        let html = tables.items[tableIndex].getRange().getHtml();
        await context.sync();
        //filter out all <p> -> </p> tag instances 
        //ungreedy *? tag for matching only to the first, global ungreedy single-line tags at end of regex
        let tableRegex = /(<p )(.|\n|\s|\r)*?(<\/p>)/gus
        //this will give me all the text within a table separated by regex - then i can use i % length and i/length to know my place in the "table"
        let elements = [...html.m_value.matchAll(tableRegex)]
        let formattedStrings = []
        for (let [i, element] of elements.entries()) {
          //makes sure the first index exists, if not creates it as an empty array, then pushes as 2d array
          if (formattedStrings[Math.floor(i / length)]) {
            formattedStrings[Math.floor(i / length)][i % length] = (element[0])
          }
          else {
            formattedStrings[Math.floor(i / length)] = []
            formattedStrings[Math.floor(i / length)][i % length] = (element[0])
          }
        }
        //we should get the row length of the table
        //formattedStrings mimics the table structure as a 2d array string
        return formattedStrings
      }
    }
  })
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

/* Takes a single <p> to </p> element and turns it into a list element if it has the necessary class*/
const convertToListElem = (pElem) => {
  let orderedRegEx = />.{1,2}<span/g;
  if (pElem.includes("class=MsoListParagraphCxSpFirst")) { //Case for if the element is the first element in a list
    //Must add extra html element codes at the beginning and end of the list to wrap the list elements together.
    pElem = listDelimiter(pElem, true); // starts a list
    pElem = pElem.replace("<p ", "<li ").replace("</p>", "</li>").replaceAll(orderedRegEx, "><span");
    pElem = pElem.replaceAll("&nbsp;", "");
  }
  else if (pElem.includes("class=MsoListParagraphCxSpMiddle")) { //Case for if the element is within the same list.
    pElem = pElem.replace("<p ", "<li ").replace("</p>", "</li>").replaceAll(orderedRegEx, "><span");
    pElem = pElem.replaceAll("&nbsp;", "");
  }
  else if (pElem.includes("class=MsoListParagraphCxSpLast")) { //Case for if the element is the last element in a list.
    pElem = listDelimiter(pElem, false); // ends a list
    pElem = pElem.replace("<p ", "<li ").replace("</p>", "</li>").replaceAll(orderedRegEx, "><span");
    pElem = pElem.replaceAll("&nbsp;", "");
  }
  else if (pElem.includes("class=MsoListParagraph")) { //Case for if the element is the only element in the list
    pElem = listDelimiter(pElem, true); // starts a list
    pElem = listDelimiter(pElem, false); // ends a list
    pElem = pElem.replace("<p ", "<li ").replace("</p>", "</li>").replaceAll(orderedRegEx, "><span");
    pElem = pElem.replaceAll("&nbsp;", "");
  }
  //Case for if the element is not part of a list is handled by just returning it back.
  return pElem;
}

/* Filters a string and changes any word-outputted lists to properly formatted html lists. INDENTING IS NOT YET IMPLEMENTED*/
const filterForLists = async (description) => {
  let startRegEx = /(<p )(.|\n|\s|\r)*?(<\/p>)/gu;
  let elemList = [...description.matchAll(startRegEx)];
  description = await convertToIndentedList(description, elemList);
  return description
}

/* Scans each element in an array of 'strings' for "style='margin-left:#.0in" where the # is the indent level 
   Then it keeps track of the current indent level as it loops through the array, processing the elements
   through convertToListElem and adding an extra <ul> or <ol> as necessary to properly turn them into html 
   lists. */
const convertToIndentedList = async (description, elemList) => {
  let indentLevel = 0;
  for (let i = 0; i < elemList.length; i++) {
    /* Use elemList[i][0] in order to reach the matched strings. */
    let elem = elemList[i][0];
    let alteredElem = "" + elem;
    let result = alteredElem.match(/style='margin-left:(\d)\.(\d)in/);
    if (result) {
      let curIndentLevel = (parseInt(result[1]) * 2 + parseInt(result[2]) * 0.2) - 1
      while (curIndentLevel > indentLevel) {
        alteredElem = listDelimiter(alteredElem, true);
        indentLevel++;
      }
      while (curIndentLevel < indentLevel) {
        alteredElem = listDelimiter(alteredElem, false, true);
        indentLevel--;
      }
    }
    else {
      let curIndentLevel = 0;
      while (curIndentLevel < indentLevel) {
        alteredElem = listDelimiter(alteredElem, false, true);
        indentLevel--;
      }
    }
    description = description.replace(elem, convertToListElem(alteredElem));
  }
  return description;
}

/* Adds a <ul> or <ol> element based on the parameters and if the element is an unordered or ordered list. */
const listDelimiter = (elem, start, endPrefix) => {
  if (start) {
    if (elem.includes(">·<span") || elem.includes(">o<span") || elem.includes(">§<span")) { // Checks for if it should start an unordered or ordered list
      elem = "<ul>" + elem;
    }
    else {
      elem = "<ol>" + elem;
    }
  }
  else {
    if (endPrefix) { //case for when it needs to end the previous list element instead of the current one
      if (elem.includes(">·<span") || elem.includes(">o<span") || elem.includes(">§<span")) { // Checks for if it should end an unordered or ordered list
        elem = "</ul>" + elem;
      }
      else {
        elem = "</ol>" + elem;
      }
    }
    else {
      if (elem.includes(">·<span") || elem.includes(">o<span") || elem.includes(">§<span")) { // Checks for if it should end an unordered or ordered list
        elem = elem + "</ul>";
      }
      else {
        elem = elem + "</ol>";
      }
    }
  }
  return elem;
}
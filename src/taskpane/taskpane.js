/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

/***********************
Initialization Functions
***********************/


import { Data, params, templates, ERROR_MESSAGES } from './model'
import {
  parseArtifacts,
  loginCall,
  retrieveStyles,
  updateSelectionArray
} from './server'
//model stores user data (login credentials and projects)
var model = new Data();
//this is a global variable that expresses whether a user is using a version of word that supports api 1.3
var versionSupport;

//boilerplate initialization with a few calls
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    setDefaultDisplay();
    setEventListeners();
    document.body.classList.add('ms-office');
    //This checks if a users Word versions supports WordApi 1.3, then sets this in a global variable
    versionSupport = Office.context.requirements.isSetSupported("WordApi", "1.3")
    // this element doesn't currently exist
    // document.getElementById("help-connection-google").style.display = "none";
  }
});

//displays the login page / 'main' div after Office.onReady initialization
const setDefaultDisplay = () => {
  document.getElementById("app-body").style.display = "flex";
}

//sets the event listeners of buttons 
const setEventListeners = () => {
  let states = params.pageStates;
  // document.getElementById('test').onclick = () => test();
  document.getElementById('btn-login').onclick = async () => await loginAttempt();
  // document.getElementById('dev-mode').onclick = () => goToState(states.dev);
  document.getElementById('send-to-spira-button').onclick = async () => await pushArtifacts();
  document.getElementById('log-out').onclick = () => goToState(states.authentication);
  document.getElementById("select-requirements").onclick = () => openStyleMappings("req-");
  document.getElementById("select-test-cases").onclick = () => openStyleMappings("test-");
  document.getElementById("confirm-req-style-mappings").onclick = () => confirmStyleMappings('req-');
  document.getElementById("confirm-test-style-mappings").onclick = () => confirmStyleMappings('test-');
  document.getElementById('product-select').onchange = () => goToState(states.artifact);
  document.getElementById("pop-up-close").onclick = () => hideElement("pop-up");
  document.getElementById("pop-up-ok").onclick = () => hideElement('pop-up');
  document.getElementById("btn-help-login").onclick = () => goToState(states.helpLogin);
  document.getElementById("btn-help-main").onclick = () => goToState(states.helpMain);
  document.getElementById('lnk-help-login').onclick = () => goToState(states.helpLink);
  document.getElementById(params.buttonIds.helpLogin).onclick = () => openHelpSection(params.buttonIds.helpLogin);
  document.getElementById(params.buttonIds.helpModes).onclick = () => openHelpSection(params.buttonIds.helpModes);
  document.getElementById(params.buttonIds.helpVersions).onclick = () => openHelpSection(params.buttonIds.helpVersions);
  // event listener for pressing enter before login.
  addEventListener('keydown', async (e) => {
    //this only does anything if the login page is viewable. 
    if (e.code == "Enter" && !document.getElementById('panel-auth').classList.contains("hidden")) {
      await loginAttempt()
    }
    return
  })
}

/****************
Testing Functions 
*****************/
//basic testing function for validating code snippet behaviour.
async function test() {
  return Word.run(async (context) => {
    let body = context.document.body
    let lists = body.lists
    context.load(body)
    context.load(lists, ['paragraphs'])
    await context.sync();
    for (let list of lists.items) {
      console.log(list.paragraphs)
    }
  })
}

/**************
Spira API calls
**************/

//Attemps to 'log in' the user by making an api call to GET /projects
const loginAttempt = async () => {
  /*disable the login button to prevent someone from pressing it multiple times, this can
  overpopulate the products selector with duplicate sets.*/
  document.getElementById("btn-login").disabled = true
  //retrieves form data from input elements
  let url = document.getElementById("input-url").value
  let username = document.getElementById("input-username").value
  let rssToken = document.getElementById("input-password").value
  //allows user to enter URL with trailing slash or not.
  if (url[url.length - 1] == "/") {
    //url cannot be changed as it is tied to the HTML DOM input object, so creates a new variable
    var finalUrl = url.substring(0, url.length - 1)
  }
  //formatting the URL as it should be to populate products / validate user credentials
  let validatingURL = finalUrl || url + params.apiComponents.loginCall + `?username=${username}&api-key=${rssToken}`;
  try {
    //call the products API to populate relevant products
    var response = await loginCall(validatingURL);
    if (response.body) {
      //if successful response, move user to main screen
      hideElement('panel-auth');
      showElement('main-screen');
      document.getElementById('main-screen').style.display = 'flex';
      document.getElementById("btn-login").disabled = false
      model.user.url = finalUrl || url;
      model.user.username = username;
      model.user.api_key = rssToken;
      model.user.userCredentials = `?username=${username}&api-key=${rssToken}`;
      //populate the products dropdown & model object with the response body.
      populateProjects(response.body)
      boldStep('product-select-text');
      //On successful login, hide error message if its visible
      clearErrors();
      return
    }
  }
  catch (err) {
    //if the response throws an error, show an error message
    displayError(ERROR_MESSAGES.login);
    document.getElementById("btn-login").disabled = false;
    return
  }
  // Turn the header check box on here so that it won't get set just by navigating
  // around the send to spira menu
  document.getElementById('header-check').checked = true;
}

/********************
HTML DOM Manipulation
********************/

// these should be self explanatory
const disableButton = (htmlId) => {
  document.getElementById(htmlId).disabled = true
}

const enableButton = (htmlId) => {
  document.getElementById(htmlId).disabled = false
}

// disables all buttons on the main page
const disableMainButtons = () => {
  let buttons = Object.values(params.buttonIds);
  for (let button of buttons) {
    disableButton(button);
  }
}

// enables all buttons after sending succeeds or fails so the user do things again
const enableMainButtons = () => {
  let buttons = Object.values(params.buttonIds);
  for (let button of buttons) {
    enableButton(button);
  }
}

// Same as disableMainButtons, but for dropdowns
const disableDropdowns = () => {
  for (let i = 1; i < 6; i++) {
    // using disableButton but it works the same for dropdowns
    disableButton(`test-style-select${i}`);
  }
  for (let i = 1; i < 6; i++) {
    disableButton(`req-style-select${i}`);
  }
  disableButton('product-select');
}

// Same as enableMainButtons, but for dropdowns
const enableDropdowns = () => {
  for (let i = 1; i < 6; i++) {
    // using enableButton but it works the same for dropdowns
    enableButton(`test-style-select${i}`);
  }
  for (let i = 1; i < 6; i++) {
    enableButton(`req-style-select${i}`);
  }
  enableButton('product-select');
}

//populates the products selector with the retrieved products
const populateProjects = (projects) => {
  addDefaultProject();
  model.projects = projects
  let dropdown = document.getElementById('product-select')
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

//Reveals the style selectors after selecting an artifact.
const openStyleMappings = async (pageTag) => {
  clearErrors();
  hideElement('send-to-spira');
  //checks the current selected artifact type then loads the appropriate menu
  if (pageTag == "req-") {
    document.getElementById("select-requirements").classList.add("activated");
    document.getElementById("select-test-cases").classList.remove("activated");
    showElement('req-style-mappings');
    hideElement('test-style-mappings');
    boldStep('req-styles-text');
  }
  //opens the test cases style mappings if test mappings is the selected artifact type
  else {
    document.getElementById("select-test-cases").classList.add("activated");
    document.getElementById("select-requirements").classList.remove("activated");
    hideElement('req-style-mappings');
    showElement('test-style-mappings');
    boldStep('test-styles-text');
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

/* validates the users styles are valid from a parsing standpoint - not that it 
is valid for the particular document being used*/
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
      //gives an error explaining they have duplicate style mappings and cannot proceed.
      else {
        displayError(ERROR_MESSAGES.duplicateStyles)
        //hides the final button if it is already displayed when a user inputs invalid styles.
        hideElement('send-to-spira');
        boldStep(pageTag + 'styles-text');
        return
      }
      Office.context.document.settings.set(pageTag + 'style' + i.toString(), setting);
    }
    //gives an error explaining they have empty style mappings.
    else {
      displayError(ERROR_MESSAGES.emptyStyles)
      boldStep(pageTag + 'styles-text');
      return
    }
  }
  //hides error on successful confirm
  clearErrors();
  //this saves the settings
  Office.context.document.settings.saveAsync()
  //show the send to spira button after this is clicked and all style selectors are populated.
  showElement('send-to-spira');
  boldStep('send-to-spira-text');
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

//clears a passed in selector element by ID 
const clearDropdownElement = (element_id) => {
  let dropdown = document.getElementById(element_id);
  while (dropdown.length > 0) {
    dropdown.remove(0);
  }
}

//updates the progress bar to visually demonstrate the progress of sending the document.
const updateProgressBar = (current, total) => {
  let width = current / total * 100;
  let bar = document.getElementById("progress-bar-progress");
  bar.style.width = width + "%";
  if (current == total) {
    document.getElementById('pop-up').classList.remove('sending');
    document.getElementById('pop-up').classList.add('sent');
    document.getElementById('pop-up-text').textContent = `Sent ${total} 
    ${document.getElementById('select-requirements').classList.contains('activated') ?
        "Requirements" : "Test Cases"} successfully!`;
    enableButton('pop-up-ok');
  }
  else {
    document.getElementById('pop-up-text').textContent = `Sending ${total} 
    ${document.getElementById('select-requirements').classList.contains('activated') ?
        "Requirements!" : "Test Cases!"}`;
  }
  return true
}

//displays the progress bar, and removes any other errors that may have been displayed
const showProgressBar = () => {
  disableButton('pop-up-ok');
  showElement('pop-up');
  document.getElementById('pop-up').classList.remove('sent');
  document.getElementById('pop-up').classList.remove('err');
  document.getElementById('pop-up').classList.add('sending');
  document.getElementById('pop-up-text').textContent = "Parsing Document..."
  document.getElementById("progress-bar-progress").style.width = "0%";
  document.getElementById("progress-bar").classList.remove("hidden");
}

//hides the progress bar when sending finishes, and resets the bar for next use
const hideProgressBar = () => {
  document.getElementById('pop-up').classList.remove('sending');
  document.getElementById('pop-up').classList.remove('sent');
  document.getElementById("progress-bar-progress").style.width = "0%";
  hideElement('progress-bar');
}

//displays an error, where the error is an object from model.js's ERROR_MESSAGES
const displayError = (error, timeOut, failedArtifact) => {
  let element = document.getElementById(error.htmlId);
  hideProgressBar();
  enableButton('pop-up-ok');
  document.getElementById('pop-up').classList.add('err');
  showElement('pop-up')
  if (timeOut) {
    element.textContent = error.message;
    setTimeout(() => {
      clearErrors();
    }, ERROR_MESSAGES.stdTimeOut);
  }
  //special error case for handling hierarchy errors 
  else if (error.message.includes("hierarchy")) {
    element.textContent = error.message.replace("{hierarchy-line}", failedArtifact.Name)
  }
  else if (error.message.includes("table")) {
    element.textContent = error.message.replace("{table-line}", failedArtifact)
  }
  else if (failedArtifact) { // This is a special case error message for more descriptive errors when sending artifacts
    element.textContent =
      `The request to the API has failed on the Artifact: '${failedArtifact.Name}'. All, if any previous Artifacts should be in Spira.`;
  }
  else {
    element.textContent = error.message;
  }
}

//clears any errors currently being displayed
const clearErrors = () => {
  hideElement('pop-up');
  document.getElementById('pop-up').classList.remove('err');
  document.getElementById('pop-up-text').textContent = "";
}


// Takes you to a specific page or UI setup within the taskpane
const goToState = (state) => {
  let states = params.pageStates;
  switch (state) {

    // Case for logging out (back to authentication page)
    case (states.authentication):

      // clear stored user data
      model.clearUser();
      // hide main selection screen
      hideElement('main-screen');

      // removes currently entered RSS token to prevent a user from leaving their login credentials
      // populated after logging out and leaving their computer.
      document.getElementById("input-password").value = ""

      // Show authentication page
      showElement('panel-auth');
      clearDropdownElement('product-select');

      // Hides style mappings
      hideElement('req-style-mappings');
      hideElement('test-style-mappings');

      // Resets artifact button colors
      document.getElementById("select-requirements").classList.remove("activated");
      document.getElementById("select-test-cases").classList.remove("activated");

      // Hides artifact text and buttons
      hideElement('artifact-select-text');
      hideElement('select-requirements');
      hideElement('select-test-cases');

      // Hides send to spira menu and enables button
      hideElement('send-to-spira');
      hideProgressBar();
      document.getElementById("send-to-spira-button").disabled = false;

      // Clear any error messages that exist
      clearErrors();
      break;

    // Currently used to get back to the sending screen from help
    case (states.products):
      hideElement('panel-auth');
      showElement('main-screen');
      break;

    // Case for after a user selects a product and now needs to select artifact type
    case (states.artifact):
      // Show artifact text and buttons
      showElement('artifact-select-text');
      showElement('select-requirements');
      showElement('select-test-cases');
      boldStep('artifact-select-text');

      // If the blank product is listed, remove it from the dropdown.            // seven spaces
      if (document.getElementById('product-select').children.item(0).textContent == "       ") {
        document.getElementById('product-select').childNodes.item(0).remove();
      }
      break;
    case (states.dev):
      //moves us to the main interface without manually entering credentials
      hideElement('panel-auth');
      document.getElementById('main-screen').classList.remove('hidden');
      document.getElementById("main-screen").style.display = "flex";
      addDefaultProject();
      let devOption = document.createElement("option");
      devOption.text = "Test";
      devOption.value = null;
      document.getElementById("product-select").add(devOption);
      boldStep('product-select-text');

      // Turn the header check box on here so that it won't get set just by navigating
      // around the send to spira menu
      document.getElementById('header-check').checked = true;
      break;

    // moves us to the help screen and makes the back button take us to the login page
    case (states.helpLogin):
      hideElement('panel-auth');
      showElement('help-screen');
      document.getElementById('btn-help-back').onclick = () => {
        hideElement('help-screen');
        goToState(states.authentication);
      };

      // Automatically shows the login section when pressing help on the login screen.
      document.getElementById('btn-help-section-login').click();
      break;

    // moves us to the help screen and makes the back button take us to the project select page
    case (states.helpMain):
      hideElement('main-screen');
      showElement('help-screen');
      document.getElementById('btn-help-back').onclick = () => {
        hideElement('help-screen');
        goToState(states.products);
      }

      // Automatically shows the guide section when pressing help on the main screen.
      document.getElementById('btn-help-section-guide').click();
      break;


    // Used by the link in the login page header
    case (states.helpLink):
      hideElement('panel-auth');
      showElement('help-screen');
      document.getElementById('btn-help-back').onclick = () => {
        hideElement('help-screen');
        goToState(states.authentication);
      };
      break;

    // takes the user back to the start of the sending process after pushing their artifacts.
    case (states.postSend):

      // add back null project and select it
      addDefaultProject();
      document.getElementById('product-select').value = null;

      // Hides style mappings
      hideElement('req-style-mappings');
      hideElement('test-style-mappings');

      // Resets artifact button colors
      document.getElementById("select-requirements").classList.remove("activated");
      document.getElementById("select-test-cases").classList.remove("activated");

      // Hides artifact text and buttons
      hideElement('artifact-select-text');
      hideElement('select-requirements');
      hideElement('select-test-cases');

      // Hides send to spira menu and enables button
      hideElement('send-to-spira');

      // Bolds the step 1 text
      boldStep('product-select-text');
      break;

  }
}

/*buttonId is the id of the button pressed on the help page
 - this then opens the relevant help section */
const openHelpSection = (buttonId) => {
  for (let buttonId of params.collections.helpButtons) {
    document.getElementById(buttonId).classList.remove('activated')
  }
  let section = buttonId.replace("btn-", "")
  document.getElementById(buttonId).classList.add('activated')
  //hides the help sections before showing the relevant one
  for (let section of params.collections.helpSections) {
    hideElement(section)
  }
  showElement(section)
}

/*adds an empty "null" option to the product selector so it doesnt default to
the order Spira sends them in (prevents users from importing somewhere they 
  didn't mean to)*/
const addDefaultProject = () => {
  let nullProject = document.createElement("option");
  nullProject.text = "       ";
  nullProject.value = null;
  document.getElementById("product-select").add(nullProject, 0);
}

//helper function for hiding any HTML elemetn with .hidden
const hideElement = (elementId) => {
  document.getElementById(elementId).classList.add('hidden');
}

//helper function for revealing any HTML elemetn with .hidden
const showElement = (elementId) => {
  document.getElementById(elementId).classList.remove('hidden');
}

//bolds the current step the user is "on"
const boldStep = (stepId) => {
  for (let eachStep of params.collections.sendSteps) {
    document.getElementById(eachStep).classList.remove('bold');
  }
  document.getElementById(stepId).classList.add('bold');
}

/*********************
Pure data manipulation
**********************/

/*determines whether you want to parse requirements or test cases then disables all
buttons and dropdowns.*/
const pushArtifacts = async () => {
  let active = document.getElementById("test-style-mappings").classList.contains('hidden')
  //if the requirements style mappings are visible that is the selected artifact.
  let artifactType = params.artifactEnums.testCases;
  if (active) {
    artifactType = params.artifactEnums.requirements
  }
  disableMainButtons();
  disableDropdowns();
  await parseArtifacts(artifactType, model, versionSupport);
  enableMainButtons();
  enableDropdowns();
  goToState(params.pageStates.postSend);
}

/* Returns an array with all of the styles intended to be used for the 
  Style dropdown menus */
const getStyles = async () => {
  let userStyles = await usedStyles();
  let allStyles = await trimStyles(Object.values(Word.Style), userStyles)
  return allStyles;
}

/* Returns a new array by filtering out styles from the first set and adding 
them to the second set*/
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
  //this is an old method, but makes sense to use here as this occurs before the user 
  let selection = await updateSelectionArray();
  for (let i = 0; i < selection.length; i++) {
    //normal is a reserved style for descriptions of requirements
    if (!styles.includes(selection[i].style) && selection[i].style != "Normal") {
      styles.push(selection[i].style);
    }
  }
  return styles;
}

export {
  disableButton,
  displayError,
  clearErrors,
  showProgressBar,
  updateProgressBar,
  hideProgressBar,
  enableButton,
  enableMainButtons,
  disableMainButtons
}

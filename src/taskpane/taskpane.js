/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const axios = require('axios')

// Global selection array, used throughout
var SELECTION = [];
//setting a user object to maintain credentials when using other parts of the add-in
var USER_OBJ = { url: "", username: "", password: "" }

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("req-style-mappings").style.display = 'none';
    document.getElementById("test-style-mappings").style.display = 'none';
    document.getElementById("empty-error").style.display = 'none';
    setEventListeners();

    document.body.classList.add('ms-office');
    // this element doesn't currently exist
    // document.getElementById("help-connection-google").style.display = "none";
  }
});

const setEventListeners = () => {
  document.getElementById('test').onclick = test;
  document.getElementById('btn-login').onclick = () => loginAttempt();
  document.getElementById('dev-mode').onclick = () => devmode();
  document.getElementById('send-artifacts').onclick = () => pushRequirements();
  document.getElementById('log-out').onclick = () => logout();
  document.getElementById("style-mappings-button").onclick = () => openStyleMappings();
  //I think theres a way to use classes to reduce this to 2 but unsure
  document.getElementById("confirm-style-mappings").onclick = () => closeStyleMappings(true, 'req-');
  document.getElementById("cancel-style-mappings").onclick = () => closeStyleMappings(false, 'req-');
  document.getElementById("confirm-test-style-mappings").onclick = () => closeStyleMappings(true, 'test-');
  document.getElementById("cancel-test-style-mappings").onclick = () => closeStyleMappings(false, 'test-');
}

const devmode = () => {
  //moves us to the main interface without manually entering credentials
  document.getElementById('panel-auth').classList.add('hidden');
  document.getElementById('main-screen').classList.remove('hidden');
}

const loginAttempt = async () => {
  //retrieves form data from input elements
  let url = document.getElementById("input-url").value
  let username = document.getElementById("input-username").value
  let rssToken = document.getElementById("input-password").value
  //allows user to enter URL with trailing slash or not.
  let apiBase = "/services/v5_0/RestService.svc/projects"
  if (url[url.length - 1] == "/") {
    //url cannot be changed as it is tied to the HTML dom object, so creates a new variable
    var finalUrl = url.substring(0, url.length - 1)
  }
  //formatting the URL as it should be to populate projects / validate user credentials
  let validatingURL = finalUrl || url + apiBase + `?username=${username}&api-key=${rssToken}`;
  try {
    var response = await axios.get(validatingURL)

    if (response.data) {

      //if successful response, move user to main screen
      document.getElementById('panel-auth').classList.add('hidden');
      document.getElementById('main-screen').classList.remove('hidden');
      //save user credentials in global object to use in future requests
      USER_OBJ = {
        url: finalUrl || url, username: username, password: rssToken
      }
      populateProjects(response.data)

      //On successful login, remove error message if its visible
      document.getElementById("login-err-message").classList.add('hidden')
      return
    }
  }
  catch (err) {
    //if the response throws an error, show an error message for 5 seconds
    //In practice this can be more specific to alert the user to different potential problems
    document.getElementById("login-err-message").classList.remove('hidden');
    return
  }
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
  let settings = retrieveStyles(pageTag)
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
  //after this, select the relevant box when compared to the users settings
}

//closes the style mapping page taking in a boolean 'result'
//pageTag is req or test depending on which page is currently open

const closeStyleMappings = (result, pageTag) => {
  //result = true when a user selects confirm to exit the style mappings page
  if (result) {
    //saves the users style preferences. this is document bound
    for (let i = 1; i <= 5; i++) {
      let setting = document.getElementById(pageTag + "style-select" + i.toString()).value
      Office.context.document.settings.set(pageTag + 'style' + i.toString(), setting);
    }
    //this saves the settings
    Office.context.document.settings.saveAsync()
  }

  document.getElementById("main-screen").classList.remove("hidden")
  document.getElementById("req-style-mappings").style.display = 'none'
  document.getElementById("test-style-mappings").style.display = 'none'
  for (let i = 1; i <= 5; i++) {
    clearDropdownElement('style-select' + i.toString());
  }
}


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

//basic testing function for validating code snippet behaviour.
export async function test() {

  return Word.run(async (context) => {
    await updateSelectionArray();
    let lines = SELECTION;

    //try catch block for backend node call to prevent errors crashing the application
    try {
      let call1 = await axios.post("http://localhost:5000/retrieve", { lines: lines })
    }
    catch (err) {
      console.log(err)
    }
    // Tests the parseRequirements Function
    let requirements = parseRequirements(lines);

    //try catch block for backend node call to prevent errors crashing the application
    try {
      let call1 = await axios.post("http://localhost:5000/retrieve", { lines: lines, headings: requirements })
    }
    catch (err) {
      console.log(err)
    }
  })
}

// Get an Array of {text, style} objects from the user's selected text, delimited by /r
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

// Parses an array of range objects based on style and turns them into
// them into requirement objects
const parseRequirements = (lines) => {
  let requirements = []
  let page = document.getElementById("artifact-select").value;
  let styles;
  if (page == 'requirements') {
    styles = retrieveStyles('req-')
  }
  else {
    styles = retrieveStyles('test-')
  }
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
          requirements[requirements.length - 1].Description = requirements[requirements.length - 1].Description + line.text
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
//clears the credentials and returns the user to the home page
const logout = () => {
  var USER_OBJ = { url: "", username: "", password: "" };
  document.getElementById('panel-auth').classList.remove('hidden');
  document.getElementById('main-screen').classList.add('hidden');
  clearDropdownElement('project-select');
  //clears all 5 style select elements (titled style-select[1-5])
  for (let i = 1; i <= 5; i++) {
    clearDropdownElement('style-select' + i.toString());
  }
}

// Send a requirement to Spira using the API -- WIP
const pushRequirements = async () => {
  await updateSelectionArray();
  // Tests the parseRequirements Function
  let requirements = parseRequirements(SELECTION);
  /*if someone has selected an area with no properly formatted text, show an error explaining
  that and then return this function to prevent sending an empty request.*/
  if (requirements.length == 0) {
    document.getElementById("empty-error").textContent = "You currently have no valid text selected. if this isincorrect, check your style mappings and set them as the relevant styles."
    document.getElementById("empty-error").style.display = 'flex';
    setTimeout(() => {
      document.getElementById('empty-error').style.display = 'none';
    }, 8000)
    return
  }
  // Tests the pushRequirements Function
  let id = document.getElementById('project-select').value;
  for (let i = 0; i < requirements.length; i++) {
    let item = requirements[i];
    const apiCall = USER_OBJ.url + "/services/v5_0/RestService.svc/projects/" + id +
      `/requirements?username=${USER_OBJ.username}&api-key=${USER_OBJ.password}`;
    // try catch block to stop application crashing if call fails
    try {
      let call = await axios.post(apiCall, { Name: item.Name, Description: item.Description, RequirementTypeId: 2 });
    }
    catch (err) {
      //shows the failed requirement to add. This should work if it fails in the middle of sending
      document.getElementById("empty-error").textContent = `The request to the API has failed on requirement: '${item.Name}'. All, if any previous requirements should be in Spira.`
      document.getElementById("empty-error").style.display = "flex";
      setTimeout(() => {
        document.getElementById('empty-error').style.display = 'none';
      }, 8000)
    }
  }
  return
}

const clearDropdownElement = (element_id) => {
  let dropdown = document.getElementById(element_id);
  while (dropdown.length > 0) {
    dropdown.remove(0);
  }
}

const handleErrors = (error) => {
  //find status code in error report, then display a summary message
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
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const axios = require('axios')

//setting a user object to maintain credentials when using other parts of the add-in
const USER_OBJ = { url: "", username: "", password: "" }

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    setEventListeners();

    document.body.classList.add('ms-office');
    // this element doesn't currently exist
    // document.getElementById("help-connection-google").style.display = "none";
  }
});

const setEventListeners = () => {
  document.getElementById('test').onclick = test;
  document.getElementById('btn-login').onclick = () => loginAttempt();
}

const loginAttempt = async () => {
  //retrieves form data from input elements
  let url = document.getElementById("input-url").value
  let username = document.getElementById("input-username").value
  let rssToken = document.getElementById("input-password").value
  //allows user to enter URL with trailing slash or not.
  let slashCheck = "/services/v5_0/RestService.svc/projects"
  if (url[url.length - 1] == "/") {
    slashCheck = slashCheck.substring(1)
  }
  //formatting the URL as it should be to populate projects / validate user credentials
  let validatingURL = url + slashCheck + `?username=${username}&api-key=${rssToken}`;
  try {
    var response = await axios.get(validatingURL)
    if (response.status == 200) {
      //if successful response, move user to main screen
      document.getElementById('panel-auth').classList.add('hidden');
      document.getElementById('main-screen').classList.remove('hidden');
      //save user credentials in global object to use in future requests
      USER_OBJ = {
        url: url, username: username, password: rssToken
      }
      return
    }
  }
  catch (err) {
    //if the response throws an error, show an error message for 5 seconds
    //In practice this can be more specific to alert the user to different potential problems
    document.getElementById("login-err-message").classList.remove('hidden');
    setTimeout(() => {
      document.getElementById("login-err-message").classList.add('hidden')
    }, 5 * 1000)
    return
  }
}

//basic function which uses Word API to extract text as a proof of concept.
export async function test() {

  return Word.run(async (context) => {
    //check for highlighted text
    //splits the selected areas by enter-based indentation. 
    let selection = context.document.getSelection();
    context.load(selection, 'text');
    await context.sync();
    if (selection.text) {
      selection = context.document.getSelection().split(['/r'])
      context.load(selection, ['text', 'styleBuiltIn'])
      await context.sync();
    }

    // if nothing is selected, select the entire body of the document
    else {
      selection = context.document.body.getRange().split(['/r']);
      context.load(selection, ['text', 'styleBuiltIn'])
      await context.sync();
    }
 
    // Testing parsing lines of text from the selection array and logging it
    let lines = []
    selection.items.forEach((item) => {
      lines.push({ text: item.text, style: item.styleBuiltIn })
    })

    // Tests the parseRequirements Function
    let requirements = parseRequirements(lines);

    //try catch block for backend node call to prevent errors crashing the application
    try {
      let call1 = await axios.post("http://localhost:5000/retrieve", { lines: lines, headings: requirements })
      // let call = await axios.post("http://localhost:5000/retrieve", { text: selection.value})
    }
    catch (err) {
      console.log(err)
    }
  })
} 

// Parses an array of range objects based on style and turns them into
// them into requirement objects
const parseRequirements = (lines) => {
  let requirements = []
  for (let i = 0; i < lines.length; i++) { 
    if (lines[i].style === "Heading1") {
      if (lines[i + 1] && lines[i + 1].style === "Normal") {
        requirements.push({name: lines[i].text, description: lines[i + 1].text})
      }
      else {
        requirements.push({name: lines[i].text, description: null});
      }
    } 
  }
  return requirements;
}

const pushRequirement = async (name, description, projectId) => {
  const APICALL = USER_OBJ.url + "/projects/" + projectId + "/requirements";

}
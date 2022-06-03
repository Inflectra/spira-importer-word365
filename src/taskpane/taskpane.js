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
  //testing response to simulate successful/failed attemps

  try {
    var response = await axios.post("http://localhost:5000/retrieve", {
      url: url, username: username, password: rssToken
    })
    if (response.status == 200) {
      document.getElementById('panel-auth').classList.add('hidden');
      document.getElementById('main-screen').classList.remove('hidden');
      //save user credentials to use in future requests
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
  //if successful response, move user to main screen



}

//basic function which uses Word API to extract text as a proof of concept.
export async function test() {

  return Word.run(async (context) => {
    /*As it stands, these functions only retrieve plain text as
    a proof of concept. Later down the road, we will use cheerio or
    a similar solution to map these as HTML elements and create an 
    artifical DOM which can be navigated to retrieve relevant info
    while perserving stylization and structure. This will use .getHtml()
    after getSelection() to replace the context.load function*/
    //check for highlighted text
    let selection = context.document.getSelection();
    context.load(selection, 'text')
    await context.sync();

    //if nothing is selected, select the entire body of the document
    if (!selection.text) {
      selection = context.document.body;
      context.load(selection, 'text')
      await context.sync();
    }
    //try catch block for backend node call to prevent errors crashing
    try {
      let call1 = await axios.post("http://localhost:5000/retrieve", { text: selection.text })
      // let call = await axios.post("http://localhost:5000/retrieve", { text: selection.value})
    }
    catch (err) {
      console.log(err)
    }
  })
} 
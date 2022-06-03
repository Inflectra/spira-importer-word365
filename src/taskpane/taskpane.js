/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const axios = require('axios')

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById('test').onclick = test;
  }
});

export async function test() {

  return Word.run(async (context) => {
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
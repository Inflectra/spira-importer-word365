const axios = require('axios')
const superagent = require('superagent');
//ignore it saying defaults doesnt exist, it does and using default does not work.
//makes sure put requests uses the proper content-type header
axios.defaults.headers.put['Content-Type'] = "application/json"
axios.defaults.headers.put['accept'] = "application/json"

import { Data, tempDataStore, params, templates } from './model.js'
import { disableButton, retrieveStyles } from './taskpane.js';

var RETRIEVE = "http://localhost:5000/retrieve"

const sendArtifacts = async (ArtifactTypeId) => {
  return Word.run(async (context) => {
    disableButton("send-to-spira-button");
    /***************************
     Start range identification
     **************************/
    let projectId = new tempDataStore(document.getElementById("project-select").value)
    let selection = context.document.getSelection();
    let splitSelection = context.document.getSelection().split(['\r']);
    context.load(selection, ['text', 'inlinePictures', 'style', 'styleBuiltIn']);
    context.load(splitSelection, ['text', 'inlinePictures', 'style', 'styleBuiltIn'])
    await context.sync();
    await axios.post(RETRIEVE, { text: selection.text })
    //If there is no selected area, parse the full body of the document instead
    if (!selection.text) {
      selection = context.document.body.getRange();
      splitSelection = context.document.body.getRange().split(['\r']);
      context.load(selection, ['text', 'inlinePictures', 'style', 'styleBuiltIn']);
      context.load(splitSelection, ['text', 'inlinePictures', 'style', 'styleBuiltIn'])
      await context.sync();
    }
    /********************** 
     Start image formatting
    ***********************/
    let imageLines = []
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
    let imageObjects = [];
    let images = selection.inlinePictures;
    for (let i = 0; i < images.items.length; i++) {
      let base64 = images.items[i].getBase64ImageSrc();
      await context.sync();
      let imageObj = new templates.Image(base64.m_value, `inline${i}.jpg`, imageLines[0])
      imageObjects.push(imageObj)
      //removes the first item in imageLines as it has been converted to an imageObject;
      imageLines.shift();
    }
    //end of image formatting
    /**********************
     Start Artifact Parsing
    ***********************/
    switch (ArtifactTypeId) {
      //this will parse the data assuming the user wants artifacts to be imported.
      case (params.artifactEnums.requirements): {
        /*this returns the styles in an array to be referenced 
        against text in the document*/
        let styles = retrieveStyles('req-');
        let descStart, descEnd;
        let requirements = [];
        let requirement = new templates.Requirement()
        for (let [i, item] of splitSelection.items.entries()) {

        }
      }


    }
  })
}

//returns the Word API range object for the selected range or full body to be centrally stored
//A different version of this will be used for google docs.
const getUserSelection = async () => {
  return Word.run(async (context) => {

    return [selection, splitSelection]
  })
}


export {
  sendArtifacts,
  getUserSelection
}
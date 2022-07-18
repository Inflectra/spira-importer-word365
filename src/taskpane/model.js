
// Leaving room for google or word-specific setup later on in development.



// Main area for global constants, IE hardcoded values
var params = {
  //enums for different artifact types
  artifactEnums: {
    requirements: 1,
    testCases: 2,
    testSteps: 7
  },
  apiComponents: {
    loginCall: "/services/v6_0/RestService.svc/projects",
    apiBase: "/services/v6_0/RestService.svc/projects/",
    postOrPutRequirement: "/requirements",
    getRequirement: "/requirements/",
    postOrPutTestCase: "/test-cases",
    getTestCase: "/test-cases/",
    postOrPutTestStep: "/test-steps",
    getTestStep: "/test-steps/",
    postOrGetTestFolders: "/test-folders",
    postImage: "/documents/file",
    //these fields will be populated when the full URL is made for this
    imageSrc: "/{project-id}/Attachment/{AttachmentId}.aspx",
    //this is the initial outdent value for the first requirement sent (-20)
    initialOutdent: "/indent/-20",
    //have to replace {project-id} with tempDataStore.currentProjectId
    outdentCall: "/requirements/{requirement-id}/outdent",
    indentCall: "/requirements/{requirement-id}/indent",
  },
  regexs: {
    //this detects multiple whitespace characters that are not new lines in a row
    nonBreakingWhitespaceRegex: /(&nbsp;){2,}/g,
    whitespaceRegex: /([^\S\n]){2,}/g,
    tableRegex: /<table(.|\n|\r)*?\/table>/g,
    tableRowRegex: /<tr(.|\n|\r)*?\/tr>/g,
    tableDataRegex: /<td(.|\n|\r)*?\/td>/g,
    //this parses out the entire body and its contents
    bodyRegex: /<body(.|\n|\r|\s)*?<\/body>/gu,
    //this parses out the body tags for removal
    bodyTagRegex: /<(\/)??body(.|\n|\r|\s)*?>/gu,
    paragraphRegex: /(<p )(.|\n|\s|\r)*?(<\/p>)/gu,
    pTagRegex: /<\/??p(.|\n|\r|\s)*?>/g,
    emptyParagraphRegex: /<p(.)*?>\&nbsp\;<\/p>/g,
    orderedRegex: /.*class=MsoListParagraph.*><span.*>(.*)<span/,
    marginRegex: /style='margin-left:(\d)\.(\d)in/,
    //images in tables always have this \n at the beginning, but not the case everywhere.
    tableImageRegex: /<img\n(.|\n|\r)*?("|\s)>/gu,
    imageRegex: /<img(.|\n|\r)*?("|\s)>/gu,
    listSpanRegex: /<span(.|\r|\n|\s)*?(<\/span>){1,2}/,
    exceptedListRegex: />(\d{1} | \.){2,}<span/u,
    firstListItemRegex: /<p class=MsoListParagraphCxSpFirst(.|\n|\r)*?\/p>/g,
    lastListItemRegex: /<p class=MsoListParagraphCxSpLast(.|\n|\r)*?\/p>/g,
    singleListItemRegex: /<p class=MsoListParagraph (.|\n|\s|\r)*?<\/p>/g,
    orderedListRegex: />(\()?([A-Za-z0-9]){1,3}(\.|\))<span/,
    //this matches the ordered list 'icon' (ie. 1.,  a., 1) ) at the start of a line
    orderedListSymbolRegex: /^[A-Za-z0-9]{1,3}(\.|\))/,
    olTagRegex: /<ol>/g,
    olClosingTagRegex: /<\/ol>/g,
    ulTagRegex: /<ul>/g,
    ulClosingTagRegex: /<\/ul>/g
  },
  //this is the html id's of buttons which will be used when enabling or disabling buttons
  buttonIds: {
    sendToSpira: "send-to-spira-button",
    selectRequirements: "select-requirements",
    selectTestCases: "select-test-cases",
    validateStylesReq: "confirm-req-style-mappings",
    validateStylesTest: "confirm-test-style-mappings",
    logout: "log-out",
    help: "btn-help-main",
    helpLogin: "btn-help-section-login",
    helpModes: "btn-help-section-guide",
    helpVersions: "btn-help-section-issues"
  },
  // This is used to move around within the add-in UI
  pageStates: {
    authentication: 0,
    products: 1,
    artifact: 2,
    req_style: 3,
    test_style: 4,
    req_send: 5,
    postSend: 6,
    helpLogin: 7,
    helpMain: 8,
    helpLink: 9,
    dev: 42
  },
  //these are collections of parameters that are likely to be used as groups
  collections: {
    helpButtons: [
      "btn-help-section-login",
      "btn-help-section-guide",
      "btn-help-section-issues"
    ],
    helpSections: [
      "help-section-login",
      "help-section-guide",
      "help-section-issues"
    ],
    sendSteps: [
      "product-select-text",
      "artifact-select-text",
      "req-styles-text",
      "test-styles-text",
      "send-to-spira-text"
    ]
  }
}

var templates = {
  //Constructor functions for requirements and test cases
  Requirement: function () {
    this.Name = "";
    this.Description = "";
    this.RequirementTypeId = 2; // This is the requirement typeId we use when sending to Spira
    this.IndentLevel = 0;
  },

  TestCase: function () {
    this.folderName = "";
    this.folderDescription = "";
    this.Name = "";
    this.testCaseDescription = "";
    this.testSteps = [];
  },

  TestStep: function () {
    this.Description = "";
    this.ExpectedResult = "";
    this.SampleData = "";
  },

  Image: function (base64, name, lineNum) {
    this.base64 = base64;
    this.name = name;
    this.lineNum = lineNum;
  },

  ListItem: function (text, indentLevel) {
    this.text = text;
    this.indentLevel = indentLevel;
  }
}

/*
 *
 * ==========
 * DATA MODEL
 * ==========
 *
 * This holds all the information about the user and all template configuration information. 
 * Future versions of this program can add their artifact to the `templateFields` object.
 *
 */
function Data() {
  //global user object
  this.user = {
    url: '',
    username: '',
    api_key: '',
    //this will be populated on login.
    userCredentials: "?username={username}&api-key={api-key}"
  }

  //function to clear global user object for logout
  this.clearUser = () => {
    /*object desctructuring makes this not a reference value, new created Data object
    should be immediately discarded by trash collector*/
    this.user = { ...new Data().user }
  }

  this.projects = [];

  this.colors = {
    primaryButton: '#0078d7',
    selectedButton: '#0A0269',
    inactiveButton: '#f4f4f4',
    progressBarProgress: '#60ec60',
    progressBarBackground: '#dadada',
    errorMessages: "#ff0000"
  };
}

/*error messages and their potential location (used to be multiple)
 as well as other relevant information about them.*/
var ERROR_MESSAGES = {
  stdTimeOut: 8000, // 8000 is 8 seconds when used in setTimeout()
  allIds: { main: "pop-up-text" },
  login: { htmlId: "pop-up-text", message: "Your credentials are invalid" },
  empty: { htmlId: "pop-up-text", message: "You currently have no valid text selected or within the body of the document. if this is incorrect, check your style mappings and set them as the relevant styles." },
  hierarchy: { htmlId: "pop-up-text", message: "Your style hierarchy is invalid for the selected area. This occured on the line of '{hierarchy-line}'" },
  table: { htmlId: "pop-up-text", message: "Your table with first cell \"{table-line}\" does not contain any data in the column allocated for test step descriptions. Please update your selected column, or remove this table from your selection." },
  failedReq: { htmlId: "pop-up-text", message: "" },
  duplicateStyles: { htmlId: "pop-up-text", message: "You currently have multiple mappings set to the same style. Please only use each style once." },
  emptyStyles: { htmlId: "pop-up-text", message: "You currently have unselected styles. Please provide a style for all provided inputs." },
  testCaseFolders: { htmlId: "pop-up-text", message: "Loading or retriving test case folders failed. If retrieving failed, you may still have your test cases imported in a new folder." }
}

export { Data, params, templates, ERROR_MESSAGES }

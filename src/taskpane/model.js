
// Leaving room for google or word-specific setup later on in development.





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
        postOrPutTestStep: "/test-steps/",
        getTestStep: "/test-steps",
        postOrGetTestFolders: "/test-folders",
        postImage: "/documents/file",
    }
}

var templates = {
//Constructor functions for requirements and test cases
    Requirement:  function(){
        this.name = "";
        this.description = "";
        this.typeId = 2; // This is the requirement typeId we use when sending to Spira
        this.indentLevel = 0; 
    },

    TestCase: function() {
        this.folderName = "";
        this.folderDescription = "";
        this.name = "";
        this.testCaseDescription = "";
        this.testSteps = [];
    }
}
// Constructor function for globally accessible data that might change.
function Data() {

    this.user = {
        url: '',
        username: '',
        api_key: '',
        //this will be populated on login.
        userCredentials: "?username={username}&api-key={api-key}"
    }

    //function to clear user object for logout
    this.clearUser = () => {
        this.user = {
            url: '',
            username: '',
            api_key: '',
            //this will be populated on login.
            userCredentials: "?username={username}&api-key={api-key}"
        }
    }

    this.currentProjectId = ""
    this.projects = [];

    this.colors = {
        primaryButton: '#0078d7',
        selectedButton: '#0A0269',
        progressBarProgress: '#60ec60',
        progressBarBackground: '#dadada',
        errorMessages: "#ff0000"
    };
}

function tempDataStore() {
    this.currentProjectId = '';
}

var ERROR_MESSAGES = {
    stdTimeOut: 8000, // 8000 is 8 seconds when used in setTimeout()
    allIds: {login: "login-err", send: "send-err"},
    loginErr: {htmlId: "login-err", message: "Your credentials are invalid"},
    empty: {htmlId: "empty-err", message: "You currently have no valid text selected or within the body of the document. if this is incorrect, check your style mappings and set them as the relevant styles."},
    hierarchy: {htmlId: "hierarchy-err", message: "Your style heirarchy is invalid for the selected area. Please make sure requirements only indent 1 additional level from the previous requirement as specified in the indent level style selectors above."},
    table: {htmlId: "table-err", message: "Your description column for one or more tables only includes empty cells or does not exist. If you do not want to send test steps - do not select tables in your document. If you do, check your selection and try again."},
    duplicateStyles: {htmlId: "duplicate-styles-err", message: "You currently have multiple mappings set to the same style. Please only use each style once."},
    failedReq: {htmlId: "failed-req-err", message: ""},
    emptyStyles: {htmlId: "empty-styles-err", message: "You currently have unselected styles. Please provide a style for all provided inputs."}
}

export {Data, tempDataStore, params, templates, ERROR_MESSAGES}

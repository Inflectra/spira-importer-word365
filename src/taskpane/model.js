
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

export {Data, tempDataStore, params}
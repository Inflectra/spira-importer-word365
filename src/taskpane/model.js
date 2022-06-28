
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

    fieldType: {

    }
}

// Main area for object templates for artifacts and such
var templateFields = {

}

// Constructor function for globally accessible data that might change.
function Data() {

    this.user = {
        url: '',
        api_key: '',
    }

    this.projects = [];

}
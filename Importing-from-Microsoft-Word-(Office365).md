# Importing from Microsoft Word (Office 2019+, iOS, Web)

This add-in works with Microsoft Word 2019+, Word with Office 365, Word in the cloud (via a web browser), and Word on iPad. The add-in lets you import or export data to and from any product in your SpiraTest, SpiraTeam, or SpiraPlan application. 

**The add-in works for:**
1. Requirements
2. Test Cases with their Test Steps

## Important Notices
Older word documents may have images embedded in a fashion not consistent with what the Word JavaScript API supports. If your word document was initially created in a version of word from 2016 or earlier, you may want to upload your document to Google Docs, then re-download it as a .docx file. This will update the embedding of images to the newer method used. If this is not an option, see the [manual version](#dealing-with-images-within-old-documents) below. 

## Installation

To install the add-in:
* Go to the **insert** tab in Word.
* Click on **"Get Add-ins"** and in the window that opens, navigate to the **store** tab.
* Search for **"Spira"** or **"SpiraPlan"**.
* When you see the correct add-in developed by Inflectra, click on the "Add" button associated with it. 
* You should now see the SpiraPlan icon labeled "SpiraPlan Importer" in your home tab. Click on it to begin.

## 1. Connect to your Spira app

You can use this add-in with SpiraTest®, SpiraTeam®, or SpiraPlan®. 

If you are using Word in the browser, make sure your SpiraPlan is accessible over the internet.

![Spira add-in login screen](img/word365-log-in-screen.png)

* **Your Spira URL:** The web address that you use to access SpiraPlan® in your browser. Use the web address you use to access Spira in your browser. This is usually of the form 'http://**company**.spiraservice.net'. Make sure you remove any suffixes from the address (e.g. Default.aspx or "/")
* **Your Username:** This is the exact same username you use to log in to Spira. (Not Case Sensitive)
* **Enter your RSS token:** You can find or generate this from your user profile page inside Spira - "{ExampleRSS}". Make sure to include the curly braces and *make sure to hit Save after generating a new RSS token.*

**If there is a problem connecting to Spira you will be notified with an error message.**

After you have logged in click **Logout** to close your connection with Spira and take you back to the add-in's login page.

## 2. Select the product you wish to import to and the type of Artifact you wish to import

The add-in provides a dropdown with the various products within your Spira instance.

![Spira product selection screen](img/word365-product-select.png)

Once you have selected your product, a second option should appear giving you a choice between Requirements or Test Cases. This selection can be changed at any time by selecting the opposite option.

![Spira Artifact selection screen](img/word365-artifact-select.png)

## 3. Select the styles used in your document to represent the relevant fields

Select the styles used in the document which represent either the hierarchy of requirements, or the test case names, folder names, and table locations for test step properties. 

### Requirements:

For requirements, the 5 indent levels represent the hierarchical relationship of requirements in spira. An indent level 1 requirement is fully outdented, while indent level 2 is the child of the above indent level 1 requirement, and so on and so forth. You may only increase the indent level by 1 per requirement parsed, as the importer enforces Spira's hierarchy rules. Failure to follow this rule will throw a hierarchy error.

## Test Steps

For test steps, the first 2 style dropdowns represent the folder name, of which the test cases below it will be put in, as well as the style which will be treated as the title of test cases. The last 3 selectors are for determining where within your tables of test steps each field is located in. Every table within the selection must have at least 1 cell in the column for Descriptions with valid text within. If this is not the case, the importer will throw an error explaining that there is an invalid table. Please note, the importer will not create a test case folder with a name which already exists as a test case folder. Instead, the tool will import the test cases into the existing folder with that name. It is case sensitive, but can be a valuable tool if you wish to add the test cases from a document into an existing test case folder.

## 4. Select the portion of your Word document you wish to import to your Spira product

This import tool gives you the option to either parse an entire document (by not having any text highlighted) or parse only sections of the document (By highlighting that section of text before clicking "Send to Spira"). If you highlight a section of text which does not contain a style which would create either a test step or a requirement, there will be an invalid selection error which prevents any importing to Spira until you have made a valid selection. Make sure no lists within the selection contain styles which are set to any style selector - for instance if Heading 1 is your indent level 1 selection, lists may not contain any Heading 1 text - parsing will throw an error or even crash the add-in depending on the specific instance. 

Once you have selected the portion of the document you wish to import, click the "Send to Spira" button to start importing. The UI will be disabled while importing, and the progress bar will show the progress of imports. You may close this window, but it will not stop the import process. If you need an "Emergency stop" button, clicking the icon in the Word taskbar you used to open the add-in will refresh the add-in and stop the tool from continuing to import artifacts. 

## Dealing with images within old documents
If for some reason you cannot upload your Word document to the cloud with Google docs, you are still able to update the embedding type of images manually (1 by 1). This is done with one of two methods. First, by saving the image, then inserting it again in your modern version of word. The quicker, yet more complex approach, is cutting the image and pasting it again in the document. Then, after pasting, press ctrl (or click on the pop-up with a clickboard on it in the bottom right corner of your image) then pressing "U" or selecting the "Picture" option. By default, word will paste images with the setting "Keep source formatting", but this is precisely the source formatting from older documents which is incompatible with the Word Javascript API this tool uses. 


## Functionality Differences to the Microsoft Word Classic plugin

**What can the Word365 add-in do that the Classic Word add-in cannot?**

* Parse test step tables without removing the first row (the "Using header rows?" option allows you to toggle this)
* Enforce hierarchy rules before sending requirements
* Send Test cases / Requirements without requiring an empty last line under the rest of the selection
* Parse the full document (The classic add-in can only parse selected text)
* NOTE: It is compatible only with Word 2019+ (includes 365, Word on Web, Word for iPad) and Spira 6.3.0.1+

**What can the Classic Word add-in do that the Word365 add-in cannot?**

* Work with versions of spira older than 6.3.0.1
* Work with versions of Word 2016 or earlier

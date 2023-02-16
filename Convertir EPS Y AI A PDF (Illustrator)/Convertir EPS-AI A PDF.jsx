/** Saves every document found in a folder as a PDF file in a user specified folder.
    The documents to be converted must have the .eps or .ai extension.
*/

// Main Code [Execution of script begins here]

try {
    // uncomment to suppress Illustrator warning dialogs
    // app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    // Get the folder where the documents to be converted are located
    var sourceFolder = null;
    sourceFolder = Folder.selectDialog('Select folder with documents to convert.', '~');

    if (sourceFolder != null) {
        var destFolder = null;
        destFolder = Folder.selectDialog('Select folder for PDF files.', '~');

        if (destFolder != null) {
            var options, i, sourceFile, sourceDoc, targetFile;
            var files = sourceFolder.getFiles(/\.(ai|eps)$/i);

            if (files.length > 0) {
                // Get the PDF options to be used
                options = this.getOptions();
                // You can tune these by changing the code in the getOptions() function.

                for (i = 0; i < files.length; i++) {
                    sourceFile = files[i]; // returns the File object
                    sourceDoc = app.open(sourceFile); // returns the Document object

                    // Get the file to save the document as pdf into
                    targetFile = this.getTargetFile(sourceDoc.name, '.pdf', destFolder);

                    // Save as pdf
                    sourceDoc.saveAs(targetFile, options);
                    sourceDoc.close();
                }
                alert('Documents saved as PDF');
            } else {
                throw new Error('There are no documents to convert in the selected folder!');
            }
        }
    }
} catch (e) {
    alert(e.message, "Script Alert", true);
}

/** Returns the options to be used for the generated files.
    @return PDFSaveOptions object
*/
function getOptions() {
    // Create the required options object
    var options = new PDFSaveOptions();
    // See PDFSaveOptions in the JavaScript Reference for available options

    // Set the options you want below:

    // For example, uncomment to set the compatibility of the generated pdf to Acrobat 7 (PDF 1.6)
    // options.compatibility = PDFCompatibility.ACROBAT7;

    // For example, uncomment to view the pdfs in Acrobat after conversion
    // options.viewAfterSaving = true;

    return options;
}

/** Returns the file to save or export the document into.
    @param docName the name of the document
    @param ext the extension the file extension to be applied
    @param destFolder the output folder
    @return File object
*/
function getTargetFile(docName, ext, destFolder) {
    var newName = "";

    // if name has no dot (and hence no extension),
    // just append the extension
    if (docName.indexOf('.') < 0) {
        newName = docName + ext;
    } else {
        var dot = docName.lastIndexOf('.');
        newName += docName.substring(0, dot);
        newName += ext;
    }

    // Create the file object to save to
    var myFile = new File(destFolder + '/' + newName);

    // Preflight access rights
    if (myFile.open("w")) {
        myFile.close();
    } else {
        throw new Error('Access is denied');
    }
    return myFile;
}

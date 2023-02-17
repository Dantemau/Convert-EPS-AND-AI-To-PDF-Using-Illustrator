/** Saves every document found in a folder and its subfolders as a PDF file in the same folder as the source document.
    The documents to be converted must have the .eps or .ai extension.
*/

// Main Code [Execution of script begins here]
/** Given a source file, returns a file to save the document as pdf into.
    The pdf extension will replace the current extension (eps or ai).
    The file will be saved in the same folder as the source file.
    If the target file already exists, an error will be thrown.
    @param sourceFile The file to save as pdf.
    @param newExtension The new extension to use for the pdf file.
    @return A File object.
*/
function getTargetFile(sourceFile) {
    var targetFolder = sourceFile.parent;
    var sourceNameNoExt = sourceFile.name.slice(0, -4);
    var targetName = sourceNameNoExt + '.pdf';
    var targetFile = new File(targetFolder + '/' + targetName);

    // Check if the target file already exists, and overwrite it if it does
    if (targetFile.exists) {
        targetFile.remove();
    }

    return targetFile;
}


try {
    // uncomment to suppress Illustrator warning dialogs
    // app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    // Get the folder where the documents to be converted are located
    var sourceFolder = null;
    sourceFolder = Folder.selectDialog('Select folder with documents to convert.', '~');

    if (sourceFolder != null) {
        var options, sourceFile, sourceDoc, targetFile;
        var files = getFilesRecursive(sourceFolder, /\.(ai|eps)$/i);

        if (files.length > 0) {
            // Get the PDF options to be used
            options = this.getOptions();
            // You can tune these by changing the code in the getOptions() function.

            var totalFiles = files.length;
            var completedFiles = 0;

            for (var i = 0; i < totalFiles; i++) {
                sourceFile = files[i]; // returns the File object
                
                // Check if the file exists
                if (!sourceFile.exists) {
                    alert('File not found: ' + sourceFile.fsName);
                    continue;
                }
                
                sourceDoc = app.open(sourceFile); // returns the Document object

                // Get the file to save the document as pdf into
                targetFile = this.getTargetFile(sourceDoc.fullName);

                // Check if the target file already exists
                if (targetFile.exists) {
                    throw new Error('The file ' + targetFile.name + ' already exists in the same folder as the source document!');
                }

                // Save as pdf
                sourceDoc.saveAs(targetFile, options);
                sourceDoc.close();
                completedFiles++;

                // Update progress indicator
                var percentComplete = (completedFiles / totalFiles) * 100;
                $.writeln('Conversion progress: ' + percentComplete.toFixed(2) + '%');
            }

            alert("We're Done. Files were saved as PDF");
        } else {
            throw new Error('There are no documents to convert in the selected folder!');
        }
    }
} catch (e) {
    alert(e.message, "Script Alert", true);
}


/** Returns an array of files found in the specified folder and its subfolders that match the given regular expression pattern.
    @param folder The folder to search in.
    @param pattern The regular expression pattern to match file names against.
    @return An array of File objects.
*/
function getFilesRecursive(folder, pattern) {
    var fileList = [];

    var files = folder.getFiles();

    for (var i = 0; i < files.length; i++) {
        var file = files[i];

        if (file instanceof Folder) {
            // Recurse into subfolder
            fileList = fileList.concat(getFilesRecursive(file, pattern));
        } else if (file instanceof File && file.name.match(pattern)) {
            // Add matching file to list
            fileList.push(file);
        }
    }

    return fileList;
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

    // Use these settings for optimal compatibility with most PDF viewers
    return options;
}

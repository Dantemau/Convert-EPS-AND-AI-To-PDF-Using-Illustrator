/** Saves every document found in a folder and its subfolders as a PDF file in a user specified folder.
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
                    sourceDoc = app.open(sourceFile); // returns the Document object

                    // Get the file to save the document as pdf into
                    targetFile = this.getTargetFile(sourceDoc.name, '.pdf', destFolder);

                    // Check if the target file already exists
                    if (targetFile.exists) {
                        throw new Error('The file ' + targetFile.name + ' already exists in the destination folder!');
                    }

                    // Save as pdf
                    sourceDoc.saveAs(targetFile, options);
                    sourceDoc.close();

                    completedFiles++;

                    // Update progress indicator
                    var percentComplete = (completedFiles / totalFiles) * 100;
                    $.writeln('Conversion progress: ' + percentComplete.toFixed(2) + '%');
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

    // Use these settings for optimal output quality
    //options.pDFPreset = "PDF/X-1a:2001";
    //options.colorDownsampling = 300;
    //options.colorDownsamplingImageThreshold = 1;
    //options.colorDownsamplingMethod = DownsampleMethod.BICUBICDOWNSAMPLE;
    //options.grayscaleDownsampling = 300;
    //options.grayscaleDownsamplingImageThreshold = 1;
    //options.grayscaleDownsamplingMethod = DownsampleMethod.BICUBICDOWNSAMPLE;

    return options;
}

/**
 * Returns the target file for the document to be saved as PDF.
 * @param {string} docName The name of the source document.
 * @param {string} ext The extension to be used for the target file (e.g. '.pdf').
 * @param {Folder} destFolder The folder to save the target file into.
 * @return File object
 */
function getTargetFile(docName, ext, destFolder) {
    // Make sure the extension starts with a period
    if (ext.indexOf('.') != 0) {
        ext = '.' + ext;
    }

    // Create the file name
    var newName = docName.replace(/\.[^\.]+$/, '') + ext;

    // Create the target file object
    var targetFile = new File(destFolder + '/' + newName);

    return targetFile;
}

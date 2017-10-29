(function () {
    "use strict";

    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#url').html(location.href);
            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                // Do something that is only available via the new APIs
                $('#insert').click(insertImage);
                $('#getMathML').click(getMathMLfromSelection);
                $('#reload').click(document.getElementById("reloadForm").submit);
            } else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or greater.');
            }
        });
    };

    function insertImage() {
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to get the current selection.
            var range = thisDocument.getSelection();

            // Load image
            var request = new XMLHttpRequest();
            request.open('GET', 'formula.base64', false);
            request.send(null);
            document.getElementById("output").innerHTML = request.responseText;

            // Queue a command to replace the selected text.
            var image = range.insertInlinePictureFromBase64(request.responseText, Word.InsertLocation.replace);
            image.altTextTitle = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi><mo>=</mo><mfrac><mn>1</mn><mn>2</mn></mfrac></math>';

            // Synchronize the document state by executing the queued commands.
            return context.sync().then(function () {
                console.log('Added an image.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + error.message);
        });
    }

    function getMathMLfromSelection() {
        Word.run(function (context) {
            // Clear the output
            document.getElementById("output").value = "";

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to get the current selection.
            var range = thisDocument.getSelection();
            var html = range.getHtml();
            
            // Synchronize the document state by executing the queued commands.
            return context.sync().then(function () {
                var v=html.value;
                var re = /<math(.*)math>'/ig;
                var r = re.exec(v);
                document.getElementById("output").value = "<math"+r[1]+"math>";
            });
        })
        .catch(function (error) {
            console.log('Error: ' + error.message);
        });
    }
    
})();
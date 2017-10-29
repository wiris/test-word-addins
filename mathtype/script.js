(function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#url').html(location.href);
                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#insertOoxml').click(insertFormula);
                    $('#getSelection').click(getSelection);
                    $('#getOoxml').click(getSelectionOoxml);
                    $('#reload').click(document.getElementById("reloadForm").submit);
                } else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertFormula() {
            Word.run(function (context) {
                console.log('Begin adding an image.');

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();
				
                var request = new XMLHttpRequest();
                request.open('GET', 'formula.xml', false);
                request.send(null);
                document.getElementById("output").value = request.responseText;

                // Queue a command to replace the selected text.
                var image = range.insertOoxml(request.responseText, Word.InsertLocation.replace);

                // Synchronize the document and attach MathML to ContentControl
                return context.sync().then(function () {
                    console.log('Added an image.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + error.message);
            });
        }

        function getSelection() {
            Word.run(function (context) {
                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                var range = thisDocument.getSelection();
                var html = range.getHtml();

                // Synchronize the document state by executing the queued commands.
                return context.sync().then(function () {
                    document.getElementById("output").value = html.value;
                });
            })
            .catch(function (error) {
                console.log('Error: ' + error.message);
            });
            
        }

        function getSelectionOoxml() {
            Word.run(function (context) {
                // Clear value
                document.getElementById("output").value = "";

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                var range = thisDocument.getSelection();
                var xml = range.getOoxml();

                // Synchronize the document state by executing the queued commands.
                return context.sync().then(function () {
                    document.getElementById("output").value = xml.value;
                });
            })
            .catch(function (error) {
                console.log('Error: ' + error.message);
            });
            
        }

})();
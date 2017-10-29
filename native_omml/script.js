(function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#url').html(location.href);
                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#insert').click(insertFormula);
                    $('#getOmml').click(getSelectionOmml);
                    $('#reload').click(document.getElementById("reloadForm").submit);
                } else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertFormula() {
            Word.run(function (context) {

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

        function getSelectionOmml() {
            Word.run(function (context) {
                // Clear value
                document.getElementById("output").value = "";

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();
                var html = range.getHtml();

                return context.sync().then(function () {
                    var value = html.value;
                    // Find text between <m:oMath> and </m:oMath>
                    var i0 = Math.max(value.indexOf("&lt;m:oMath "),value.indexOf("&lt;m:oMath&gt;"));
                    var end = "&lt;/m:oMath&gt;";
                    var i1 = value.indexOf(end);
                    if (i0>=0 && i1>=0) {
                        i1 += end.length;
                        var escaped = value.substring(i0,i1);
                        escaped = escaped.replaceAll("&lt;","<");
                        escaped = escaped.replaceAll("&gt;",">");
                        document.getElementById("output").value = escaped;
                    }
                });
            })
            .catch(function (error) {
                console.log('Error: ' + error.message);
            });
            
        }

        String.prototype.replaceAll = function(find, replace) {
            var str = this;
            return str.replace(new RegExp(find.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1"), 'g'), replace);
        };
})();
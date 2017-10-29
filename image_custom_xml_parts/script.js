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
                    $('#getMathML').click(getMathMLfromSelection);
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
                request.open('GET', 'formula.base64', false);
                request.send(null);
                document.getElementById("output").innerHTML = request.responseText;

                // Insert ContentControl
                var myContentControl = range.insertContentControl();

                // Add image to ContentControl
                var image = myContentControl.insertInlinePictureFromBase64(request.responseText, Word.InsertLocation.replace);

                // Synchronize the document and attach MathML to ContentControl
                return context.sync().then(function () {
                    console.log('Added an image.');
                    var mathml = '<math xmlns="http://www.w3.org/1998/Math/MathML" id="'+myContentControl.id+'"><mi>x</mi><mo>=</mo><mfrac><mn>1</mn><mn>2</mn></mfrac></math>';

                    var customXmlParts = thisDocument.customXmlParts;
                    if (typeof customXmlParts == 'undefined') {
                        console.log("customXmlParts is undefined using Office.context.document.customXmlParts")
                        customXmlParts = Office.context.document.customXmlParts;
                    }
                    customXmlParts.addAsync(mathml,
                        function() {
                            console.log("Added mathml.");
                        });
                    });
            })
            .catch(function (error) {
                console.log('Error: ' + error.message);
            });
        }

        function getMathMLfromSelection() {
            Word.run(function (context) {
                
                // Clear value
                document.getElementById("output").value = "";

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                var range = thisDocument.getSelection();
                var contentControl = range.parentContentControl;
                
                // Synchronize the document state by executing the queued commands.
                return context.sync().then(function () {
                    console.log(contentControl.id);
                    
                    var customXmlParts = thisDocument.customXmlParts;
                    if (typeof customXmlParts == 'undefined') {
                        console.log("customXmlParts is undefined using Office.context.document.customXmlParts")
                        customXmlParts = Office.context.document.customXmlParts;
                    }

                    // Get the list of all MathML parts using the MathML namespace
                    customXmlParts.getByNamespaceAsync("http://www.w3.org/1998/Math/MathML", function (eventArgs) {
                        console.log("Found " + eventArgs.value.length + " parts with this namespace");
                        for (var i in eventArgs.value) {
                            eventArgs.value[i].getXmlAsync(function (asyncResult) {
                                // Get the MathML and extract the id
                                var mathml = asyncResult.value;
                                console.log(mathml);
                                var re = /id="([^"]*)"/ig;
                                var r = re.exec(mathml);
                                // Is the desired MathML ?
                                if (r[1]==contentControl.id) {
                                    document.getElementById("output").value = mathml;
                                }
                            });
                        }
                    });
                }); 
            })
            .catch(function (error) {
                console.log('Error: ' + error.message);
            });
        }
        
        })();
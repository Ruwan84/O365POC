var xhr;
var serviceRequest;

(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            serviceRequest = new Object();
            serviceRequest.ArtifactType = "docx";
            serviceRequest.ContentType = "Application/docx";
            serviceRequest.Content = "";
            serviceRequest.Id = 0;
            serviceRequest.IsInline = true;
            serviceRequest.Name = "";
            serviceRequest.Size = 0;

            getFileName();
            document.getElementById('submit').addEventListener("click",
                function () {
                    loadingStatus();
                    sendFile();
                });



        });
    };   

    function loadingStatus() {
        var statusInfo = document.getElementById("status");
        statusInfo.innerHTML = '<img src="https://nylo365addinsonazurev2.azurewebsites.net/Images/spinner.gif">';
        document.getElementById('status').scrollIntoView();
    }

    // Create a function for writing to the status div. 
    function updateStatus(message) {
        var statusInfo = document.getElementById("status");
        statusInfo.innerHTML = '<div class="alert alert-success float-left">' + message + '</br>';
    }

    // Get all the content from a PowerPoint or Word document in 100-KB chunks of text.
    function sendFile() {
        Office.context.document.getFileAsync("compressed",
            { sliceSize: 100000 },
            function (result) {

                if (result.status == Office.AsyncResultStatus.Succeeded) {

                    // Get the File object from the result.
                    var myFile = result.value;
                    var state = {
                        file: myFile,
                        counter: 0,
                        sliceCount: myFile.sliceCount
                    };

                    //updateStatus("Getting file of " + myFile.size + " bytes");

                    getSlice(state);
                }
                else {
                    updateStatus(result.status);
                }
            });
    }

    function getFileName() {

        Office.context.document.getFilePropertiesAsync(function (asyncResult) {
            var fileUrl = asyncResult.value.url;
            if (fileUrl == "") {
                //var fileName = document.getElementById("fileName");
                //fileName.innerHTML = "Document1.docx";                
                document.getElementById("TextBoxTitle").value = "Document1.docx";
            }
            else {
                var filePath = fileUrl.split("/");
                //var fileName = document.getElementById("fileName");
                //fileName.innerHTML = filePath[filePath.length - 1];
                document.getElementById("TextBoxTitle").value = filePath[filePath.length - 1];
            }            
        });
    }

    // Get a slice from the file and then call sendSlice.
    function getSlice(state) {
        state.file.getSliceAsync(state.counter, function (result) {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                //updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
                sendSlice(result.value, state);
            }
            else {
                updateStatus(result.status);
            }
        });
    }

    ////function sendSlice1(slice, state) {
    ////    var data = slice.data;

    ////    // If the slice contains data, create an HTTP request.
    ////    if (data) {

    ////        // Encode the slice data, a byte array, as a Base64 string.
    ////        // NOTE: The implementation of myEncodeBase64(input) function isn't 
    ////        // included with this example. For information about Base64 encoding with
    ////        // JavaScript, see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Base64_encoding_and_decoding.
    ////        //var fileData = myEncodeBase64(data);
    ////        var fileData = data;

    ////        // Create a new HTTP request. You need to send the request 
    ////        // to a webpage that can receive a post.
    ////        var request = new XMLHttpRequest();

    ////        // Create a handler function to update the status 
    ////        // when the request has been sent.
    ////        request.onreadystatechange = function () {
    ////            if (request.readyState == 4) {

    ////                updateStatus("Sent " + slice.size + " bytes.");
    ////                state.counter++;

    ////                if (state.counter < state.sliceCount) {
    ////                    getSlice(state);
    ////                }
    ////                else {
    ////                    closeFile(state);
    ////                }
    ////            }
    ////        }

    ////        //request.open("POST", "https://localhost:44320/api/AttachmentService");
    ////        request.open("POST", "https://localhost:44355/default.aspx");


    ////        request.setRequestHeader("Slice-Number", slice.index);

    ////        // Send the file as the body of an HTTP POST 
    ////        // request to the web server.
    ////        request.send(fileData);
    ////    }
    ////}

    //function makeServiceRequest() {
    function sendSlice(slice, state) {
        var data = slice.data;
        // If the slice contains data, create an HTTP request.
        if (data) {
            var fileData = base64js.fromByteArray(data);

            xhr = new XMLHttpRequest();

            // Update the URL to point to your service location.
            //xhr.open("POST", "https://nylo365webapionazure.azurewebsites.net/api/WordService", true);
            //xhr.open("POST", "https://localhost:44313/api/WordService", true);
            xhr.open("POST", "https://nylo365webapionazurev2.azurewebsites.net/api/WordService", true);
            
            xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            xhr.onreadystatechange = requestReadyStateChange;

            //var fileName = document.getElementById("fileName").innerHTML;
            serviceRequest.ArtifactType = "docx";
            serviceRequest.ContentType = "Application/docx";
            serviceRequest.Content = fileData;            
            serviceRequest.IsInline = true;
            serviceRequest.Size = 0;
            //serviceRequest.Name = fileName;
            //serviceRequest.Id = slice;            
            //serviceRequest.Attachment = slice._data$p$0 || slice.$0_0;
            serviceRequest.Name = $('#TextBoxTitle').val();
            //serviceRequest.Title = $('#TextBoxTitle').val();
            serviceRequest.Function = $('#DropDownListFuncations option:selected').text() + '|' + $('#DropDownListFuncations option:selected').val();
            serviceRequest.DocumentType = $('#DropDownListDocumentTypes option:selected').text() + '|' + $('#DropDownListDocumentTypes option:selected').val();
            serviceRequest.LineOfBusiness = $('#DropDownListLineofBusiness option:selected').text() + '|' + $('#DropDownListLineofBusiness option:selected').val();
            serviceRequest.BusinessArea = $('#DropDownListBusinessAreas option:selected').text() + '|' + $('#DropDownListBusinessAreas option:selected').val();
            serviceRequest.SubBusinessArea = $('#DropDownListSubBusinessAreas option:selected').text() + '|' + $('#DropDownListSubBusinessAreas option:selected').val();
            serviceRequest.SubFunction = $('#DropDownListSubFunction option:selected').text() + '|' + $('#DropDownListSubFunction option:selected').val();
            serviceRequest.Tower = $('#DropDownListTower option:selected').text();
            serviceRequest.SubTower = $('#TextBoxSubTower').val();
            serviceRequest.Application = $('#DropDownAppName option:selected').text();
            serviceRequest.Project = $('#DropDownProjectName option:selected').text();
            serviceRequest.ExpiryDate = $('#fileUploadDatePikerExpDate').val();
            serviceRequest.Keyword = $('#TextBoxKeyword').val();
            serviceRequest.Comments = $('#FileUploadVersionComments').val();

            //console.log('Name - ' + serviceRequest.Name);            
            //console.log('Function - ' + serviceRequest.Function);
            //console.log('DocumentType - ' + serviceRequest.DocumentType);
            //console.log('LineOfBusiness - ' + serviceRequest.LineOfBusiness);
            //console.log('BusinessArea - ' + serviceRequest.BusinessArea);
            //console.log('SubFunction - ' + serviceRequest.SubFunction);
            //console.log('Tower - ' + serviceRequest.Tower);
            //console.log('SubTower - ' + serviceRequest.SubTower);
            //console.log('Application - ' + serviceRequest.Application);
            //console.log('Project - ' + serviceRequest.Project);
            //console.log('ExpiryDate - ' + serviceRequest.ExpiryDate);
            //console.log('Keyword - ' + serviceRequest.Keyword);
            //console.log('Comments - ' + serviceRequest.Comments);
            

            // requestReadyStateChange function.
            xhr.send(JSON.stringify(serviceRequest));
        }
    };

    // Handles the response from the JSON web service.
    function requestReadyStateChange() {
        if (xhr.readyState == 4) {
            if (xhr.status == 200) {
                var response = JSON.parse(xhr.responseText);
                console.log(response.Message);
                updateStatus(response.Message);
            }
        }
    };

    function closeFile(state) {

        // Close the file when you're done with it.
        state.file.closeAsync(function (result) {

            // If the result returns as a success, the
            // file has been successfully closed.
            if (result.status == "succeeded") {
                updateStatus("File closed.");
            }
            else {
                updateStatus("File couldn't be closed.");
            }
        });
    }

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(errorHandler);
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();

function showHideView() {
    $("#view-all").toggleClass('collapse');

    var viewText = $(".view-all-link").text();
    $(".view-all-link").text((viewText === 'View More') ? 'Hide' : 'View More');
}

var xhr;
var serviceRequest;

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            initApp();
        });
    };

    function initApp() {
        
        if (Office.context.mailbox.item.attachments == undefined) {


        } else if (Office.context.mailbox.item.attachments.length == 0) {


        } else {

            var attachmentCountMessage = Office.context.mailbox.item.attachments.length + " attachment(s) availale in this email. Click on Submit button to send them to the KMP";
            document.getElementById("attachments").innerHTML = attachmentCountMessage;

            serviceRequest = new Object();
            serviceRequest.attachmentToken = "";
            serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            serviceRequest.attachments = new Array();
        }
    };

    

})();

function submitAttachments() {
    Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
};

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status == "succeeded") {
        serviceRequest.attachmentToken = asyncResult.value;
        makeServiceRequest();
    }
    else {
        showMessage("Error occured when retrieving the access token.");
    }
}

function makeServiceRequest() {
    var attachment;
    xhr = new XMLHttpRequest();

    // Update the URL to point to your service location.
    xhr.open("POST", "https://nylo365webapionazure.azurewebsites.net/api/OutlookService", true);

    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    xhr.onreadystatechange = requestReadyStateChange;

    // Translate the attachment details into a form easily understood by WCF.
    for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
        attachment = Office.context.mailbox.item.attachments[i];
        attachment = attachment._data$p$0 || attachment.$0_0;

        if (attachment !== undefined) {
            serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
        }
    }

    // Send the request. The response is handled in the 
    // requestReadyStateChange function.
    xhr.send(JSON.stringify(serviceRequest));
};


// Handles the response from the JSON web service.
function requestReadyStateChange() {
    if (xhr.readyState == 4) {
        if (xhr.status == 200) {
            var response = JSON.parse(xhr.responseText);
            if (!response.isError) {
                // The response indicates that the server recognized
                // the client identity and processed the request.
                // Show the response.
                var names = "<h2>Below attachments are uploaded to the KMP: </h2><br />";

                for (i = 0; i < response.AttachmentNames.length; i++) {
                    names += response.AttachmentNames[i] + "<br />";
                }
                showMessage(names); 
            } else {
                showMessage(response.message); 
            }
        } else {
            if (xhr.status == 404) {
                showMessage("The app server could not be found.");
            } else {
                showMessage("There was an unexpected error: " + xhr.status + " -- " + xhr.statusText);
            }
        }
    }
};

// Shows the service response.
function showResponse(response) {
    document.getElementById("response").innerHTML = "Submittde attachments: " + response.attachmentsProcessed;
};

function showMessage(message) {
    document.getElementById("message").innerHTML = message;
};



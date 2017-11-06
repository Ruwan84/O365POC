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

            for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
                var itemName = Office.context.mailbox.item.attachments[i].name;
                $('#attachments').append('<input class="checkbox-1" type="checkbox" name="myCheckbox" /><label>' + itemName + '</label><br />');
            }
            console.log('NEW CONTENT');
            showMessage('<div><b>Please select attachments for uploading to KM Portal</b></div>');

            //serviceRequest = new Object();
            //serviceRequest.attachmentToken = "";
            //serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            //serviceRequest.attachments = new Array();
        }
    };



})();

function showHideView() {
    $("#view-all").toggleClass('collapse');

    var viewText = $(".view-all-link").text();
    $(".view-all-link").text((viewText === 'View More') ? 'Hide' : 'View More');
}

function submitAttachments() {

    serviceRequest = new Object();
    serviceRequest.attachmentToken = "";
    serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
    serviceRequest.attachments = new Array();

    // Clear existing message if exists
    emptyMessage();

    if (isItemsSelected() == true) {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
    else {
        showMessage('<div class="alert alert-danger">Please select at least One attachment for uploading to KM Portal</div>');
    }
};

function isItemsSelected() {
    var isSelected = false;
    var selected = [];
    var n = $('#attachments input:checked').length;
    isSelected = (n === 0 ? false : true);

    return isSelected;
}

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

    var selected = [];
    $('#attachments input:checked').each(function () {
        selected.push($(this).next('label').text());
    });

    // Update the URL to point to your service location.
    //xhr.open("POST", "https://nylo365webapionazure.azurewebsites.net/api/OutlookService", true);
    //xhr.open("POST", "https://localhost:44313/api/OutlookService", true);
    xhr.open("POST", "https://nylo365webapionazurev2.azurewebsites.net/api/OutlookService", true);

    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    xhr.onreadystatechange = requestReadyStateChange;

    // Translate the attachment details into a form easily understood by WCF.
    var attachmentCount = 0;
    for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {

        attachment = Office.context.mailbox.item.attachments[i];
        var itemName = attachment.name;

        if (jQuery.inArray(itemName, selected) > -1) {
            attachment = attachment._data$p$0 || attachment.$0_0;
            attachment.Function = $('#DropDownListFuncations option:selected').text() + '|' + $('#DropDownListFuncations option:selected').val();
            attachment.DocumentType = $('#DropDownListDocumentTypes option:selected').text() + '|' + $('#DropDownListDocumentTypes option:selected').val();
            attachment.LineOfBusiness = $('#DropDownListLineofBusiness option:selected').text() + '|' + $('#DropDownListLineofBusiness option:selected').val();
            attachment.BusinessArea = $('#DropDownListBusinessAreas option:selected').text() + '|' + $('#DropDownListBusinessAreas option:selected').val();
            attachment.SubBusinessArea = $('#DropDownListSubBusinessAreas option:selected').text() + '|' + $('#DropDownListSubBusinessAreas option:selected').val();
            attachment.SubFunction = $('#DropDownListSubFunction option:selected').text() + '|' + $('#DropDownListSubFunction option:selected').val();
            attachment.Tower = $('#DropDownListTower option:selected').text();
            attachment.SubTower = $('#TextBoxSubTower').val();
            attachment.Application = $('#DropDownAppName option:selected').text();
            attachment.Project = $('#DropDownProjectName option:selected').text();
            attachment.ExpiryDate = $('#fileUploadDatePikerExpDate').val();
            attachment.Keyword = $('#TextBoxKeyword').val();
            attachment.Comments = $('#FileUploadVersionComments').val();

            if (attachment !== undefined) {
                serviceRequest.attachments[attachmentCount] = JSON.parse(JSON.stringify(attachment));
                attachmentCount++;
                console.log(attachmentCount);
            }
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
                var attachmentInfoList = "";

                for (i = 0; i < response.AttachmentNames.length; i++) {
                    attachmentInfoList += '<span class="doc-style"><b>' + response.AttachmentNames[i] + '</b></span><br/>';
                }
                var names = '<div class="alert alert-success float-left">Following attachments are uploaded to the KM Portal.</br>' + attachmentInfoList + '</div>';
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
    document.getElementById("response").innerHTML = "Submitted attachments: " + response.attachmentsProcessed;
};

function showMessage(message) {
    document.getElementById("message").innerHTML = message;
};

function emptyMessage() {
    document.getElementById("message").innerHTML = '';
    document.getElementById("message").innerHTML = '<div class="text-align-center"><img src="https://nylo365addinsonazurev2.azurewebsites.net/Images/spinner.gif"></div>';
    document.getElementById('message').scrollIntoView();
};



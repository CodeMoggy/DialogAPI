
var dlg;


(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            
        });
    };

})();

//The MIT License (MIT)

//Copyright (c) Microsoft Corporation

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

function getContacts(event) {

    // determine if the version of Excel supports the required requirementSet
    if (Office.context.requirements.isSetSupported('DialogAPI', '1.1')) {

        // defines which page to open when the dialog is launched
        var url = "https://localhost:44370/app/auth.html";

        Office.context.ui.displayDialogAsync(url, { height: 40, width: 40, requireHTTPS: true }, function (result) {
            dlg = result.value;
            
            // add an event handler when the dialog message has been received
            dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, dialogMessageReceived);
        });

        event.completed()
    }
    else {
        // todo - need an alternative path
    }
}


function dialogMessageReceived(result) {
    if (result && JSON.parse(result.message).status === "success") {


        //close the dialog and call into Graph
        dlg.close();

        // grab the token from the result...this is the AAD access token that will be used to authenticate the callout to the Graph API
        var _token = JSON.parse(result.message);

        $.ajax({
            url: "https://graph.microsoft.com/v1.0/me/contacts",
            headers: {
                "accept": "application/json",
                "Authorization": "Bearer " + _token.accessToken
            },
            success: function (data) {
                var _officeTable = new Office.TableData();
                _officeTable.headers = ["First", "Last", "PrimaryEmail", "Company", "WorkPhone", "MobilePhone"];

                $(data.value).each(function (i, e) {
                    var _item = [
                        e.givenName,
                        e.surname,
                        (e.emailAddresses.length > 0) ? e.emailAddresses[0].address : "",
                        (e.companyName) ? e.companyName : "",
                        (e.businessPhones.length > 0) ? e.businessPhones[0] : "",
                        (e.mobilePhone) ? e.mobilePhone : ""];
                    _officeTable.rows.push(_item);
                });

                Office.context.document.setSelectedDataAsync(_officeTable, {
                    coercionType: Office.CoercionType.Table,
                    cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }]
                }, function (asyncResult) {
                    if (asyncResult.status !== Office.AsyncResultStatus.Failed) {
                        //create a table binding
                        Office.context.document.bindings.addFromSelectionAsync(
                            Office.BindingType.Table, { id: "ContactsBinding" },
                            function (result) {
                                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                                    // todo - add error handling
                                }
                                else {
                                    //get the binding
                                    Office.context.document.bindings.getByIdAsync("ContactsBinding", function (result) {
                                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                                            //add BindingDataChanged handler
                                            var _binding = result.value;
                                            _binding.addHandlerAsync(Office.EventType.BindingDataChanged, function () {
                                                // for the purpose of this project no event handler is required
                                            });
                                        }
                                    });
                                }
                            });
                    }
                    else {
                        // todo
                    }
                });
            },
            error: function (err) {
                var e = err;
            }
        });
    }
    else {
        // todo - add error handling
    }
}
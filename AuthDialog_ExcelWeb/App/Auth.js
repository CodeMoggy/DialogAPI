﻿//The MIT License (MIT)

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

(function () {
    "use strict";


    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(signIn);
    };

    function signIn() {
        var response = { "status": "none", "accessToken": "" };

        window.config = {
            instance: "https://login.microsoftonline.com/",
            tenant: app.tenant,
            clientId: app.clientId,
            endpoints: {
                "msgraph": "https://graph.microsoft.com"
            }
        }

        // Setup auth context
        var authContext = new AuthenticationContext(window.config);
        authContext.redirectUri = app.redirectUri;
        authContext.handleWindowCallback();

        var isCallback = authContext.isCallback(window.location.hash);
        var user = authContext.getCachedUser();

        // Check if the user is cached
        if (!user) {
            authContext.login();
        }
        else {
            // Get access token for graph
            authContext.acquireToken("https://graph.microsoft.com", function (error, token) {
                // Check for success
                if (error || !token) {
                    // Handle ADAL Error
                    response.status = "error";
                    response.accessToken = null;
                    Office.context.ui.messageParent(JSON.stringify(response));
                }
                else {
                    // Return the roken to the parent
                    response.status = "success";
                    response.accessToken = token;
                    Office.context.ui.messageParent(JSON.stringify(response));
                }
            });
        }
    }
})();
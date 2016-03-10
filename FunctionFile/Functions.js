/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
Office.initialize = function() {
}

function saveToOneDrivebak(event) {
    if (true) {
        //if (accessToken === "") {
        Office
            .context
            .ui
            .displayDialogAsync(
            "https://localhost:8443/signin.html",
            {
                height: 320,
                width: 240,
                requireHTTPS: true
            },
            function(result) {
                _dlg = result.value;
                _dlg.addEventHandler(
                    Microsoft
                        .Office
                        .WebExtension
                        .EventType
                        .DialogMessageReceived,
                    function(arg) {
                        var message = JSON.parse(arg.message);
                        console.log("Status: " + message.status);
                        console.log("Token: " + message.accessToken);
                        if (message.status == "success") {
                            _dlg.close();
                            accessToken = message.accessToken;
                            doStuff(event, "Brian");
                        }
                    });
            });

    } else {
        doStuff(event, "Robert");
    }
}

// This is our access token to OneDrive
var accessToken = "";

function saveToOneDrive(event) {
    if (!authenticated()) {
        authenticate(event);
    } else {
        // implement
    }
}

function authenticated() {
    return "" !== accessToken;
}

function authenticate(event) {
    var CLIENT_ID = "ffe6420a-cc97-4ed6-9928-351b9b0ff697",
        REDIRECT_URI = "https://localhost:8443/authorize.html",
        GRAPH_ID = "https://graph.microsoft.com",

        authUrl = "https://login.microsoftonline.com/common/oauth2"
            + "/authorize"
            + "?response_type=code"
            + "&client_id=" + CLIENT_ID
            + "&redirect_uri=" + REDIRECT_URI
            + "&resource=" + GRAPH_ID;
    
    Office
        .context
        .ui
        .displayDialogAsync(
        authUrl,
        {
            height: 320,
            width: 240,
            requireHTTPS: true
        },
        function(result) {
            _dlg = result.value;
            _dlg.addEventHandler(
                Microsoft
                    .Office
                    .WebExtension
                    .EventType
                    .DialogMessageReceived,
                function(arg) {
                    var message = JSON.parse(arg.message);
                    console.log("Status: " + message.status);
                    console.log("Token: " + message.accessToken);
                    if (message.status == "success") {
                        _dlg.close();
                        accessToken = message.accessToken;
                        doStuff(event, "Brian");
                    }
                });
        });
}

function doStuff(event, token) {
    Office.context.mailbox.item.notificationMessages.addAsync("subject", {
        type: "informationalMessage",
        icon: "blue-icon-16",
        message: "Token: " + token,
        persistent: false
    });
    // TODO implement
    event.completed();
}
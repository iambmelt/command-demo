/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
Office.initialize = function() {
}

// This is our access token to OneDrive
var accessToken = "";
var _event;
function saveToOneDrive(eventContext) {
    _event = eventContext;
    if (!authenticated()) {
        authenticate();
    } else {
        // TODO implement
        console.log(accessToken);
        doStuff(accessToken);
    }
}

function authenticated() {
    accessToken = window.localStorage.getItem('accessToken');
    return "" !== accessToken;
}


function authenticate() {
    var TENANT_ID = "ddfb6627-bdfd-4532-88cf-bfd6b4404248",
        AUTH_ENDPOINT = "https://login.microsoftonline.com/"
            + TENANT_ID
            + "/oauth2",
        CLIENT_ID = "ffe6420a-cc97-4ed6-9928-351b9b0ff697",
        REDIRECT_URI = "https://localhost:8443/authorize.html",
        GRAPH_ID = "https://graph.microsoft.com",

        authUrl =
            AUTH_ENDPOINT
            + "/authorize"
            + "?response_type=code"
            + "&client_id=" + CLIENT_ID
            + "&redirect_uri=" + REDIRECT_URI
            + "&resource=" + GRAPH_ID;

    Office
        .context
        .ui
        .displayDialogAsync(authUrl, {
            height: 40,
            width: 40,
            requireHTTPS: true
        }, onDialogOpen);
}

var dialog;
function onDialogOpen(result) {
    dialog = result.value;
    dialog.addEventHandler(
        Microsoft
            .Office
            .WebExtension
            .EventType
            .DialogMessageReceived,
        onMessageReceived);
}

function onMessageReceived(msg) {
    var debug = true;
    if (debug) {
        // not currently able to see the msg
        // return to the parent...
        dialog.close();
    } else {
        var message = JSON.parse(msg.message);
        console.log("Status: " + message.status);
        console.log("Token: " + message.accessToken);
        if (message.status == "success") {
            dialog.close();
            accessToken = message.accessToken;
            doStuff("Brian");
        } else {
            dialog.close();
        }
    }
}

function doStuff(token) {
    _event.completed();
    Office
        .context
        .mailbox
        .item
        .notificationMessages
        .addAsync("subject", {
            type: "informationalMessage",
            icon: "blue-icon-16",
            message: "Token: " + token,
            persistent: false
        });
    // TODO implement
    // _event.completed();
}
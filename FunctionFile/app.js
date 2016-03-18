/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var app,
    saveToOneDrive;

Office.initialize = function() {
    var app = new App();
    saveToOneDrive = function(eventContext) {
        if (!app.isAuthenticated()) {
            app.authenticate();
        }
        else {
            app.notify(app.token.access_token.slice(0, 5));
        }
    }

    function App() {
        this.token = null;
        this.dialog = null;

        this.isAuthenticated = function() {
            if (this.token == null) {
                this.token = window.localStorage.getItem('accessToken');
            }

            return this.token !== null;
        }

        this.authenticate = function() {
            var TENANT_ID = "ddfb6627-bdfd-4532-88cf-bfd6b4404248",
                AUTH_ENDPOINT = "https://login.microsoftonline.com/"
                    + TENANT_ID
                    + "/oauth2",
                CLIENT_ID = "ffe6420a-cc97-4ed6-9928-351b9b0ff697",
                REDIRECT_URI = "https://localhost:8443/authorize.html",
                GRAPH_ID = "https://graph.microsoft.com";

            // FIXME - generate an actual nonce
            var authUrl = AUTH_ENDPOINT
                + "/authorize"
                + "?response_type=id_token+token"
                + "&client_id=" + CLIENT_ID
                + "&scope=openid%20https%3A%2F%2Fgraph.microsoft.com%2Ffiles.readwrite"
                + "&nonce=23232432465433"
                + "&resource=" + GRAPH_ID;

            Office.context.ui.displayDialogAsync(authUrl, {
                height: 40,
                width: 40,
                requireHTTPS: true
            }, this.onDialogOpen);
        }

        this.onDialogOpen = function(result) {
            this.dialog = result.value;
            this.dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, onMessageReceived);
            this.dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, onDialogDismissed);
        }

        this.onMessageReceived = function(message) {
            var debug = true;
            this.notify(message);
        }

        this.onDialogDismissed = function(msg) {
            if (this.isAuthenticated()) {
                this.notify('Logged in :)');
            }
            else {
                this.notify('ERROR!');
            }
        }

        this.notify = function(message) {
            Office.context.mailbox.item.notificationMessages
                .addAsync("subject", {
                    type: "informationalMessage",
                    icon: "blue-icon-16",
                    message: message,
                    persistent: false
                });
        }
    }
}
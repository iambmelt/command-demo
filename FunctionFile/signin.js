Office.initialize = function (reason) {
        $(document).ready(initAuth);
}

function initAuth() {
    var CLIENT_ID = "ffe6420a-cc97-4ed6-9928-351b9b0ff697",
        REDIRECT_URI = "https://localhost:8443/authorize.html",
        GRAPH_ID = "https://graph.microsoft.com",
    
        authUrl = "https://login.microsoftonline.com/common/oauth2" +
            + "/authorize"
            + "?response_type=code"
            + "&client_id=" + CLIENT_ID
            + "&redirect_uri=" + REDIRECT_URI
            + "&resource=" + GRAPH_ID;
       
       //window.location.replace(authUrl);
       window.location.href = "authUrl";
}
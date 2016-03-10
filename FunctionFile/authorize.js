Office.initialize = function (reason) {
        $(document).ready(completeAuth);
}

function getToken(qs) {
    var kvs = qs.split('&');
        
    for (var i = 0; i < kvs.length; i++) {
        var kv = kvs[i].split("=");
        var key = kv[0];
        var value = kv[1];
        if (key.toLowerCase() === "code") {
            return value;
        }
    }
    
    return ""
}

function completeAuth() {
    Office.context.ui.messageParent(JSON.stringify({
        status: "success",
        accessToken: getToken(location.hash.substring(1))
    }));
}
    
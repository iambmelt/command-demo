Office.initialize = function(reason) {
    //$(document).ready(completeAuth);
    $(document).ready(function() {
        // Ok, which of these is supposedly correct?
        //Office.context.ui.messageParentAsync("Hello");
        //Office.context.ui.messageParent("Hello");
    });
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

    // FIXME handle this condition
    return "";
}

function completeAuth() {
    Office.context.ui.messageParent(JSON.stringify({
        status: "success",
        accessToken: getToken(location.hash.substring(1))
    }));
}

$(document).ready(function() {
    var token = getToken(window.document.URL);
    if (token !== null) {
        window.localStorage.setItem('accessToken', JSON.stringify(token));
        $('#success').show();
        $('#failed').hide();
    }
    else {
        $('#success').hide();
        $('#failed').show();
    }

    function getToken(hash) {
        var parts = hash.split('?');
        if (parts == null || parts.length <= 0) return null;

        var rightPart = parts.length == 2 ? parts[1] : parts[0];
        var token = getTokenFromString(rightPart);
        return token;
    }

    function getTokenFromString(hash) {
        var params = {},
            regex = /([^&=]+)=([^&]*)/g;

        var matches = regex.exec(hash);
        if (matches == null) return null;

        for (var i = matches.length - 1; i >= 0; i--) {
            params[decodeURIComponent(matches[i])] = decodeURIComponent(matches[i]);
        }

        return params;
    }
});
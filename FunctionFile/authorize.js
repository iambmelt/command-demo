$(document).ready(function() {
    var token = getToken(window.document.URL);
    $('#token').text('THe token is:' + JSON.stringify(token));
    window.localStorage.setItem('accessToken', JSON.stringify(token));
});

function getToken(hash) {
    var parts = hash.split('?');
    if (parts && parts.length > 0) {
        var rightPart = parts.length == 2 ? parts[1] : parts[0];
        var token = getTokenFromString(rightPart);
        return token;
    }

    return '';
}

function getTokenFromString(hash) {
    let params = {},
        regex = /([^&=]+)=([^&]*)/g,
        m;

    while ((m = regex.exec(hash)) !== null) {
        params[decodeURIComponent(m[1])] = decodeURIComponent(m[2]);
    }

    return params;
}

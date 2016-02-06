Office.initialize = function (reason) {
        $(document).ready(initAuth);
}

function initAuth(){

    var response = {"status":"none", "accessToken": ""};

    window.config = {
        instance: 'https://login.microsoftonline.com/',
        //tenant: 'dmpimemsftdev.onmicrosoft.com',
        clientId: 'b6f31d09-b541-41ec-b085-41eab842cfdb',
        postLogoutRedirectUri: "https://agave-node.azurewebsites.net/aadAuth.html",
        cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost.
        endpoints: {
            /* 'target endpoint to be called': 'target endpoint's resource ID' */
            'https://lists.office.com': 'https://lists.office.com'
        }

    }
    var authContext = new AuthenticationContext(window.config);
    authContext.redirectUri = "https://agave-node.azurewebsites.net/aadAuth.html";
    authContext.handleWindowCallback(window.location.hash);

    var isCallback = false;
    var user = authContext.getCachedUser();
    if (user) {

    } else {
        authContext.login();
    }

    if (authContext.getCachedUser()) {
        //$('#display-name').text("HELLO " + user.userName + "!");

        authContext.acquireToken(authContext.config.clientId, function (error, token) {
            // Handle ADAL Error
            if (error || !token) {
                console.log('ADAL Error Occurred: ' + error);
            }
            //console.log("id token: " + token);
        });
        authContext.acquireToken("https://lists.office.com", function (error, token) {
            // Handle ADAL Error
            if (error || !token) {
                console.log('ADAL Error Occurred: ' + error);
            } else {
                response.status = "success";
                response.accessToken = token;
                Office.context.ui.messageParent(JSON.stringify(response));
            }
        });
    }
}
    
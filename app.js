var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// uncomment after placing your favicon in /public
//app.use(favicon(path.join(__dirname, 'public', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

//QbDemo - Additional dependencies and middle layers
var session = require('express-session');
app.use(session({ resave: false, saveUninitialized: false, secret: 'smith' }));
var QuickBooks = require('node-quickbooks');
var request = require('request');
var qs = require('querystring');
var debug = require('debug')('myfirstexpressapp:server');

var consumerKey = 'qyprdgdAdo87q3etZEooPVomcAIGFr';
var consumerSecret = 'Ty24EXbvhqb8OEhBhc4B2VheuTKEQgUveI7fzgqa';
var accessToken = {};
var realmId = '';

//QbDemo - Set up HTTP routes for auth flow
app.get('/', function (req, res) {
    res.redirect('/home.html');
})

app.get('/requestToken', function (req, res) {
    var postBody = {
        url: QuickBooks.REQUEST_TOKEN_URL,
        oauth: {
            callback: 'https://agave-node.azurewebsites.net/callback/',
            consumer_key: consumerKey,
            consumer_secret: consumerSecret
        }
    }
    
    debug(postBody);
    
    request.post(postBody, function (e, r, data) {
        var requestToken = qs.parse(data)
        req.session.oauth_token_secret = requestToken.oauth_token_secret
        debug(requestToken)
        res.redirect(QuickBooks.APP_CENTER_URL + requestToken.oauth_token)
    })
})

app.get('/callback', function (req, res) {
    var postBody = {
        url: QuickBooks.ACCESS_TOKEN_URL,
        oauth: {
            consumer_key: consumerKey,
            consumer_secret: consumerSecret,
            token: req.query.oauth_token,
            token_secret: req.session.oauth_token_secret,
            verifier: req.query.oauth_verifier,
            realmId: req.query.realmId
        }
    }

    request.post(postBody, function (e, r, data) {
            
        //save the accessToken & realmId
        accessToken = qs.parse(data)
        realmId = postBody.oauth.realmId

        debug(accessToken)
        debug(postBody.oauth.realmId)

        res.redirect('/close.html?_host_Info=Excel|Win32|16.01|en-US|telemetry|isDialog') 
        //Note: The query string is only needed to workaround a known issue which causes context loss on server-side redirects.
    })
})

app.get('/getToken', function (req, res) {

    debug("Requested: " + accessToken)
    res.set("Expires", "0");
    res.send(accessToken);

})


app.get('/clearToken', function (req, res) {

    debug("Token cleared")
    accessToken = {};
    res.set("Expires", "0");
    res.status(200).end();

})

//Qbdemo - Setup routes for querying data
app.get('/getAccounts', function (req, res) {

    if (authenticated(res))
        getAccounts(function (accounts) { res.send(accounts) });

})

app.get('/getPurchases', function (req, res) {
    if (authenticated(res))
        getPurchases(function (purchases) { res.send(purchases) });

})

function authenticated(res) {
    var authenticated = false;

    if (accessToken.oauth_token &&
        accessToken.oauth_token_secret) {
        authenticated = true;
    }
    else {
        res.status(401).end();
        debug("Unauthenticated Request")
    }

    return authenticated;

}

function getAccounts(callback) {
  
    // save the access token somewhere on behalf of the logged in user
    var qbo = new QuickBooks(consumerKey,
        consumerSecret,
        accessToken.oauth_token,
        accessToken.oauth_token_secret,
        realmId,
        true, // use the Sandbox
        false); // turn debugging on

    // test out account access
    qbo.findAccounts(function (_, accounts) {
        debug(accounts.QueryResponse.Account.length + " accounts retrieved.")
        callback(accounts);
    })

}

function getPurchases(callback) {
  
    // save the access token somewhere on behalf of the logged in user
    var qbo = new QuickBooks(consumerKey,
        consumerSecret,
        accessToken.oauth_token,
        accessToken.oauth_token_secret,
        realmId,
        true, // use the Sandbox
        false); // turn debugging on

    // test out account access
    qbo.findPurchases(function (_, purchases) {
        debug(purchases.QueryResponse.Purchase.length + " purchases retrieved.")
        callback(purchases);
    })

}




module.exports = app;

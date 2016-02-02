//QboDemo - Main Code

var QuickBooks = require('node-quickbooks');
var request    = require('request');
var qs         = require('querystring');
var debug = require('debug')('myfirstexpressapp:server');

var consumerKey    = 'qyprdgdAdo87q3etZEooPVomcAIGFr',
    consumerSecret = 'Ty24EXbvhqb8OEhBhc4B2VheuTKEQgUveI7fzgqa',
    accessToken    = {contents:"empty"},
    realmId        = ''

exports.init = function (mainApp){
    var app = mainApp;

    app.get('/',function(req,res){
    res.redirect('/start');
    })

    app.get('/start', function(req, res) {
    res.render('home.ejs', {port:6000, appCenter: QuickBooks.APP_CENTER_BASE})
    })

    app.get('/requestToken', function(req, res) {
    var postBody = {
        url: QuickBooks.REQUEST_TOKEN_URL,
        oauth: {
        //callback:        'https://localhost:' + 6000 + '/callback/',
        callback:        'https://agave-node.azurewebsites.net/callback/',
        consumer_key:    consumerKey,
        consumer_secret: consumerSecret
        }
    }
    request.post(postBody, function (e, r, data) {
        var requestToken = qs.parse(data)
        req.session.oauth_token_secret = requestToken.oauth_token_secret
        debug(requestToken)
        res.redirect(QuickBooks.APP_CENTER_URL + requestToken.oauth_token)
    })
    })

    app.get('/callback', function(req, res) {
    var postBody = {
        url: QuickBooks.ACCESS_TOKEN_URL,
        oauth: {
        consumer_key:    consumerKey,
        consumer_secret: consumerSecret,
        token:           req.query.oauth_token,
        token_secret:    req.session.oauth_token_secret,
        verifier:        req.query.oauth_verifier,
        realmId:         req.query.realmId
        }
    }   
    
    request.post(postBody, function (e, r, data) {
            
            //save the accessToken & realmId
            accessToken = qs.parse(data)
            realmId = postBody.oauth.realmId
            
            debug(accessToken)
            debug(postBody.oauth.realmId)
            
            res.redirect('/close.html')

        })
    })


    app.get('/getAccounts', function(req, res) {

        if(authenticated(res))
            getAccounts(function(accounts) {res.send(accounts)});
    
    })

    app.get('/getPurchases', function(req, res) {
        if(authenticated(res))
            getPurchases(function(purchases) {res.send(purchases)});
    
    })

    app.get('/token', function(req, res) {

        debug("Requested: " + accessToken)
        res.set("Expires", "0");
        res.send(accessToken);
    
    })

    app.get('/renderAccounts', function(req, res) {
        getAccounts(function(accounts) {
            res.send("<html><body>" + JSON.stringify(accounts) + "</body></html>");
        })

    })

}

function authenticated(res){
    var authenticated = false;
    
    if(accessToken.oauth_token &&
        accessToken.oauth_token_secret) {
           authenticated = true;
        }
        else {
            res.status(401).end();
            debug("Unauthenticated Request")
        }
        
    return authenticated;
    
}

var getAccounts = function(callback){
  
      // save the access token somewhere on behalf of the logged in user
    var qbo = new QuickBooks(consumerKey,
                         consumerSecret,
                         accessToken.oauth_token,
                         accessToken.oauth_token_secret,
                         realmId,
                         true, // use the Sandbox
                         false); // turn debugging on

    // test out account access
    qbo.findAccounts(function(_, accounts) {
      debug(accounts.QueryResponse.Account.length + " accounts retrieved.")
        callback(accounts);
    })
    
}

var getPurchases = function(callback){
  
      // save the access token somewhere on behalf of the logged in user
    var qbo = new QuickBooks(consumerKey,
                         consumerSecret,
                         accessToken.oauth_token,
                         accessToken.oauth_token_secret,
                         realmId,
                         true, // use the Sandbox
                         false); // turn debugging on

    // test out account access
    qbo.findPurchases(function(_, purchases) {
      debug(purchases.QueryResponse.Purchase.length + " purchases retrieved.")
        callback(purchases);
    })
    
}
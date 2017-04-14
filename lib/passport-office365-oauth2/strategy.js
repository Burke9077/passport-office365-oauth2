/**
 * Module dependencies.
 */
var util = require('util'),
    https = require('https'),
    OAuth2Strategy = require('passport-oauth').OAuth2Strategy,
    InternalOAuthError = require('passport-oauth').InternalOAuthError;


/**
 * `Strategy` constructor.
 *
 * The Office 365 authentication strategy closely resembles Azure authentication
 * with a few minor tweeks.
 *
 * Options:
 *   - `clientID`           specifies the client id of the application that is registered in Azure Active Directory
 *   - `clientSecret`       secret used to establish ownership of the client Id
 *   - `callbackURL`        URL to which Azure AD will redirect the user after obtaining authorization
 *   - `scope`              List of resources requested (User.Read)
 *   - `authorizationURL`   Authorization url for the basis of all calls (eg: https://login.microsoftonline.com/common/oauth2/v2.0/authorize)
 *   - `tokenURL`           Token URL
 *
 * Examples:
 *
 *     var Office365Oauth2Strategy = require('passport-office365-oauth2').Strategy;
 *     passport.use("office365", new Office365Oauth2Strategy ({
 *         clientID: 'yourClientId',
 *         clientSecret: 'yourClientSecret',
 *         callbackURL: 'https://www.example.net/auth/azureadoauth2/callback',
 *         tokenURL: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
 *         authorizationURL: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
 *         scope: 'User.Read'
 *       },
 *       function (accessToken, refresh_token, params, profile, done) {
 *         // Deal with the user data in your own way
 *       }
 *     ));
 *
 * @param {Object} options
 * @param {Function} verify
 * @api public
 */
function Strategy(options, verify) {
    options = options || {};

    var base_url = options.authorizationURL;

    OAuth2Strategy.call(this, options, verify);

    this.name = 'office365_oauth2';
}

/**
 * Inherit from `OAuth2Strategy`.
 */
util.inherits(Strategy, OAuth2Strategy);

/**
 * Authenticate request by delegating to Azure AD using OAuth.
 *
 * @param {Object} req
 * @api protected
 */
Strategy.prototype.authenticate = function(req, options) {
    if (!options.resource && this.resource) { // include default resource as authorization parameter
        options.resource = this.resource;
    }

    // Call the base class for standard OAuth authentication.
    OAuth2Strategy.prototype.authenticate.call(this, req, options);
};

/**
 * Retrieve user profile from Azure AD.
 *
 * This function constructs a normalized profile, with the following properties:
 *
 *   - `provider`         always set to `office365_oauth2`
 *   - `id`
 *   - `username`
 *   - `displayName`
 *
 * @param {String} accessToken
 * @param {Function} done
 * @api protected
 */
Strategy.prototype.userProfile = function(accessToken, callback) {
    // Get the user profile
    var options = {
        host: 'graph.microsoft.com',
        path: '/v1.0/me',
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json',
            Authorization: 'Bearer ' + accessToken
        }
    };

    https.get(options, function(response) {
        var body = '';
        response.on('data', function(d) {
            body += d;
        });
        response.on('end', function() {
            var error;
            if (response.statusCode === 200) {
                return callback(null, JSON.parse(body));
            } else {
                error = new Error();
                error.code = response.statusCode;
                error.message = response.statusMessage;
                // The error body sometimes includes an empty space
                // before the first character, remove it or it causes an error.
                body = body.trim();
                error.innerError = JSON.parse(body).error;
                return callback(error);
            }
        });
    }).on('error', function(e) {
        return callback(e);
    });
};

/**
 * Return extra Azure AD-specific parameters to be included in the authorization
 * request.
 *
 * @param {Object} options
 * @return {Object}
 * @api protected
 */
Strategy.prototype.authorizationParams = function(options) {
    return options;
};

/**
 * Expose `Strategy`.
 */
module.exports = Strategy;

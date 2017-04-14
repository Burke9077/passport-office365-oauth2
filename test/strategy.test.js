var expect = require('chai').expect;
var AzureAdOAuth2Strategy = require('..');

describe('Strategy', function() {

    var strategy = new AzureAdOAuth2Strategy({
            clientID: 'yourClientId',
            clientSecret: 'yourClientSecret',
            callbackURL: 'https://www.example.net/auth/azureadoauth2/callback',
            tokenURL: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            authorizationURL: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
            scope: 'User.Read'
        },
        function() {});

    it('should be named office365_oauth2', function() {
        expect(strategy.name).to.equal('office365_oauth2');
    });
});

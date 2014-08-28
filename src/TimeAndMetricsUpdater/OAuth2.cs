using System;
using System.Diagnostics;
using System.IO;
using Google.GData.Client;
using Newtonsoft.Json;
using Microsoft.VisualBasic;

namespace TimeAndMetricsUpdater
{
    public static class OAuth2
    {
        public static GOAuth2RequestFactory GetOAuthFactory() {
            // OAuth2Parameters holds all the parameters related to OAuth 2.0.
            var parameters = new OAuth2Parameters();

            if (!File.Exists("oauth.parameters"))
                throw new Exception("Missing oauth.parameters");

            parameters = JsonConvert.DeserializeObject<OAuth2Parameters>(File.ReadAllText("oauth.parameters"));
            if (parameters.AccessToken == null)
                parameters = Initialize(parameters);

            ////////////////////////////////////////////////////////////////////////////
            // STEP 5: Make an OAuth authorized request to Google
            ////////////////////////////////////////////////////////////////////////////

            // Initialize the variables needed to make the request
            var requestFactory = new GOAuth2RequestFactory(null, "TimeAndMetrics", parameters);
            return requestFactory;
        }

        private static OAuth2Parameters Initialize(OAuth2Parameters parameters) {
            ////////////////////////////////////////////////////////////////////////////
            // STEP 1: Configure how to perform OAuth 2.0
            ////////////////////////////////////////////////////////////////////////////

            // Space separated list of scopes for which to request access.
            var SCOPE = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";

            // This is the Redirect URI for installed applications.
            // If you are building a web application, you have to set your
            // Redirect URI at https://code.google.com/apis/console.
            var REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";

            ////////////////////////////////////////////////////////////////////////////
            // STEP 2: Set up the OAuth 2.0 object
            ////////////////////////////////////////////////////////////////////////////

            // Set your Redirect URI, which can be registered at
            // https://code.google.com/apis/console.
            parameters.RedirectUri = REDIRECT_URI;

            ////////////////////////////////////////////////////////////////////////////
            // STEP 3: Get the Authorization URL
            ////////////////////////////////////////////////////////////////////////////

            // Set the scope for this particular service.
            parameters.Scope = SCOPE;

            parameters.AccessType = "offline";

            // Get the authorization url.  The user of your application must visit
            // this url in order to authorize with Google.  If you are building a
            // browser-based application, you can redirect the user to the authorization
            // url.
            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
            Process.Start(authorizationUrl);
            parameters.AccessCode = Interaction.InputBox("Please authorize your OAuth request token in the browser " +
                "window that just opened. Once that is complete, paste in your access code " +
                    "to continue...", "Time and Metrics Sync: Google Authorization");

            ////////////////////////////////////////////////////////////////////////////
            // STEP 4: Get the Access Token
            ////////////////////////////////////////////////////////////////////////////

            // Once the user authorizes with Google, the request token can be exchanged
            // for a long-lived access token.  If you are building a browser-based
            // application, you should parse the incoming request token from the url and
            // set it in OAuthParameters before calling GetAccessToken().
            OAuthUtil.GetAccessToken(parameters);
            File.WriteAllText("oauth.parameters", JsonConvert.SerializeObject(parameters));
            return parameters;
        }
    }
}

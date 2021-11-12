// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <IndexRouterSnippet>
var express = require("express");
var router = express.Router();

/* GET home page. */
router.get("/", async function (req, res, next) {
  let token = "";
  if (req.app.locals.msalClient) {
    const account = await req.app.locals.msalClient
      .getTokenCache()
      .getAccountByHomeId(req.session.userId);

    if (account) {
      // Attempt to get the token silently
      // This method uses the token cache and
      // refreshes expired tokens as needed
      const response = await req.app.locals.msalClient.acquireTokenSilent({
        scopes: process.env.OAUTH_SCOPES.split(","),
        redirectUri: process.env.OAUTH_REDIRECT_URI,
        account: account,
      });
      // First param to callback is the error,
      // Set to null in success case
      token = response.accessToken;
    }
  }
  let params = {
    token,
    active: { home: true },
  };

  res.render("index", params);
});

module.exports = router;
// </IndexRouterSnippet>

/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
* This sample shows how to:
*    - Get the current user's metadata
*    - Get the current user's profile photo
*    - Attach the photo as a file attachment to an email message
*    - Upload the photo to the user's root drive
*    - Get a sharing link for the file and add it to the message
*    - Send the email
*/
const express = require('express');
const router = express.Router();
const graphHelper = require('../utils/graphHelper.js');
const emailer = require('../utils/emailer.js');
const passport = require('passport');
const request = require('superagent');
const http = require('http');
// ////const fs = require('fs');
// ////const path = require('path');

const clientState = 'patata';

// Get the home page.
router.get('/', (req, res) => {
  // check if user is authenticated
  if (!req.isAuthenticated()) {
    res.render('login');
  } else {
    renderSendMail(req, res);
  }
});

// Authentication request.
router.get('/login',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      res.redirect('/');
    });

// Authentication callback.
// After we have an access token, get user data and load the sendMail page.
router.get('/token',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      graphHelper.getUserData(req.user.accessToken, (err, user) => {
        if (!err) {
          requestSubscription(req.user);

          req.user.profile.displayName = user.body.displayName;
          req.user.profile.emails = [{ address: user.body.mail || user.body.userPrincipalName }];
          renderSendMail(req, res);
        } else {
          renderError(err, res);
        }
      });
    });

// Load the sendMail page.
function renderSendMail(req, res) {
  res.render('sendMail', {
    display_name: req.user.profile.displayName,
    email_address: req.user.profile.emails[0].address
  });
}

const subscriptions = {};

function requestSubscription(user) {
  const subcriptionConfig = {
    changeType: 'created, updated',
    notificationUrl: 'https://e9c6d0f2.ngrok.io/listen',
    resource: '/me/events',
    expirationDateTime: new Date(Date.now() + 86400000).toISOString(),
    clientState
  };

  request.post('https://graph.microsoft.com/beta/subscriptions')
    .set({
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + user.accessToken
    })
    .send(subcriptionConfig)
    .end((err, res) => {
      if (err) {
        console.error('Error trying to subscribe to resources: ' + err.message);
        return;
      }

      console.log('Subscribed to notifications: ' + JSON.stringify(res.body));
      subscriptions[res.body.id] = {
        userToken: user.accessToken
      };
    });
}

/* Default listen route */
router.post('/listen', (req, res, next) => {
  let clientStatesValid;

  // If there's a validationToken parameter in the query string,
  // then this is the request that Office 365 sends to check
  // that this is a valid endpoint.
  // Just send the validationToken back.
  if (req.query && req.query.validationToken) {
    res.send(req.query.validationToken);
    // Send a status of 'Ok'
    res.status(200).end(http.STATUS_CODES[200]);
  } else {
    clientStatesValid = false;

    // First, validate all the clientState values in array
    for (let i = 0; i < req.body.value.length; i++) {
      if (req.body.value[i].clientState !== clientState) {
        // If just one clientState is invalid, we discard the whole batch
        clientStatesValid = false;
        break;
      } else {
        clientStatesValid = true;
      }
    }

    // Send a status of 'Accepted'
    // If the clientState field doesn't have the expected value,
    // this request might NOT come from Microsoft Graph.
    // However, you should still return the same status that you'd
    // return to Microsoft Graph to not alert possible impostors
    // that you have discovered them.
    res.status(202).end(http.STATUS_CODES[202]);

    // If all the clientStates are valid, then process the notification
    if (clientStatesValid) {
      for (let i = 0; i < req.body.value.length; i++) {
        const notification = req.body.value[i];
        processNotification(notification, req);
      }
    } 
  }
});

// Get subscription data from the database
// Retrieve the actual event data from Office 365.
// Send the message data to the socket.
function processNotification(notification) {
  const subIds = Object.keys(subscriptions);

  let subscription;
  for(let i = 0; i < subIds.length; i++) {
    const id = subIds[i];
    subscription = id === notification.subscriptionId && subscriptions[id];
  }

  if (!subscription) {
    // Receiving subscription from a different session, ignore.
    return;
  }

  console.log(`Received ${notification.changeType} notification for resource ${notification.resource}`);

  request.get('https://graph.microsoft.com/beta/' + notification.resource)
    .set({ Authorization: 'Bearer ' + subscription.userToken })
    .end((err, res) => {
      if (err) {
        console.err('Error requesting event information: ' + err.message);
        return;
      }
      
      console.log('Event information received: ' + JSON.stringify(res.body));
    });
}

// Do prep before building the email message.
// The message contains a file attachment and embeds a sharing link to the file in the message body.
function prepForEmailMessage(req, callback) {
  const accessToken = req.user.accessToken;
  const displayName = req.user.profile.displayName;
  const destinationEmailAddress = req.body.default_email;
  // Get the current user's profile photo.
  graphHelper.getProfilePhoto(accessToken, (errPhoto, profilePhoto) => {
    // //// TODO: MSA flow with local file (using fs and path?)
    if (!errPhoto) {
        // Upload profile photo as file to OneDrive.
        graphHelper.uploadFile(accessToken, profilePhoto, (errFile, file) => {
          // Get sharingLink for file.
          graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
            const mailBody = emailer.generateMailBody(
              displayName,
              destinationEmailAddress,
              link.webUrl,
              profilePhoto
            );
            callback(null, mailBody);
          });
        });
      }
      else {
        var fs = require('fs');
        var readableStream = fs.createReadStream('public/img/test.jpg');
        var picFile;
        var chunk;
        readableStream.on('readable', function() {
          while ((chunk=readableStream.read()) != null) {
            picFile = chunk;
          }
      });
      
      readableStream.on('end', function() {

        graphHelper.uploadFile(accessToken, picFile, (errFile, file) => {
          // Get sharingLink for file.
          graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
            const mailBody = emailer.generateMailBody(
              displayName,
              destinationEmailAddress,
              link.webUrl,
              picFile
            );
            callback(null, mailBody);
          });
        });
      });
      }
  });
}

// Send an email.
router.post('/sendMail', (req, res) => {
  const response = res;
  const templateData = {
    display_name: req.user.profile.displayName,
    email_address: req.user.profile.emails[0].address,
    actual_recipient: req.body.default_email
  };
  prepForEmailMessage(req, (errMailBody, mailBody) => {
    if (errMailBody) renderError(errMailBody);
    graphHelper.postSendMail(req.user.accessToken, JSON.stringify(mailBody), (errSendMail) => {
      if (!errSendMail) {
        response.render('sendMail', templateData);
      } else {
        if (hasAccessTokenExpired(errSendMail)) {
          errSendMail.message += ' Expired token. Please sign out and sign in again.';
        }
        renderError(errSendMail, response);
      }
    });
  });
});

router.get('/disconnect', (req, res) => {
  req.session.destroy(() => {
    req.logOut();
    res.clearCookie('graphNodeCookie');
    res.status(200);
    res.redirect('/');
  });
});

// helpers
function hasAccessTokenExpired(e) {
  let expired;
  if (!e.innerError) {
    expired = false;
  } else {
    expired = e.forbidden &&
      e.message === 'InvalidAuthenticationToken' &&
      e.response.error.message === 'Access token has expired.';
  }
  return expired;
}
/**
 * 
 * @param {*} e 
 * @param {*} res 
 */
function renderError(e, res) {
  e.innerError = (e.response) ? e.response.text : '';
  res.render('error', {
    error: e
  });
}

module.exports = router;

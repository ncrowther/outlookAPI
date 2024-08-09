'use strict';

var utils = require('../utils/writer.js');
var OutlookAPI = require('../service/OutlookAPIService');

module.exports.bookmeeting = function bookmeeting(req, res, next, body) {

  const authheader = req.headers.authorization;
  console.log(req.headers);

  const auth = new Buffer.from(authheader.split(' ')[1],
    'base64').toString().split(':');
  const user = auth[0];
  const pass = auth[1];

  OutlookAPI.bookmeeting(body, pass)
    .then(function (response) {
      utils.writeJson(res, response);
    })
    .catch(function (response) {
      utils.writeJson(res, response);
    });
};

module.exports.planmeeting = function planmeeting(req, res, next, body) {

  const authheader = req.headers.authorization;
  console.log(req.headers);

  const auth = new Buffer.from(authheader.split(' ')[1],
    'base64').toString().split(':');
  const user = auth[0];
  const pass = auth[1];

  OutlookAPI.planmeeting(body, pass)
    .then(function (response) {
      utils.writeJson(res, response);
    })
    .catch(function (response) {
      utils.writeJson(res, response);
    });
};

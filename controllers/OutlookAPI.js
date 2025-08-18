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

module.exports.sendemail = function sendemail(req, res, next, body) {

  const authheader = req.headers.authorization;
  console.log(req.headers);

  const auth = new Buffer.from(authheader.split(' ')[1],
    'base64').toString().split(':');
  const user = auth[0];
  const pass = auth[1];

  OutlookAPI.sendemail(body, pass)
    .then(function (response) {
      utils.writeJson(res, response);
    })
    .catch(function (response) {
      utils.writeJson(res, response);
    });
};

module.exports.getCalendar = function getCalendar(req, res, next, body) {

  const authheader = req.headers.authorization;
  console.log(req.headers);

  const startDate = req.query.startDate + 'T00:00Z'
  const endDate = req.query.endDate + 'T23:00Z'

  console.log("*********" + startDate);

  const auth = new Buffer.from(authheader.split(' ')[1], 'base64').toString().split(':');
  const user = auth[0];
  const pass = auth[1];
  
  let token = pass

  console.log("****TOKEN*****" + token);

  OutlookAPI.getCalendar(body, token, startDate, endDate)
    .then(function (response) {
      utils.writeJson(res, response);
    })
    .catch(function (response) {
      utils.writeJson(res, response);
    });
};
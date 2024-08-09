'use strict';

var request = require('request');
var util = require('util');


/**
 * Outlook API
 * PLan a meeting though Microsoft Graph API
 *
 * body PlanRequest OutlookAPI prompt
 * apiKey String Api Key
 * returns PlanResponse
 **/
exports.planmeeting = function (body, apiKey) {
  return new Promise(function (resolve, reject) {

    console.log('Plan meeting ApiKey: ', apiKey);

    var options = {
      'method': 'POST',
      'url': 'https://graph.microsoft.com/v1.0/me/findMeetingTimes',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + apiKey
      },
      body: JSON.stringify({
        "attendees": [
          {
            "emailAddress": {
              "address": body.recipient,
            },
            "type": "Required"
          }
        ],
        "timeConstraint": {
          "timeslots": [
            {
              "start": {
                "dateTime": body.windowStartDateTime,
                "timeZone": "UTC"
              },
              "end": {
                "dateTime": body.windowEndDateTime,
                "timeZone": "UTC"
              }
            }
          ]
        },
        "locationConstraint": {
          "isRequired": "false",
          "suggestLocation": "true",
          "locations": [
            {
              "displayName": "Conf Room 32/1368",
              "locationEmailAddress": "conf32room1368@imgeek.onmicrosoft.com"
            }
          ]
        },
        "meetingDuration": body.duration
      })

    };

    request(options, function (error, response) {

      if (error) throw new Error(error);

      console.log(response.body);

      var body = JSON.parse(response.body)
      var slots = [];
      var i = 0;

      if (body.meetingTimeSuggestions) {
        body.meetingTimeSuggestions.forEach(function (suggestion) {

          console.log('***Suggestion ' + suggestion)

          let startSlot = suggestion.meetingTimeSlot.start.dateTime.replace(":00.0000000", ":00Z");
          let endSlot = suggestion.meetingTimeSlot.end.dateTime.replace(":00.0000000", ":00Z");

          var timeslot = {
            "startDateTime": startSlot,
            "endDateTime": endSlot,
          }

          var slot = timeslot

          slots[i] = slot;
          i++;

        })
      }

      var response = {
        "slots": slots
      }

      resolve(response);

    });

  })
}


/**
 * Outlook API
 * Book a meeting though Microsoft Graph API
 *
 * body Request OutlookAPI prompt
 * apiKey String Api Key
 * returns Response
 **/
exports.bookmeeting = function (body, apiKey) {

  return new Promise(function (resolve, reject) {

    console.log('ApiKey: ', apiKey);

    var subject = body.subject;
    console.log('Subject: ', subject);

    var options = {
      'method': 'POST',
      'url': 'https://graph.microsoft.com/v1.0/me/events',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + apiKey
      },
      body: JSON.stringify({
        "subject": body.subject,
        "body": {
          "contentType": "HTML",
          "content": body.messageBody
        },
        "start": {
          "dateTime": body.startDateTime, //"2024-08-09T18:00:25.357Z",
          "timeZone": "UTC"
        },
        "end": {
          "dateTime": body.endDateTime, //"2024-08-09T18:00:25.357Z",
          "timeZone": "UTC"
        },
        "location": {
          "displayName": body.location
        },
        "attendees": [
          {
            "emailAddress": {
              "address": body.recipient
            },
            "type": "required"
          }
        ],
        "allowNewTimeProposals": true
      })

    };

    request(options, function (error, response) {

      if (error) throw new Error(error);

      console.log(response.body);

      var returnString = "Failed to Book"
      if (response.statusCode === 201) {
        returnString = "Meeting Booked"
      }
  
      response = {
        "response": returnString
      }

      resolve(response);

    });

  })
}


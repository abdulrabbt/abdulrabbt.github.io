Office.initialize = function() {};

function isIE() {
  return typeof msCrypto === "object";
}

function isSupported() {
  return typeof crypto === "object";
}

function getRandom() {
  var letters_pool = "23456789ABCDEFGHJKLMNPQRST";
  return letters_pool[Math.floor(Math.random() * letters_pool.length)];
}

function arrayBufferToString(buffer) {
  var binary = "";
  var bytes = new Uint8Array(buffer);
  var len = bytes.byteLength;
  for (var i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return binary;
}

function genPass(length) {
  var result = "";
  for (var i = 0; i < length; ++i) {
    result += getRandom();
  }
  return result;
}

function roundToHour(date) {
  p = 60 * 60 * 1000; // milliseconds in an hour
  return new Date(Math.round(date.getTime() / p) * p);
}

function constructURL(domain, room, password, encryption) {
  return "https://" + domain + "/#" + room + "/" + password + "/" + encryption;
}

function callback(domain, room, password, encryption, event) {
  var currentEmail = Office.context.mailbox.userProfile.emailAddress;
  var url = constructURL(domain, room, password, encryption);

  // set the starting date to be beginning of next hour
  // ends a half hour after
  var startTime = roundToHour(new Date());
  var endTime = new Date(startTime);
  endTime.setMinutes(endTime.getMinutes() + 30);

  Office.context.mailbox.displayNewAppointmentForm({
    requiredAttendees: [currentEmail],
    location: "Online",
    subject: "",
    resources: [],
    start: startTime,
    end: endTime,
    body:
      "\n-- Do not delete or change any of the following text. -- " +
      "\nTo join the meeting, check the info and link below: " +
      "\nMeeting ID (access code): " +
      room +
      "\nMeeting password: " +
      password +
      "\nMeeting encryption: " +
      encryption +
      "\n\nJoin the meeting: " +
      url +
      "\nThis meeting is powered by Meet.sa (a SITE product).",
  });

  event.completed();
}

function genUrl(event) {
  console.log(event);
  var room = genPass(8);
  var password = genPass(14);
  var domain = event.source.id;
  var cryptObt;

  // If msCrypto is present, then use it
  if (isIE()) {
    // NOTE: msCrypto is only supported in IE11 (+ Outlook for win)
    cryptObt = msCrypto.subtle.generateKey(
      { name: "AES-GCM", length: 128 },
      true,
      ["encrypt", "decrypt"]
    );

    cryptObt.oncomplete = function(e) {
      var ob = msCrypto.subtle.exportKey("jwk", e.target.result);
      ob.oncomplete = function(e2) {
        var result = e2.target.result;
        var k = JSON.parse(arrayBufferToString(result)).k;
        return callback(domain, room, password, k, event);
      };
    };
  } else if (isSupported()) {
    cryptObt = crypto.subtle
      .generateKey({ name: "AES-GCM", length: 128 }, true, [
        "encrypt",
        "decrypt",
      ])
      .then(function(e) {
        return crypto.subtle.exportKey("jwk", e);
      })
      .then(function(e2) {
        var k = e2.k;
        return callback(domain, room, password, k, event);
      });
  } else {
    // TODO: crypto is not supported
    return event.completed();
  }
}

Office.onReady(function() {});

Office.initialize = function() {};

// function isIE() {
//   return document.documentMode;
// }

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

function genUrl(event) {
  var room = genPass(8);
  var password = genPass(14);
  var url = "https://" + event.source.id + "/#" + room + "/" + password + "/";
  var cryptObt;
  // If msCrypto is present, then use it
  if (typeof msCrypto === 'object') {
    // NOTE: msCrypto is only supported in IE11 (+ Outlook for win)
    cryptObt = msCrypto.subtle.generateKey(
      { name: "AES-GCM", length: 128 },
      true,
      ["encrypt", "decrypt"]
    );

    cryptObt.oncomplete = function(e) {
      var ob = msCrypto.subtle.exportKey("jwk", e.target.result);
      ob.oncomplete = function(e2) {
        console.log("from oncomplete", e2);
        var result = e2.target.result;
        var k = JSON.parse(arrayBufferToString(result)).k;
        var currentEmail = Office.context.mailbox.userProfile.emailAddress;
        Office.context.mailbox.displayNewAppointmentForm({
          requiredAttendees: [currentEmail],
          location: "Online",
          subject: "Ray Meeting",
          resources: [],
          body: "\n\n\n\nJoin Ray Meeting \n " + url + k,
        });

        event.completed();
      };
    };
  } else {
    cryptObt = crypto.subtle.generateKey(
      { name: "AES-GCM", length: 128 },
      true,
      ["encrypt", "decrypt"]
    );

    cryptObt.oncomplete = function(e) {
      var ob = crypto.subtle.exportKey("jwk", e.target.result);
      ob.oncomplete = function(e2) {
        console.log("from oncomplete", e2);
        var result = e2.target.result;
        var k = JSON.parse(arrayBufferToString(result)).k;
        var currentEmail = Office.context.mailbox.userProfile.emailAddress;
        Office.context.mailbox.displayNewAppointmentForm({
          requiredAttendees: [currentEmail],
          location: "Online",
          subject: "Ray Meeting",
          resources: [],
          body: "\n\n\n\nJoin Ray Meeting \n " + url + k,
        });

        event.completed();
      };
    };
  }
}

Office.onReady(function() {});

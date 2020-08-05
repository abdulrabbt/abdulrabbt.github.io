Office.initialize = function () {};

function isIE() {
  return document.documentMode;
}

function getRandom() {
  var letters_pool = "23456789ABCDEFGHJKLMNPQRST";
  return letters_pool[Math.floor(Math.random() * letters_pool.length)];
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
  msCrypto.subtle
    .generateKey({ name: "AES-GCM", length: 128 }, true, ["encrypt", "decrypt"])
    .then(function (key) {
      let key_exported = msCrypto.subtle.exportKey("jwk", key);
      key_exported = JSON.parse(arrayBufferToString(key_exported.result)).k;

      Office.context.mailbox.displayNewAppointmentForm({
        requiredAttendees: [],
        location: "Online",
        subject: "Ray Meeting",
        resources: [],
        body: "\n\n\n\nJoin Ray Meeting \n " + url + key_exported,
      });
      event.completed();
    });
}

Office.onReady(function () {});

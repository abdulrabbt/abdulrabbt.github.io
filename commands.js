Office.initialize = function () {};

function isIE() {
  return document.documentMode;
}

function getRandom() {
  var letters_pool = "23456789ABCDEFGHJKLMNPQRST";
  return letters_pool[Math.floor(Math.random() * letters_pool.length)];
}

function arrayBufferToString( buffer ) {
  var binary = '';
  var bytes = new Uint8Array( buffer );
  var len = bytes.byteLength;
  for (var i = 0; i < len; i++) {
      binary += String.fromCharCode( bytes[ i ] );
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
  var crypt = msCrypto.subtle.generateKey(
    { name: "AES-GCM", length: 128 },
    true,
    ["encrypt", "decrypt"]
  );
  crypt.oncomplete = function (e) {
    let key_exported = msCrypto.subtle.exportKey("jwk", e.target.result);
    key_exported = JSON.parse(arrayBufferToString(key_exported.result)).k;

    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees: [],
      location: "Online",
      subject: "Ray Meeting",
      resources: [],
      body: "\n\n\n\nJoin Ray Meeting \n " + url + key_exported,
    });
    event.completed();
  };
}

Office.onReady(function () {});

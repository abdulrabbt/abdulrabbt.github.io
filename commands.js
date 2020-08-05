Office.initialize = function() {};

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

// async function generateKey() {
//   var key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 128 }, true, ["encrypt", "decrypt"]);
//   var key_exported = await crypto.subtle.exportKey("jwk", key);
//   return key_exported.k;
// }

// async function generateKeyIE() {
//   let key = await msCrypto.subtle.generateKey({ name: "AES-GCM", length: 128 }, true, ["encrypt", "decrypt"]);
//   key = key.result;

//   let key_exported = msCrypto.subtle.exportKey("jwk", key);
//   key_exported = JSON.parse(arrayBufferToString(key_exported.result));

//   return key_exported.k;
// }

function genUrl(event) {
  var room = genPass(8);
  var password = genPass(14);
  // var encryptionKey = "tlyYLfRxK1YPXaChLQAcPQ";
  var encryptionKey;
  if (isIE()) {
    encryptionKey = "testing ";
  } else {
    encryptionKey = "testing IE";
  }
  var url = "https://" + event.source.id + "/#" + room + "/" + password + "/" + encryptionKey;

  Office.context.mailbox.displayNewAppointmentForm({
    requiredAttendees: [],
    location: "Online",
    subject: "Ray Meeting",
    resources: [],
    body: "\n\n\n\nJoin Ray Meeting \n " + url
  });
  event.completed();
}

Office.onReady(function() {});

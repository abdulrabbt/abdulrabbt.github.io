Office.initialize = function() {};

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

async function generateKey() {
    console.log("key");
    var key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 128 }, true, ["encrypt", "decrypt"]);
    var key_exported = await crypto.subtle.exportKey("jwk", key);
    return key_exported.k;
  };

console.log("test")
async function genUrl(event) {
  var room = genPass(8);
  var password = genPass(14);
  var encryptionKey = await generateKey(); //"tlyYLfRxK1YPXaChLQAcPQ";
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

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

async function generateKey() {
    console.log("key");
    var key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 128 }, true, ["encrypt", "decrypt"]);
    var key_exported = await crypto.subtle.exportKey("jwk", key);
    return key_exported.k;
  };

function generateKeyIE() {
    console.log("key IE");
    return msCrypto.subtle.generateKey({ name: "AES-GCM", length: 128 }, true, ["encrypt", "decrypt"]).then( function (key) {
      return msCrypto.subtle.exportKey("jwk", key.result).then(function(key_exported){
        return JSON.parse(arrayBufferToString(key_exported.result)).k
      })
          })
  }

console.log("test")
async function genUrl(event) {
  var room = genPass(8);
  var password = genPass(14);
  // var encryptionKey = await generateKey(); //"tlyYLfRxK1YPXaChLQAcPQ";
  var encryptionKey
  if(isIE()){
    encryptionKey = await generateKeyIE()
  } else {
    encryptionKey = await generateKey();
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

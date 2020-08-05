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
    console.log(e)
//     console.log("111",msCrypto.subtle.exportKey("jwk", e.target.result))
    var ob = msCrypto.subtle.exportKey("jwk", e.target.result);
    console.log("ob", ob.result)
    ob.oncomplete = function(e2) {
      console.log("from oncomplete", e2); 
      var result = e2.target.result;
      console.log("result", result);
      console.log("buffer from oncomplete", arrayBufferToString(result));
    }
    
//     console.log("result",msCrypto.subtle.exportKey("jwk", e.target.result).result)
//     console.log("buffer",arrayBufferToString(ob.result))
//     let key_exported = msCrypto.subtle.exportKey("jwk", e.target.result);
//     key_exported = JSON.parse(arrayBufferToString(key_exported.result)).k;

    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees: [],
      location: "Online",
      subject: "Ray Meeting",
      resources: [],
      body: "\n\n\n\nJoin Ray Meeting \n " + url,
    });
    event.completed();
  };
}

Office.onReady(function () {});

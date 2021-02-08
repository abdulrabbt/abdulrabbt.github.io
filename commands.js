Office.initialize = function () {};

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
  var meetUrl = "https://meet.sa/";
  var isRestricted = domain == "site-internal.meet.sa";
  // set the starting date to be beginning of next hour
  // ends a half hour after
  var startTime = roundToHour(new Date());
  var endTime = new Date(startTime);
  endTime.setMinutes(endTime.getMinutes() + 30);

  // if desktop use <br /> otherwise use /n
  var platform = Office.context.platform;
  var nlC = "\n"; // new line char
  if (platform === "OfficeOnline") {
    nlC = "<br />";
    url = '<a href="' + url + '">' + url + "</a>";
    meetUrl = '<a href="' + meetUrl + '">' + meetUrl + "</a>";
  }
  var body =
    nlC +
    nlC +
    "-- Do not delete or change any of the following text. -- " +
    nlC +
    nlC +
    "To join the meeting, follow the link and info below:" +
    nlC +
    url +
    nlC +
    nlC +
    (isRestricted
      ? "Note: VPN or Company's internal network is required" + nlC + nlC
      : "") +
    "Meeting ID (access code): " +
    room +
    nlC +
    nlC +
    "Meeting password: " +
    password +
    nlC +
    nlC +
    "Meeting encryption: " +
    encryption +
    nlC +
    nlC +
    "This meeting is powered by Meet.sa. For any inquiries, please contact support@meet.sa" +
    nlC +
    meetUrl;

  if (platform === "Mac") {
    nlC = "\r\n";
    body = "To join the meeting, please ";
    if (isRestricted) {
      body += "use the company network or a VPN and ";
    }
    body += "follow this link: " + url;
  }

  var showApptForm = function () {
    Office.context.mailbox.displayNewAppointmentForm({
      requiredAttendees: [currentEmail],
      location: "Online",
      subject: "",
      resources: [],
      start: startTime,
      end: endTime,
      // NOTE: web only supports HTML (\n doesn't work), desktop doesn't supports HTML (\n works, while <br /> doesn't)
      body: body,
    });
  };

  if (isRestricted) {
    var dialog;
    Office.context.ui.displayDialogAsync(
      "https://abdulrabbt.github.io/dialog.html",
      { height: 25, width: 30, displayInIframe: true },
      function (asyncResult) {
        if (asyncResult.status != "failed") {
          // In addition to general system errors, there are 3 specific errors for
          // displayDialogAsync that you can handle individually.
          // https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/blob/master/SimpleDialogSampleWeb/DialogHelper.js
          dialog = asyncResult.value;
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            function (arg) {
              dialog.close();
              if (arg.message == "yes") {
                showApptForm();
              }
              event.completed();
            }
          );
          dialog.addEventHandler(
            Office.EventType.DialogEventReceived,
            function () {
              dialog.close();
              event.completed();
            }
          );
        }
      }
    );
  } else {
    showApptForm();
    event.completed();
  }
}

function genUrl(event) {
  var room = genPass(5);
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

    cryptObt.oncomplete = function (e) {
      var ob = msCrypto.subtle.exportKey("jwk", e.target.result);
      ob.oncomplete = function (e2) {
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
      .then(function (e) {
        return crypto.subtle.exportKey("jwk", e);
      })
      .then(function (e2) {
        var k = e2.k;
        return callback(domain, room, password, k, event);
      });
  } else {
    // TODO: crypto is not supported
    return event.completed();
  }
}

Office.onReady(function () {});

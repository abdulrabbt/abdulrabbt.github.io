Office.initialize = function() {};

function genUrl(event) {
  // console.log("TESTING");
//   var room = "room"; // genPass(8);
//   var password = "password"; // genPass(14);
//   var encryptionKey = "tlyYLfRxK1YPXaChLQAcPQ"; // await generateKey();
//   var url = `https://${event.source.id}/#${room}/${password}/${encryptionKey}`;

  Office.context.mailbox.displayNewAppointmentForm({
    requiredAttendees: [],
    location: "Online",
    subject: "Ray Meeting",
    resources: [],
    body: "\n\n\n\nJoin Ray Meeting \n"
  });
//   Office.context.mailbox.item.notificationMessages.addAsync("subject", {
//     type: "informationalMessage",
//     icon: "Icon.16x16",
//     message: "Subject: 11",
//     persistent: false
//   });
  event.completed();
}

Office.onReady(function() {

});

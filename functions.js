function reportEmail(event) {
  Office.context.mailbox.item.forwardAsync({
    toRecipients: ["spam@wervik.be"],
    subject: "Verdachte e-mail doorgestuurd voor controle",
    htmlBody: "Deze e-mail werd gemeld door een gebruiker ter controle.",
    attachments: []
  }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      Office.context.mailbox.item.itemId && Office.context.mailbox.item.removeAsync(function (result) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("success", {
          type: "informationalMessage",
          message: "De e-mail is doorgestuurd naar IT en wordt verwijderd.",
          icon: "icon16",
          persistent: false
        });
        event.completed();
      });
    } else {
      Office.context.mailbox.item.notificationMessages.replaceAsync("error", {
        type: "errorMessage",
        message: "Er trad een fout op bij het doorsturen."
      });
      event.completed();
    }
  });
}
Office.initialize = function () {};

// Charge la bibliothèque Office.js.
Office.onReady();

// Fonction d’assistance pour ajouter un message d’état à la barre de notifications.
function statusUpdate(icon, text, event) {
  const details = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: icon,
    message: text,
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", details, { asyncContext: event }, asyncResult => {
    const event = asyncResult.asyncContext;
    event.completed();
  });
}
//hello world function
function showHelloWorldMessage(event) {
    Office.context.mailbox.item.notificationMessages.addAsync("helloWorldMessage", {
        type: "informationalMessage",
        message: "Hello, World!",
        icon: "icon16",
        persistent: false
    });
    event.completed();
}


// Affiche une barre de notification.
function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!", event);
}

// Mappe le nom de fonction spécifié dans le manifeste à son équivalent JavaScript.
Office.actions.associate("defaultStatus", defaultStatus);
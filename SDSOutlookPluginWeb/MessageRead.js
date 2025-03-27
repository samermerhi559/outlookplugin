(function () {
  "use strict";

  let messageBanner;

  // Charge la bibliothèque Office.js.
  Office.onReady(function (reason) {
    $(() => {
      const element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
    });
  });

  // Crée une liste de noms de pièces jointes.
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      let returnString = "";
      
      for (let i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Met en forme les détails du contact par nom, nom de famille et adresse e-mail donnés
    function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // Crée une liste de détails de contact mis en forme. $loc_script_mail_commands_read_js_comment3$)$$   function buildEmailAddressesString(addresses) {     if (addresses && addresses.length > 0) {       let returnString = "";        for (let i = 0 ; i < addresses.length ; i++) {         if (i > 0) {           returnString = returnString + « <br/> »;         }         returnString = returnString + buildEmailAddressString(addresses[i]);       }        return returnString;     }      return « None »;   }    // $$LOC(Charge les propriétés à partir de l’objet de base Item, puis charge les
  // propriétés spécifiques au message.
  function loadProps() {
    const item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }

  // Fonction d'assistance pour afficher les notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();
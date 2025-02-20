function onMessageSendHandler(event) {
    Office.context.mailbox.item.body.getAsync(
      "text",
      { asyncContext: event },
      isMessageSensitive
    );
  }
  
/* 2 Approaches here. We are following the Approach 1 for simplicity.
Approach 1: Fetch email body using Office.js, 
            send it to server, 
            run AI on the payload, 
            return true or false for isMsgSensitive. 
            If true, 
                send the itemid to addin server. 
                server will encrypt the email body and send it. 
            Else, send the email as is.

Approach 2. Send only itemid to addin server, 
            Server will get payload using itemid (client credentials auth
            run AI on the payload, 
            return true or false for isMsgSensitive. 
            If true, 
                send the itemid to addin server. 
                server will encrypt the email body and send it. 
            Else, send the email as is.
*/

  function isMessageSensitive(asyncResult){
    // Here, run the AI model to check if the email body contains sensitive data.
    isMsgSensitive = true;
    const event = asyncResult.asyncContext;
    console.log(asyncResult);
    if(isMsgSensitive){
        console.log("Encrypt and Send");
        event.completed({
            allowEvent: false, // set to true if message need not be encrypted, else false
            cancelLabel: "Encrypt and Send", // Label for the button that allows the user to cancel the send operation
            errorMessage: "Looks like your message has sensitive data. Do you want to encrypt?",
            // TIP: In addition to the formatted message, it's recommended to also set a
            // plain text message in the errorMessage property for compatibility on
            // older versions of Outlook clients.
            errorMessageMarkdown: "Looks like your message has sensitive data. Do you want to encrypt?",
            commandId: "EncryptAndSend" // the function that is mapped to this commandId will be executed when the user clicks on "Encrypt and Send". 
            // Check the function name in manifest.xml.
          });
    } else {
        console.log("Sending the email");
        event.completed({
            allowEvent: true
          });
    }
  }
  
  
  // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
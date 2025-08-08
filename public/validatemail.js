console.log("Addin loaded...");

Office.onReady(() => {
    console.log("Debug: Office.onReady called");
    console.log("Debug: Registering validateBody");
    Office.actions.associate("validateBody", validateBody);
    console.log("Debug: Handler registered successfully");
}).catch((error) => {
    console.error("Debug: Office.js failed:", error);
});

var mailboxItem;

function operatingSytem() { 
    var contextInfo = Office.context.diagnostics;
    console.log('Office application: ' + contextInfo.host);
    console.log('Platform: ' + contextInfo.platform);
    console.log('Office version: ' + contextInfo.version);
}

function validateBody(event){
    operatingSytem();
    mailboxItem = Office.context.mailbox.item;
    mailboxItem.body.getAsync("html", {asyncContext: event}, checkBodyOnlyOnSend);
}

function checkBodyOnlyOnSend(asyncResult){
    var bodyContent = asyncResult.value;
    console.log("body output:", bodyContent);
    if(bodyContent.includes("block")){
        mailboxItem.notificationMessages.addAsync('NoSend', {type: 'errorMessage', message: 'Mail is Blocked'});
        asyncResult.asyncContext.completed({allowEvent: false});
    }
    else{
        asyncResult.asyncContext.completed({allowEvent: true});
    }
}
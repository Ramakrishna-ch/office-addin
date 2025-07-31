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

function validateBody(event){
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

Office.actions.associate("validateBody", validateBody);
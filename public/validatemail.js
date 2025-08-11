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

async function dspAgentserverCheck(params) {
    fetch("https://localhost:5001/validate", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: "test@test.com"
    })
        .then((response) => response.json())
        .then((result) => console.log(result))
        .catch((error) => console.error(error));
}

function validateBody(event){
    operatingSytem();
    mailboxItem = Office.context.mailbox.item;
    mailboxItem.body.getAsync("html", {asyncContext: event}, checkBodyOnlyOnSend);
}

function checkBodyOnlyOnSend(asyncResult){
    var bodyContent = asyncResult.value;
    console.log("body output:", bodyContent);
    dspAgentserverCheck();
    if(bodyContent.includes("block")){
        mailboxItem.notificationMessages.addAsync('NoSend', {type: 'errorMessage', message: 'Mail is Blocked'});
        asyncResult.asyncContext.completed({allowEvent: false});
    }
    else{
        asyncResult.asyncContext.completed({allowEvent: true});
    }
}
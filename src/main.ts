Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('dumpButton').onclick = showEmailPreview;
    }
});


async function showEmailPreview() {
    document.getElementById('emailPreview').innerText = await getMessageContent()
}


async function getMessageContent(): Promise<string> {
    const item = Office.context.mailbox.item;
    if(item.itemType === Office.MailboxEnums.ItemType.Message) {
        try {
            const emailBody = await getBody(item, Office.CoercionType.Text);
            return emailBody;
        } catch (error) {
            console.log("Fehler beim Abrufen des E-Mail-Bodys:", error);
            return "";
        }
    } else {
        console.log("Das aktuelle Element ist keine Nachricht.");
        return "";
    }
}


function getBody(item: Office.MessageRead | Office.MessageCompose, coercionType: Office.CoercionType): Promise<string> {
    return new Promise((resolve, reject) => {
        item.body.getAsync(coercionType, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('dumpButton').onclick = showEmailPreview;
    }
});


async function showEmailPreview() {
    document.getElementById('emailPreview').innerText = await getMessageText()
}


async function getMessageText(): Promise<string> {
    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
        try {
            const emailBody = await new Promise<string>((resolve, reject) => {
                item.body.getAsync(Office.CoercionType.Html, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value as string);
                    } else {
                        reject(result.error);
                    }
                });
            });
            return emailBody;
        } catch (error) {
            console.log("Fehler beim Abrufen des E-Mail-Bodys:", error);
            return ""; // or handle the error appropriately
        }
    } else {
        console.log("Das aktuelle Element ist keine Nachricht.");
        return ""; // or handle the case where it's not a message
    }
}



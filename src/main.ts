

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('dumpButton').onclick = dumpMessageContent;
    }
});

function dumpMessageContent() {
    let emailText = Office.context.mailbox.item.body.getAsync(Office.CoercionType.HTML);
    document.getElementById('emailPreview').innerText = emailText;
}
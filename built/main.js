//import 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'
Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('dumpButton').onclick = dumpMessageContent;
    }
});
function dumpMessageContent() {
    var emailText = Office.context.mailbox.item.body.getAsync(Office.CoercionType.HTML);
    document.getElementById('emailPreview').innerText = emailText;
}

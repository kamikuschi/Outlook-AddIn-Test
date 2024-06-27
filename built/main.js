var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('dumpButton').onclick = showEmailPreview;
    }
});
function showEmailPreview() {
    return __awaiter(this, void 0, void 0, function* () {
        document.getElementById('emailPreview').innerText = yield getMessageContent();
    });
}
function getMessageContent() {
    return __awaiter(this, void 0, void 0, function* () {
        const item = Office.context.mailbox.item;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            try {
                const emailBody = yield getBody(item, Office.CoercionType.Text);
                return emailBody;
            }
            catch (error) {
                console.log("Fehler beim Abrufen des E-Mail-Bodys:", error);
                return "";
            }
        }
        else {
            console.log("Das aktuelle Element ist keine Nachricht.");
            return "";
        }
    });
}
function getBody(item, coercionType) {
    return new Promise((resolve, reject) => {
        item.body.getAsync(coercionType, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            }
            else {
                reject(result.error);
            }
        });
    });
}
//# sourceMappingURL=main.js.map
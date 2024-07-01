var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import { ErrorHandler } from "./errorhandler.js";
export class OutlookInterface extends ErrorHandler {
    constructor() {
        super();
    }
    getMessageText() {
        return __awaiter(this, void 0, void 0, function* () {
            const item = Office.context.mailbox.item;
            if (!item) {
                this.setError("Es konnte kein Element im Postfach gefunden werden");
                return "";
            }
            if (item.itemType === Office.MailboxEnums.ItemType.Message) {
                const emailBody = yield new Promise((resolve, reject) => {
                    item.body.getAsync(Office.CoercionType.Html, (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolve(result.value);
                        }
                        else {
                            this.setError("Fehler beim Abrufen des E-Mail-Bodys", result.error);
                            reject(result.error);
                        }
                    });
                });
                return emailBody;
            }
            else {
                this.setError("Das aktuelle Element ist keine Nachricht");
                return "";
            }
        });
    }
    getMessageFile() {
        return __awaiter(this, void 0, void 0, function* () {
            this.clearErrors();
            const item = Office.context.mailbox.item;
            if (!item) {
                this.setError("Es konnte kein Element im Postfach gefunden werden");
                return "";
            }
            if (item.itemType === Office.MailboxEnums.ItemType.Message) {
                const emailFile = yield new Promise((resolve, reject) => {
                    //const options: Office.AsyncContextOptions = { asyncContext: { currentItem: item } };
                    item.getAsFileAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolve(result.value);
                        }
                        else {
                            this.setError("Fehler beim Abrufen des E-Mail-Bodys", result.error);
                            reject(result.error);
                        }
                    });
                });
                return atob(emailFile);
            }
            else {
                this.setError("Das aktuelle Element ist keine Nachricht");
                return "";
            }
        });
    }
    getAttachments() {
        return __awaiter(this, void 0, void 0, function* () {
            this.clearErrors();
            const item = Office.context.mailbox.item;
            if (!item) {
                this.setError("Es konnte kein Element im Postfach gefunden werden");
                return null;
            }
            if (item.itemType === Office.MailboxEnums.ItemType.Message) {
                const attachments = yield new Promise((resolve, reject) => {
                    //const options: Office.AsyncContextOptions = { asyncContext: { currentItem: item } };
                    item.getAttachmentsAsync((result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolve(result.value);
                        }
                        else {
                            this.setError("Fehler beim Abrufen der Anhänge", result.error);
                            reject(result.error);
                        }
                    });
                });
                return attachments;
            }
            else {
                this.setError("Das aktuelle Element ist keine Nachricht");
                return null;
            }
        });
    }
}
//# sourceMappingURL=outlookinterface.js.map
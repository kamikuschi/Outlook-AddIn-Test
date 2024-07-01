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
        this.item = Office.context.mailbox.item;
    }
    isMessage() {
        if (!this.item) {
            this.setError("Es konnte kein Element im Postfach gefunden werden");
            return false;
        }
        if (this.item.itemType !== Office.MailboxEnums.ItemType.Message) {
            this.setError("Das aktuelle Element ist keine Nachricht");
            return false;
        }
        return true;
    }
    getMessageFile() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.isMessage()) {
                const emailFile = yield new Promise((resolve, reject) => {
                    this.item.getAsFileAsync((result) => {
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
                return "";
            }
        });
    }
    getMessageProperties() {
        if (this.isMessage()) {
            let subject = this.item.subject;
            let sender = this.item.sender.emailAddress;
            let date = this.item.dateTimeCreated;
            let recipients = [];
            this.item.to.forEach((recipient) => {
                recipients.push(recipient.emailAddress);
            });
            let ccs = [];
            this.item.cc.forEach((cc) => {
                ccs.push(cc.emailAddress);
            });
            return new MessageProperties(subject, sender, recipients, ccs, date);
        }
        else {
            return new MessageProperties;
        }
    }
}
class MessageProperties {
    constructor(subject, sender, recipients, ccs, date) {
        this.subject = subject || "";
        this.sender = sender || "";
        this.recipients = recipients || [];
        this.ccs = ccs || [];
        this.date = date || new Date;
    }
}
/*async getMessageText(): Promise<string> {
    const item = Office.context.mailbox.item;
    if(!item) {
        this.setError("Es konnte kein Element im Postfach gefunden werden");
        return "";
    }
    if(item.itemType === Office.MailboxEnums.ItemType.Message) {
        const emailBody = await new Promise<string>((resolve, reject) => {
            item.body.getAsync(Office.CoercionType.Html, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value as string);
                } else {
                    this.setError("Fehler beim Abrufen des E-Mail-Bodys", result.error as Error);
                    reject(result.error);
                }
            });
        });
        return emailBody;
    } else {
        this.setError("Das aktuelle Element ist keine Nachricht");
        return "";
    }
}*/ 
//# sourceMappingURL=outlookinterface.js.map
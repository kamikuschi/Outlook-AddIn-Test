import { ErrorHandler } from "./errorhandler.js";

export class OutlookInterface extends ErrorHandler {
    constructor() {
        super();
    }

    async getMessageText(): Promise<string> {
        const item = Office.context.mailbox.item;
        if(!item) {
            this.setError("Es konnte kein Element im Postfach gefunden werden")
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
    }

    async getMessageFile(): Promise<string> { // as EML file
        this.clearErrors()
        const item = Office.context.mailbox.item;
        if(!item) {
            this.setError("Es konnte kein Element im Postfach gefunden werden")
            return "";
        }
        if(item.itemType === Office.MailboxEnums.ItemType.Message) {
            const emailFile = await new Promise<string>((resolve, reject) => {
                //const options: Office.AsyncContextOptions = { asyncContext: { currentItem: item } };
                item.getAsFileAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value as string);
                    } else {
                        this.setError("Fehler beim Abrufen des E-Mail-Bodys", result.error as Error);
                        reject(result.error);
                    }
                });
            });
            return atob(emailFile);
        } else {
            this.setError("Das aktuelle Element ist keine Nachricht");
            return "";
        }

    }

    async getAttachments(): Promise<Office.AttachmentDetailsCompose[] | null> {
        this.clearErrors()
        const item = Office.context.mailbox.item;
        if(!item) {
            this.setError("Es konnte kein Element im Postfach gefunden werden")
            return null;
        }
        if(item.itemType === Office.MailboxEnums.ItemType.Message) {
            const attachments = await new Promise<Office.AttachmentDetailsCompose[]>((resolve, reject) => {
                //const options: Office.AsyncContextOptions = { asyncContext: { currentItem: item } };
                item.getAttachmentsAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        this.setError("Fehler beim Abrufen der Anhänge", result.error as Error);
                        reject(result.error);
                    }
                });
            });
            return attachments;
        } else {
            this.setError("Das aktuelle Element ist keine Nachricht");
            return null;
        }

    }
}
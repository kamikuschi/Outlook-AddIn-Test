import { ErrorHandler } from "./errorhandler.js";

export class OutlookInterface extends ErrorHandler {
    public readonly item: (Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead) | undefined;
    
    constructor() {
        super();
        this.item = Office.context.mailbox.item;
    }

    isMessage(): boolean {
        if(!this.item) {
            this.setError("Es konnte kein Element im Postfach gefunden werden");
            return false;
        }
        if(this.item.itemType !== Office.MailboxEnums.ItemType.Message) {
            this.setError("Das aktuelle Element ist keine Nachricht");
            return false;
        }
        return true;
    }



    async getMessageFile(): Promise<string> { // as EML file
        if(this.isMessage()) {
            const emailFile = await new Promise<string>((resolve, reject) => {
                this.item!.getAsFileAsync((result) => {
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
            return "";
        }
    }

    getMessageProperties(): MessageProperties {
        if(this.isMessage()) {
            let subject: string = this.item!.subject;
            let sender: string = this.item!.sender.emailAddress;
            let date: Date = this.item!.dateTimeCreated;
            let recipients: Array<string> = [];
            this.item!.to.forEach((recipient) => {
                recipients.push(recipient.emailAddress);
            })
            let ccs: Array<string> = [];
            this.item!.cc.forEach((cc) => {
                ccs.push(cc.emailAddress);
            })

            return new MessageProperties(subject, sender, recipients, ccs, date);
        } else {
            return new MessageProperties;
        }
    }
}

class MessageProperties {
    public subject: string;
    public sender: string;
    public recipients: Array<string>;
    public ccs: Array<string>;
    public date: Date;
    constructor(subject?: string, sender?: string, recipients?: Array<string>, ccs?: Array<string>, date?: Date) {
        this.subject = subject || ""
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
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
import { OutlookInterface } from "./outlookinterface.js";
Office.onReady(info => { main(info); });
function main(info) {
    let outlookInterface = new OutlookInterface;
    let errorHandler = new ErrorHandler;
    let sendToHistoryButton;
    if (info.host === Office.HostType.Outlook) {
        sendToHistoryButton = document.getElementById('sendToHistoryButton');
        if (!sendToHistoryButton) {
            errorHandler.setError("Fehler im HTML-Dokument", new Error("Element sendToHistoryButton isn't defined."));
            return;
        }
        sendToHistoryButton.addEventListener('click', () => __awaiter(this, void 0, void 0, function* () {
            console.log(outlookInterface.getMessageProperties());
            //download("message.eml", await outlookInterface.getMessageFile())
        }));
    }
}
function download(filename, data) {
    let element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(data));
    element.setAttribute('download', filename);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
}
/*async function setIframePreview(iframeId: string, contentPromise: Promise<string>) {
    const iframe = document.getElementById(iframeId) as HTMLIFrameElement;

    if (iframe && iframe.contentWindow && iframe.contentDocument) {
        try {
            // Warte, bis die Promise aufgelöst wird und erhalte den Inhalt
            const content = await contentPromise;

            // Setze den Inhalt des Iframes
            iframe.contentDocument.open();
            iframe.contentDocument.write(content);
            iframe.contentDocument.close();
        } catch (error) {
            console.error('Fehler beim Laden des Inhalts:', error);
        }
    } else {
        console.error(`Iframe mit ID ${iframeId} nicht gefunden oder Zugriff nicht möglich.`);
    }
}*/
//# sourceMappingURL=main.js.map
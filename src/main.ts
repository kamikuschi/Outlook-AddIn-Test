import { ErrorHandler } from "./errorhandler.js";
import { OutlookInterface } from "./outlookinterface.js"

Office.onReady(info => {main(info)});

function main(info: { host: Office.HostType; platform: Office.PlatformType; }) {
    let outlookInterface: OutlookInterface = new OutlookInterface;
    let errorHandler: ErrorHandler = new ErrorHandler;
    let sendToHistoryButton: HTMLElement | null;

    if (info.host === Office.HostType.Outlook) {
        sendToHistoryButton = document.getElementById('sendToHistoryButton');

        if(!sendToHistoryButton) {
            errorHandler.setError("Fehler im HTML-Dokument", new Error("Element sendToHistoryButton isn't defined."));
            return;
        }

        sendToHistoryButton.addEventListener('click', async () => {
            console.log(outlookInterface.getMessageProperties());
            //download("message.eml", await outlookInterface.getMessageFile())
        });
    }
}


function download(filename: string, data: string) {
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

import { ErrorHandler } from "./errorhandler.js";
import { OutlookInterface } from "./outlookinterface.js"

Office.onReady(info => {main(info)});

function main(info: { host: Office.HostType; platform: Office.PlatformType; }) {
    let outlookInterface: OutlookInterface = new OutlookInterface;
    let errorHandler: ErrorHandler = new ErrorHandler;
    let dumpButton: HTMLElement | null;
    
    if (info.host === Office.HostType.Outlook) {
        dumpButton = document.getElementById('dumpButton');
        if(!dumpButton) {
            errorHandler.setError("Das Skript kann nicht auf den Knopf zugreifen");
            return;
        }
        dumpButton.addEventListener('click', async () => {
            //setIframePreview('emailPreview', outlookInterface.getMessageFile());
            //console.log(outlookInterface.getMessageFile());
            
            download("message.eml", await outlookInterface.getMessageFile())
        });

        dumpButton.addEventListener('click', async () => {
            //setIframePreview('emailPreview', outlookInterface.getMessageFile());
            //console.log(outlookInterface.getMessageFile());
            
            download("message.eml", await outlookInterface.getMessageFile())
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


async function setIframePreview(iframeId: string, contentPromise: Promise<string>) {
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
}


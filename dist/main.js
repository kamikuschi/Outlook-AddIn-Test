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
    let dumpButton;
    if (info.host === Office.HostType.Outlook) {
        dumpButton = document.getElementById('dumpButton');
        if (!dumpButton) {
            errorHandler.setError("Das Skript kann nicht auf den Knopf zugreifen");
            return;
        }
        dumpButton.addEventListener('click', () => __awaiter(this, void 0, void 0, function* () {
            //setIframePreview('emailPreview', outlookInterface.getMessageFile());
            //console.log(outlookInterface.getMessageFile());
            download("message.eml", yield outlookInterface.getMessageFile());
        }));
        dumpButton.addEventListener('click', () => __awaiter(this, void 0, void 0, function* () {
            //setIframePreview('emailPreview', outlookInterface.getMessageFile());
            //console.log(outlookInterface.getMessageFile());
            download("message.eml", yield outlookInterface.getMessageFile());
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
function setIframePreview(iframeId, contentPromise) {
    return __awaiter(this, void 0, void 0, function* () {
        const iframe = document.getElementById(iframeId);
        if (iframe && iframe.contentWindow && iframe.contentDocument) {
            try {
                // Warte, bis die Promise aufgelöst wird und erhalte den Inhalt
                const content = yield contentPromise;
                // Setze den Inhalt des Iframes
                iframe.contentDocument.open();
                iframe.contentDocument.write(content);
                iframe.contentDocument.close();
            }
            catch (error) {
                console.error('Fehler beim Laden des Inhalts:', error);
            }
        }
        else {
            console.error(`Iframe mit ID ${iframeId} nicht gefunden oder Zugriff nicht möglich.`);
        }
    });
}
//# sourceMappingURL=main.js.map
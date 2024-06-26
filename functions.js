// Initialize the Office Add-in.
Office.onReady(() => {
    // If needed, Office.js is ready to be called
  });

// The command function.
async function sayHello(event) {
    try {
        await Office.context.mailbox.item.body.setAsync(
            "Hello world!",
            {
                coercionType: "html", // Write text as HTML
            }
        );
    } catch (error) {
        // Hier könntest du die Fehlerbehandlung verbessern
        console.error(error);
    }

    // Signalisiere, dass die Verarbeitung abgeschlossen ist
    event.completed();
}

// You must register the function with the following line.
Office.actions.associate("sayHello", sayHello);
export class ErrorHandler {
    constructor() {
        this._messages = [];
        this._exceptions = [];
    }
    setError(errorMessage, exception) {
        this._messages.push(errorMessage);
        if (exception) {
            this._exceptions.push(exception);
        }
        else {
            this._exceptions.push(new Error("None"));
        }
        console.log(errorMessage + ": " + this._exceptions[this._exceptions.length - 1].message);
    }
    /*public logErrors() {
        for(let i = 0; i < this._messages.length; i++) {
            console.log(this._messages[i] + ': ' + this._exceptions[i].message);
        }
    }*/
    clearErrors() {
        this._messages = [];
        this._exceptions = [];
    }
}
//# sourceMappingURL=errorhandler.js.map
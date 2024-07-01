export class ErrorHandler {
    private _messages: Array<string>;
    private _exceptions: Array<Error>;

    constructor() {
        this._messages = [];
        this._exceptions = [];
    }

    public setError(errorMessage: string, exception?: Error) {
        this._messages.push(errorMessage);
        if(exception) {
            this._exceptions.push(exception);
        } else {
            this._exceptions.push(new Error("None"))
        }
        for(let i = 0; i < this._messages.length; i++) {
            console.log(this._messages[i] + ': ' + this._exceptions[i].message);
        }
    }

    /*public logErrors() {
        for(let i = 0; i < this._messages.length; i++) {
            console.log(this._messages[i] + ': ' + this._exceptions[i].message);
        }
    }*/

    public clearErrors() {
        this._messages = [];
        this._exceptions = [];
    }
}
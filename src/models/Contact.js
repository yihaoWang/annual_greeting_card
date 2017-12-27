export default class Contact {

    constructor(data) {
        this._email = data.email || null;
        this._identity = [data.identity || null];
        this._name = [data.name || null];
        this._nickname = [data.nickname || null];
        this._unit = [data.unit || null];
        this._department = [data.department || null];
        this._paperCard = data.department || false;
        this._annualReport = data.annualReport || false;
        this._annualReceipt = data.annualReceipt || false;
    }

    getEmail() {
        return this._email;
    }
}

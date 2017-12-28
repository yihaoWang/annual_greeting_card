import clone from 'lodash/clone';
import trim from 'lodash/trim';
import compact from 'lodash/compact';

export default class Contact {

    constructor() {
        this.email = null;
        this.names = new Set();
        this.addresses = new Set();
        this.identities = new Set();
        this.nicknames = new Set();
        this.units = new Set();
        this.departments = new Set();
        this.paperCard = false;
        this.annulReport = false;
        this.annulReceipt = false;
    }

    static fromObject(object) {
        let result = new Contact();

        result.setEmail(object.email);
        result.addNames([object.name]);
        result.addAddresses([object.address]);
        result.addIdentities([object.identity]);
        result.addNicknames([object.nickname]);
        result.addUnits([object.unit]);
        result.addDepartments([object.department]);
        result.addPaperCard(object.paperCard);
        result.addAnnulReport(object.annualReport);
        result.addAnnulReceipt(object.annualReceipt);

        return result;
    }

    setEmail(email) {
        this.email = (trim(email) || null);
    }

    addNames(names) {
        this.names = new Set([
            ...this.names,
            ...compact(names),
        ]);
    }

    addAddresses(addresses) {
        this.addresses = new Set([
            ...this.addresses,
            ...compact(addresses),
        ]);
    }

    addIdentities(identities) {
        this.identities = new Set([
            ...this.identities,
            ...compact(identities),
        ]);
    }

    addNicknames(nicknames) {
        this.nicknames = new Set([
            ...this.nicknames,
            ...compact(nicknames),
        ]);
    }

    addUnits(units) {
        this.units = new Set([
            ...this.units,
            ...compact(units),
        ]);
    }

    addDepartments(departments) {
        this.departments = new Set([
            ...this.departments,
            ...compact(departments),
        ]);
    }

    addPaperCard(paperCard) {
        this.paperCard = (this.paperCard || paperCard || false);
    }

    addAnnulReport(annulReport) {
        this.annulReport = (this.annulReport || annulReport || false);
    }

    addAnnulReceipt(annulReceipt) {
        this.annulReceipt = (this.annulReceipt || annulReceipt || false);
    }

    get nameAddressList() {
        let result = [];

        for (let name of this.names) {
            for (let address of this.addresses) {
                result.push(`${name}/${address}`);
            }
        }

        return result;
    }

    merge(other) {
        let result = clone(this);

        if (!result.email) {
            result.setEmail(other.email);
        }

        result.addNames([...other.names]);
        result.addAddresses([...other.addresses]);
        result.addIdentities([...other.identities]);
        result.addNicknames([...other.nicknames]);
        result.addUnits([...other.units]);
        result.addDepartments([...other.departments]);
        result.addPaperCard(other.paperCard);
        result.addAnnulReport(other.annulReport);
        result.addAnnulReceipt(other.annulReceipt);

        return result;
    }

    toJSON() {
        return JSON.stringify({
            email: this.email,
            names: [...this.names],
            addresses: [...this.addresses],
            identities: [...this.identities],
            nicknames: [...this.nicknames],
            units: [...this.units],
            departments: [...this.departments],
            paperCard: this.paperCard,
            annulReport: this.annulReport,
            annulReceipt: this.annulReceipt,
        });
    }
}

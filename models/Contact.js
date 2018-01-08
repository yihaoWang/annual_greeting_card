import clone from 'lodash/clone';
import trim from 'lodash/trim';
import compact from 'lodash/compact';

const SEPARATOR = '/';

function joinFromSet(set, separator = SEPARATOR) {
    if (set.size === 0) {
        return '';
    }

    return [...set].join(separator);
}

export default class Contact {

    constructor() {
        this.email = null;
        this.name = null;
        this.addresses = new Set();
        this.identities = new Set();
        this.nicknames = new Set();
        this.units = new Set();
        this.departments = new Set();
        this.paperCard = false;
        this.annulReport = false;
        this.annulReceipt = false;
        this.er = false;

        this._count = 1;
    }

    static fromObject(object) {
        const email = object.email;
        const name = object.name;
        const address = object.address;

        if (
            ((email === undefined) && (name === '地址貼上的名字')) ||
            ((email === undefined) && (name === undefined) && (address === undefined))
        ) {
            throw new Error('Invalid row');
        }

        const result = new Contact();
        const identity = object.identity;

        result.setEmail(email);
        result.setName(name);
        result.addAddresses([address]);
        result.addIdentities(identity ? identity.split(SEPARATOR) : []);
        result.addNicknames([object.nickname]);
        result.addUnits([object.unit]);
        result.addDepartments([object.department]);
        result.addPaperCard(object.paperCard);
        result.addAnnulReport(object.annualReport);
        result.addAnnulReceipt(object.annualReceipt);
        result.addEr(object.er);

        return result;
    }

    setEmail(email) {
        this.email = (trim(email) || null);
    }

    setName(name) {
        this.name = (trim(name) || null);
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
        this.paperCard = (
            this.paperCard ||
            (paperCard === true) || (paperCard === 'Y') ||
            false
        );
    }

    addAnnulReport(annulReport) {
        this.annulReport = (
            this.annulReport ||
            (annulReport === true) || (annulReport === 'Y') ||
            false
        );
    }

    addAnnulReceipt(annulReceipt) {
        this.annulReceipt = (
            this.annulReceipt ||
            (annulReceipt === true) || (annulReceipt === 'Y') ||
            false
        );
    }

    addEr(er) {
        this.er = (this.er || er || false);
    }

    merge(other) {
        this._count++;
        let result = clone(this);

        if (!result.email) {
            result.setEmail(other.email);
        }

        if (!result.name) {
            result.setName(other.name);
        }

        result.addAddresses([...other.addresses]);
        result.addIdentities([...other.identities]);
        result.addNicknames([...other.nicknames]);
        result.addUnits([...other.units]);
        result.addDepartments([...other.departments]);
        result.addPaperCard(other.paperCard);
        result.addAnnulReport(other.annulReport);
        result.addAnnulReceipt(other.annulReceipt);
        result.addEr(other.er);

        return result;
    }

    toJSON() {
        return {
            email: this.email,
            name: this.name,
            addresses: joinFromSet(this.addresses),
            identities: joinFromSet(this.identities),
            nicknames: joinFromSet(this.nicknames),
            units: joinFromSet(this.units),
            departments: joinFromSet(this.departments),
            paperCard: this.paperCard ? 'Y' : 'N',
            annulReport: this.annulReport ? 'Y' : 'N',
            annulReceipt: this.annulReceipt ? 'Y' : 'N',
            count: this._count,
        };
    }
}

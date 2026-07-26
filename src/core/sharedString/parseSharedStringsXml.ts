import { getElementByName, getElementsByName } from "../utils";


export const parseSharedStringsXml = (str: string): string[] => {
    const strings: string[] = [];
    const xmlDoc = new DOMParser().parseFromString(str, 'text/xml');
    const sstElement = getElementByName(xmlDoc, 'sst');
    const stringItems = getElementsByName(sstElement, 'si');

    stringItems.forEach(item => {
        // A shared string can contain several rich-text runs. Each <si> is one
        // indexed value; flattening every <t> shifts all subsequent indexes.
        strings.push(getElementsByName(item, 't').map(text => text.textContent ?? '').join(''));
    });

    return strings;
};

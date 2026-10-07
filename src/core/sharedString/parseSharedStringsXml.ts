import { parseXml } from '../xml';
import { XlsxLimits } from '../../types';
import { resolveLimits } from '../security';
import { getElementByName, getElementsByName } from "../utils";


export const parseSharedStringsXml = (str: string, limits: XlsxLimits = resolveLimits()): string[] => {
    const strings: string[] = [];
    const xmlDoc = parseXml(str, 'sst', limits);
    const sstElement = getElementByName(xmlDoc, 'sst');
    const stringItems = getElementsByName(sstElement, 'si');

    stringItems.forEach(item => {
        // A shared string can contain several rich-text runs. Each <si> is one
        // indexed value; flattening every <t> shifts all subsequent indexes.
        strings.push(getElementsByName(item, 't').filter(t => t.parentElement?.localName !== 'rPh').map(text => text.textContent ?? '').join(''));
    });

    return strings;
};

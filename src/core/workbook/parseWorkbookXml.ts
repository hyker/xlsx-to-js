import { relationshipId, parseXml } from '../xml';
import { XlsxLimits } from '../../types';
import { boundedNumber, resolveLimits } from '../security';
import { getChildByName, getChildrenByName } from "../utils";
import { Workbook } from "./types";

export const parseWorkbookXml = (str: string, limits: XlsxLimits = resolveLimits()): Workbook => {
    const workbook: Workbook = {
        workSheets: [],
    }; 
    const xmlDoc = parseXml(str, 'workbook', limits);
    const workbookElement = xmlDoc.documentElement;
    const sheetsElement = getChildByName(workbookElement, 'sheets');
    if (!sheetsElement) throw new Error('Workbook is missing sheets');
    const sheetsArray = getChildrenByName(sheetsElement, 'sheet');
    if (!sheetsArray.length) throw new Error('Workbook has no worksheets');

    workbook.date1904 = ['1', 'true'].includes(getChildByName(workbookElement, 'workbookPr')?.getAttribute('date1904') ?? '0');
    if (sheetsArray.length > limits.maxSheets) throw new Error('Sheet count exceeds resource limits');
    const ids = new Set<string>();
    const sheetIds = new Set<number>();
    if (sheetsArray) {
        sheetsArray.forEach(x => {
            const relId = relationshipId(x);
            const sheetId = boundedNumber(x.getAttribute('sheetId'), 'sheet id', 0xffffffff, 1, true);
            const state = x.getAttribute('state') ?? 'visible';
            if (!relId || ids.has(relId) || sheetIds.has(sheetId)) throw new Error('Missing or duplicate sheet relationship');
            if (!['visible', 'hidden', 'veryHidden'].includes(state)) throw new Error('Invalid sheet visibility');
            ids.add(relId);
            sheetIds.add(sheetId);
            workbook.workSheets.push({
                relationshipId: relId,
                state: state as 'visible' | 'hidden' | 'veryHidden',
                id: sheetId,
                name: x.getAttribute('name') ?? '',
                dimention: '',
                data: [],
                columnStyles: [],
                rowStyles: [],
                mergeCells: [],
                drawings: [],
                defaultColWidth: 8.43, // Excel default
                baseColWidth: 8, // typical base width in chars
                defaultRowHeight: 15, // points
                zeroHeight: false,
            });
        });
    }

    return workbook;
};

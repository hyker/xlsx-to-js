import { parseXml } from '../xml';
import { XlsxLimits } from '../../types';
import { resolveLimits } from '../security';

import { getElementByName, getElementsByName } from "../utils";
import { Theme } from "./types";

export const parseThemeXml = (str: string, limits: XlsxLimits = resolveLimits()): Theme[] => {
    const themes: Theme[] = [];
    const xmlDoc = parseXml(str, 'a:theme', limits);
    const themeElement = getElementByName(xmlDoc, 'a:theme');
    const themesElement = getElementByName(themeElement, 'a:themeElements');
    const clrSchemeArray = getElementsByName(themesElement, 'a:clrScheme');

    if (clrSchemeArray) {
        clrSchemeArray.forEach(x => {
            // lt1 (Light 1)
            if (getElementByName(x, 'a:lt1')) {
                const elem = getElementByName(x, 'a:lt1');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 0,
                    name: 'lt1',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // dk1 (Dark 1)
            if (getElementByName(x, 'a:dk1')) {
                const elem = getElementByName(x, 'a:dk1');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? ''

                themes.push({
                    id: 1,
                    name: 'dk1',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // lt2 (Light 2)
            if (getElementByName(x, 'a:lt2')) {
                const elem = getElementByName(x, 'a:lt2');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 2,
                    name: 'lt2',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // dk2 (Dark 2)
            if (getElementByName(x, 'a:dk2')) {
                const elem = getElementByName(x, 'a:dk2');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 3,
                    name: 'dk2',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // accent1 (Accent 1)
            if (getElementByName(x, 'a:accent1')) {
                const elem = getElementByName(x, 'a:accent1');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 4,
                    name: 'accent1',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // accent2 (Accent 2)
            if (getElementByName(x, 'a:accent2')) {
                const elem = getElementByName(x, 'a:accent2');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 5,
                    name: 'accent2',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // accent3 (Accent 3)
            if (getElementByName(x, 'a:accent3')) {
                const elem = getElementByName(x, 'a:accent3');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 6,
                    name: 'accent3',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // accent4 (Accent 4)
            if (getElementByName(x, 'a:accent4')) {
                const elem = getElementByName(x, 'a:accent4');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 7,
                    name: 'accent4',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // accent5 (Accent 5)
            if (getElementByName(x, 'a:accent5')) {
                const elem = getElementByName(x, 'a:accent5');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 8,
                    name: 'accent5',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // accent6 (Accent 6)
            if (getElementByName(x, 'a:accent6')) {
                const elem = getElementByName(x, 'a:accent6');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 9,
                    name: 'accent6',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // hlink (Hyperlink)
            if (getElementByName(x, 'a:hlink')) {
                const elem = getElementByName(x, 'a:hlink');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 10,
                    name: 'hlink',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
            // folHlink (Followed Hyperlink)
            if (getElementByName(x, 'a:folHlink')) {
                const elem = getElementByName(x, 'a:folHlink');
                const sysClr = getElementByName(elem, 'a:sysClr');
                const srgbClr = getElementByName(elem, 'a:srgbClr');
                const rgbColor = sysClr?.getAttribute('lastClr') ?? srgbClr?.getAttribute('val') ?? '';
                themes.push({
                    id: 11,
                    name: 'folHlink',
                    val: rgbColor !== '' ? ('#FF' + rgbColor) : '',
                });
            }
        });
    }

    return themes;
};
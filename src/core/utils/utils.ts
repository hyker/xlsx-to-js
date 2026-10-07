import { matchesName } from '../xml';
import { parseRange, parseReference, resolveLimits } from '../security';
import { XlsxLimits } from '../../types';

function modifyHex(hex: string) {
    if (hex.length == 4) {
      hex = hex.replace('#', '');
    }
    if (hex.length == 3) {
      hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    }
    return hex;
}

export function argbToHex(argb: string) {
    if (argb.startsWith('#')) {
        argb = argb.slice(1);
    }

    const a = parseInt(argb.slice(0, 2), 16) / 255;
    const r = parseInt(argb.slice(2, 4), 16);
    const g = parseInt(argb.slice(4, 6), 16);
    const b = parseInt(argb.slice(6, 8), 16);

    const rFinal = Math.round(r * a);
    const gFinal = Math.round(g * a);
    const bFinal = Math.round(b * a);

    const hex = `#${rFinal.toString(16).padStart(2, '0')}${gFinal.toString(16).padStart(2, '0')}${bFinal.toString(16).padStart(2, '0')}`;

    return hex.toUpperCase();
}

export function hexToRgb(hex: string) {
    const x = { r: 0, g: 0, b: 0 };
    hex = hex.replace('#', '')
    if (hex.length != 6) {
      hex = modifyHex(hex);
    }
    x.r = parseInt(hex.slice(0, 2), 16);
    x.g = parseInt(hex.slice(2, 4), 16);
    x.b = parseInt(hex.slice(4, 6), 16);

    return x;
}
  
export function rgbToHex(r: number, g: number, b: number) {
    return "#" + (1 << 24 | r << 16 | g << 8 | b).toString(16).slice(1);
}

export function getColumnIndex(column: string): number {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
      index *= 26;
      index += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return index;
}

export function getRangeArray<T>(range: string, defValue?: T, limits: XlsxLimits = resolveLimits()): T[][] {
  const { end } = parseRange(range, limits);
  return Array.from({ length: end.row }, () => defValue === undefined ? [] :
    Array.from({ length: end.col }, () => typeof defValue === 'object' ? { ...defValue } : defValue));
}

export function getPositionInArray(cell: string): { row: number, col: number } {
  const pos = parseReference(cell);
  return { row: pos.row - 1, col: pos.col - 1 };
}

export function getElementByName(children?: Element | Document, name?: string): Element | undefined {
  if (!children || !name) return undefined;
  const local = name.split(':').pop()!;
  const elements = children.getElementsByTagNameNS('*', local);
  for (let i = 0; i < elements.length; i++) if (matchesName(elements[i], name)) return elements[i];
  return undefined;
}

export function getElementsByName(children?: Element | Document, name?: string): Element[] {
  if (!children || !name) return [];
  return Array.from(children.getElementsByTagNameNS('*', name.split(':').pop()!)).filter(e => matchesName(e, name));
}

export function getChildrenByName(parent: Element | undefined, name: string): Element[] {
  return parent ? Array.from(parent.children).filter(child => matchesName(child, name)) : [];
}

export function getChildByName(parent: Element | undefined, name: string): Element | undefined {
  return getChildrenByName(parent, name)[0];
}

export function getSheetDimension(sheetData: Element, limits: XlsxLimits = resolveLimits()): string {
  let maxRow = 0, maxCol = 0;
  for (const cell of getElementsByName(sheetData, 'c')) {
    const { row, col } = parseReference(cell.getAttribute('r') ?? '', limits);
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
  }
  if (!maxRow) return '';
  let letters = '';
  for (let n = maxCol; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + (n - 1) % 26) + letters;
  return `A1:${letters}${maxRow}`;
}

export function excelSerialToJSDate(serial: number, date1904 = false): Date {
  // Serial 60 is Excel's fictional 1900-02-29; represent it as 1900-02-28.
  const days = date1904 ? serial : serial < 60 ? serial : serial - 1;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  const result = new Date(epoch + days * 86400000);
  if (!Number.isFinite(result.getTime())) throw new Error('Invalid date serial');
  return result;
}

export function positionFromExt(start: { col: number; row: number }, extValue?: string | null): number {
  const EMU_PER_PIXEL = 9525;
  const extInPx = extValue ? parseInt(extValue, 10) / EMU_PER_PIXEL : 0;
  const CELL_WIDTH_PX = 64;

  return start.col + Math.floor(extInPx / CELL_WIDTH_PX);
}

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

export function getRangeArray<T>(range: string, defValue?: T): T[][] {
  let end: string;
  if (range.includes(':')) {
      [, end] = range.split(':');
  } else {
      end = range;
  }

  const startRow = 1;
  const startColIndex = getColumnIndex('A');

  const endColumn = end.match(/[A-Z]+/)![0];
  const endRow = parseInt(end.match(/\d+/)![0]);

  const endColIndex = getColumnIndex(endColumn);

  const array: T[][] = [];
  for (let row = startRow; row <= endRow; row++) {
      const rowArray: T[] = [];
      for (let col = startColIndex; col <= endColIndex; col++) {
        if (defValue) {
          rowArray.push(defValue);
        } else {
          rowArray.push();
        }
      }
      array.push(rowArray);
  }

  return array;
}

export function getPositionInArray(cell: string): { row: number, col: number } {
  const column = cell.match(/[A-Z]+/)![0];
  const row = parseInt(cell.match(/\d+/)![0]);

  const colIndex = getColumnIndex(column);
  const rowIndex = row - 1;

  return { row: rowIndex, col: colIndex - 1 };
}

export function getElementByName(children?: Element | Document, name?: string) {
  if (children && name) {
    const e = [...children.getElementsByTagName(name)][0];
    const x = [...children.getElementsByTagName(`x:${name}`)][0];
    return e ?? x;
  } else {
      return undefined;
  }
}

export function getElementsByName(children?: Element | Document, name?: string) {
  if (children && name) {
    const e = [...children.getElementsByTagName(name)];
    const x = [...children.getElementsByTagName(`x:${name}`)];
    return ((e && e.length > 0 ? e : undefined) ?? (x && x.length > 0 ? x : undefined)) ?? [];
  } else {
    return [];
  }
}

export function getSheetDimension(sheetData: Element): string {
  const cellElements = sheetData.querySelectorAll("c[r]");
  
  if (cellElements.length === 0) return "";

  let minRow = Number.MAX_VALUE, maxRow = Number.MIN_VALUE;
  let minCol = Number.MAX_VALUE, maxCol = Number.MIN_VALUE;

  const parseCellReference = (ref: string): { col: string, row: number } => {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`Invalid cell reference: ${ref}`);
    const [_, col, row] = match;
    return { col, row: parseInt(row, 10) };
  };

  cellElements.forEach(cell => {
    const cellRef = cell.getAttribute("r");
    if (cellRef) {
      const { col, row } = parseCellReference(cellRef);
      const colIndex = getColumnIndex(col);

      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);

      minCol = Math.min(minCol, colIndex);
      maxCol = Math.max(maxCol, colIndex);
    }
  });

  const columnName = (index: number): string => {
    let name = '';
    while (index > 0) {
      const remainder = (index - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      index = Math.floor((index - 1) / 26);
    }
    return name;
  };

  return `${columnName(minCol)}${minRow}:${columnName(maxCol)}${maxRow}`;
}

export function excelSerialToJSDate(serial: number): Date {
  const excelStartDate = new Date(1900, 0, 1);
  const jsDate = new Date(excelStartDate.getTime() + (serial - 2) * 24 * 60 * 60 * 1000);
  
  return jsDate;
}

export function positionFromExt(start: { col: number; row: number }, extValue?: string | null): number {
  const EMU_PER_PIXEL = 9525;
  const extInPx = extValue ? parseInt(extValue, 10) / EMU_PER_PIXEL : 0;
  const CELL_WIDTH_PX = 64;

  return start.col + Math.floor(extInPx / CELL_WIDTH_PX);
}

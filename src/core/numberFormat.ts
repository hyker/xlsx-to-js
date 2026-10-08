import { excelSerialToJSDate } from './utils';

/** Returns display text, or undefined when the format is unsupported for this value. */
export type NumberFormatter = (value: number) => string | undefined;

/** Built-in formats with fixed codes (ECMA-376 §18.8.30). 14 and 22 follow the
 * locale's short date, like Excel. Currency formats 5-8 and East Asian formats
 * depend on the authoring locale and are left unformatted.
 */
const BUILTIN: Record<number, string> = {
    1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
    15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0',
};

/** Excel rejects longer codes; this also bounds per-cell output from literals. */
const MAX_FORMAT_CODE = 255;

interface LocaleInfo {
    decimal: string; group: string;
    months: string[]; monthsShort: string[]; days: string[]; daysShort: string[];
    shortDate: Intl.DateTimeFormat;
}

const locales = new Map<string, LocaleInfo>();

function localeInfo(locale: string | undefined): LocaleInfo {
    const key = locale ?? '';
    let info = locales.get(key);
    if (info) return info;
    const parts = new Intl.NumberFormat(locale).formatToParts(1234567.5);
    const name = (options: Intl.DateTimeFormatOptions, count: number, date: (i: number) => number) => {
        const format = new Intl.DateTimeFormat(locale, { ...options, timeZone: 'UTC' });
        return Array.from({ length: count }, (_, i) => format.format(date(i)));
    };
    // 2006-01-01 was a Sunday, matching Excel's weekday order.
    const day = (i: number) => Date.UTC(2006, 0, 1 + i), month = (i: number) => Date.UTC(2006, i, 1);
    info = {
        decimal: parts.find(p => p.type === 'decimal')?.value ?? '.',
        group: parts.find(p => p.type === 'group')?.value ?? ',',
        months: name({ month: 'long' }, 12, month), monthsShort: name({ month: 'short' }, 12, month),
        days: name({ weekday: 'long' }, 7, day), daysShort: name({ weekday: 'short' }, 7, day),
        shortDate: new Intl.DateTimeFormat(locale, { timeZone: 'UTC', year: 'numeric', month: 'numeric', day: 'numeric' }),
    };
    locales.set(key, info);
    return info;
}

type Token =
    | { t: 'lit'; v: string }
    | { t: 'digit'; v: '0' | '#' | '?' }
    | { t: 'point' } | { t: 'comma' } | { t: 'percent' } | { t: 'text' } | { t: 'general' }
    | { t: 'exp'; plus: boolean }
    | { t: 'date'; v: string }
    | { t: 'elapsed'; v: 'h' | 'm' | 's'; n: number }
    | { t: 'ampm'; v: string }
    | { t: 'subsec'; n: number };

interface Section { tokens: Token[]; date: boolean }

function splitSections(code: string): string[] {
    const sections: string[] = [];
    let start = 0;
    for (let i = 0; i < code.length; i++) {
        const c = code[i];
        if (c === '"') { const end = code.indexOf('"', i + 1); i = end < 0 ? code.length : end; }
        else if (c === '\\' || c === '_' || c === '*') i++;
        else if (c === '[') { const end = code.indexOf(']', i + 1); i = end < 0 ? code.length : end; }
        else if (c === ';') { sections.push(code.slice(start, i)); start = i + 1; }
    }
    sections.push(code.slice(start));
    return sections;
}

/** Returns undefined for unsupported constructs: conditions and fractions. */
function tokenize(section: string): Section | undefined {
    const tokens: Token[] = [];
    let lastTime: 'h' | 's' | undefined;
    for (let i = 0; i < section.length;) {
        const c = section[i], lower = c.toLowerCase();
        if (c === '"') {
            const end = section.indexOf('"', i + 1);
            if (end < 0) return undefined;
            tokens.push({ t: 'lit', v: section.slice(i + 1, end) }); i = end + 1;
        } else if (c === '\\') { tokens.push({ t: 'lit', v: section[i + 1] ?? '' }); i += 2; }
        else if (c === '_') { tokens.push({ t: 'lit', v: ' ' }); i += 2; }
        else if (c === '*') i += 2;
        else if (c === '[') {
            const end = section.indexOf(']', i + 1);
            if (end < 0) return undefined;
            const body = section.slice(i + 1, end);
            i = end + 1;
            const elapsed = /^(h+|m+|s+)$/i.exec(body);
            if (elapsed) {
                const v = body[0].toLowerCase() as 'h' | 'm' | 's';
                tokens.push({ t: 'elapsed', v, n: body.length });
                if (v !== 'm') lastTime = v;
            } else if (body[0] === '$') {
                const symbol = body.slice(1).split('-')[0];
                if (symbol) tokens.push({ t: 'lit', v: symbol });
            } else if (/^[<>=]/.test(body)) return undefined;
            // Colors and other modifiers do not change the text.
        } else if (section.slice(i, i + 7).toLowerCase() === 'general') { tokens.push({ t: 'general' }); i += 7; }
        else if (section.slice(i, i + 5).toUpperCase() === 'AM/PM') { tokens.push({ t: 'ampm', v: section.slice(i, i + 5) }); i += 5; }
        else if (section.slice(i, i + 3).toUpperCase() === 'A/P') { tokens.push({ t: 'ampm', v: section.slice(i, i + 3) }); i += 3; }
        else if (c === '0' || c === '#' || c === '?') { tokens.push({ t: 'digit', v: c }); i++; }
        else if (c === '.') {
            let n = 0;
            while (section[i + 1 + n] === '0') n++;
            if (lastTime === 's' && n) { tokens.push({ t: 'subsec', n: Math.min(n, 3) }); i += 1 + n; }
            else { tokens.push({ t: 'point' }); i++; }
        } else if (c === ',') { tokens.push({ t: 'comma' }); i++; }
        else if (c === '%') { tokens.push({ t: 'percent' }); i++; }
        else if (c === '@') { tokens.push({ t: 'text' }); i++; }
        else if ((c === 'E' || c === 'e') && (section[i + 1] === '+' || section[i + 1] === '-')) {
            tokens.push({ t: 'exp', plus: section[i + 1] === '+' }); i += 2;
        } else if ('ymdhseb'.includes(lower) || lower === 'g') {
            let n = 1;
            while (section[i + n]?.toLowerCase() === lower) n++;
            if (lower !== 'g' && lower !== 'b') tokens.push({ t: 'date', v: lower.repeat(n) });
            if (lower === 'h' || lower === 's') lastTime = lower;
            i += n;
        } else if (c === '/' && tokens.some(token => token.t === 'digit')) return undefined;
        else { tokens.push({ t: 'lit', v: c }); i++; }
    }
    const date = tokens.some(token => token.t === 'date' || token.t === 'elapsed' || token.t === 'ampm');
    // Digits and separators are literal in date sections ('dd.mm.yyyy'), except
    // fractional seconds handled above.
    const literal = (token: Token): Token => token.t === 'digit' ? { t: 'lit', v: token.v } : token.t === 'point' ? { t: 'lit', v: '.' }
        : token.t === 'comma' ? { t: 'lit', v: ',' } : token.t === 'percent' ? { t: 'lit', v: '%' } : token;
    return { tokens: date ? tokens.map(literal) : tokens, date };
}

/** Rounds in decimal, so 1.005 at two places gives 1.01 like Excel, not 1.00. */
function fixed(value: number, places: number): string | undefined {
    const scaled = Number((value * 10 ** places).toPrecision(15));
    if (!Number.isFinite(scaled) || Math.abs(scaled) >= 1e21) return undefined;
    return (Math.round(scaled) / 10 ** places).toFixed(places);
}

/** Excel's General: up to 11 significant digits, scientific outside 1e-9..1e11. */
function general(value: number): string {
    const abs = Math.abs(value);
    if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
        const [mantissa, exponent] = value.toExponential(5).split('e');
        return `${mantissa.replace(/\.?0+$/, '')}E${exponent[0] === '-' ? '-' : '+'}${exponent.replace(/^[+-]/, '').padStart(2, '0')}`;
    }
    return String(Number(value.toPrecision(11)));
}

function formatNumber(tokens: Token[], value: number, locale: LocaleInfo): string | undefined {
    const exp = tokens.findIndex(token => token.t === 'exp');
    const mantissaEnd = exp < 0 ? tokens.length : exp;
    const point = tokens.findIndex((token, i) => token.t === 'point' && i < mantissaEnd);
    const intEnd = point < 0 ? mantissaEnd : point;
    const intDigits: number[] = [], fracDigits: number[] = [], expDigits: number[] = [];
    tokens.forEach((token, i) => {
        if (token.t !== 'digit') return;
        (i < intEnd ? intDigits : i < mantissaEnd ? fracDigits : expDigits).push(i);
    });
    // A comma between integer placeholders groups thousands; any other comma
    // after a digit placeholder scales by 1,000.
    let grouping = false, scale = 0;
    tokens.forEach((token, i) => {
        if (token.t !== 'comma') return;
        const before = intDigits.some(d => d < i) || fracDigits.some(d => d < i);
        if (intDigits.some(d => d < i) && intDigits.some(d => d > i)) grouping = true;
        else if (before) scale++;
    });
    value = value * 100 ** tokens.filter(token => token.t === 'percent').length / 1000 ** scale;
    let exponent = 0;
    if (exp >= 0) {
        const step = Math.max(1, intDigits.length);
        if (value !== 0) {
            exponent = Math.floor(Math.floor(Math.log10(value)) / step) * step;
            const rounded = Number(fixed(value / 10 ** exponent, fracDigits.length));
            if (rounded >= 10 ** step) exponent += step;
        }
        value /= 10 ** exponent;
    }
    const text = fixed(value, fracDigits.length);
    if (text === undefined) return undefined;
    const [whole, fraction = ''] = text.split('.');
    const output = new Map<number, string>();
    // Integer placeholders fill from the right; surplus digits go to the first.
    let digits = whole === '0' ? '' : whole, cursor = digits.length;
    for (let k = intDigits.length - 1; k >= 0; k--) {
        const placeholder = (tokens[intDigits[k]] as { v: string }).v;
        output.set(intDigits[k], cursor > 0 ? digits[--cursor] : placeholder === '0' ? '0' : placeholder === '?' ? ' ' : '');
    }
    if (intDigits.length && cursor > 0) output.set(intDigits[0], digits.slice(0, cursor) + output.get(intDigits[0]));
    if (grouping && intDigits.length) {
        const joined = intDigits.map(i => output.get(i)).join('');
        const lead = /^ */.exec(joined)![0], body = joined.slice(lead.length);
        output.set(intDigits[0], lead + body.replace(/\B(?=(\d{3})+(?!\d))/g, locale.group));
        for (const i of intDigits.slice(1)) output.set(i, '');
    }
    // Trailing zeros are dropped for '#' and padded for '?'.
    let trailing = true;
    for (let k = fracDigits.length - 1; k >= 0; k--) {
        const placeholder = (tokens[fracDigits[k]] as { v: string }).v, digit = fraction[k];
        if (trailing && digit === '0' && placeholder !== '0') output.set(fracDigits[k], placeholder === '?' ? ' ' : '');
        else { trailing = false; output.set(fracDigits[k], digit); }
    }
    const exponentText = String(Math.abs(exponent)).padStart(expDigits.length, '0');
    expDigits.forEach((i, k) => output.set(i, k ? '' : exponentText));
    let result = '';
    tokens.forEach((token, i) => {
        if (token.t === 'digit') result += output.get(i);
        else if (token.t === 'lit') result += token.v;
        else if (token.t === 'point') result += (!intDigits.length && digits ? digits : '') + locale.decimal;
        else if (token.t === 'percent') result += '%';
        else if (token.t === 'exp') result += `E${exponent < 0 ? '-' : token.plus ? '+' : ''}`;
        else if (token.t === 'general') result += general(value);
    });
    return result;
}

function formatDate(tokens: Token[], serial: number, date1904: boolean, locale: LocaleInfo): string | undefined {
    if (serial < 0 || (!date1904 && serial >= 2958466) || (date1904 && serial >= 2957004)) return undefined;
    const places = Math.max(0, ...tokens.map(token => token.t === 'subsec' ? token.n : 0));
    // Round once at the displayed precision so 23:59:59.6 carries into the next day.
    const total = Math.round(Number((serial * 86400 * 10 ** places).toPrecision(15))) / 10 ** places;
    const days = Math.floor(total / 86400), seconds = total - days * 86400;
    const date = excelSerialToJSDate(days, date1904);
    const hour = Math.floor(seconds / 3600), minute = Math.floor(seconds % 3600 / 60), second = Math.floor(seconds % 60);
    const twelveHour = tokens.some(token => token.t === 'ampm');
    const pad = (n: number, width: number) => String(n).padStart(width, '0');
    // 'm' is minutes after an hour or before a second token, otherwise month.
    const minutes = (i: number): boolean => {
        for (let j = i - 1; j >= 0; j--) {
            const token = tokens[j];
            if (token.t === 'date' || token.t === 'elapsed') { if (token.v[0] === 'h') return true; break; }
        }
        for (let j = i + 1; j < tokens.length; j++) {
            const token = tokens[j];
            if (token.t === 'date' || token.t === 'elapsed') return token.v[0] === 's';
        }
        return false;
    };
    let result = '';
    tokens.forEach((token, i) => {
        if (token.t === 'lit') result += token.v;
        else if (token.t === 'ampm') {
            const pm = hour >= 12, [am, p] = token.v.split('/');
            result += pm ? p.length === 1 ? p : (am === am.toLowerCase() ? 'pm' : 'PM') : am.length === 1 ? am : (am === am.toLowerCase() ? 'am' : 'AM');
        } else if (token.t === 'subsec') result += locale.decimal + pad(Math.round((seconds - Math.floor(seconds)) * 10 ** token.n), token.n);
        else if (token.t === 'elapsed') {
            const value = Math.floor(token.v === 'h' ? total / 3600 : token.v === 'm' ? total / 60 : total);
            result += pad(value, token.n);
        } else if (token.t === 'date') {
            const n = token.v.length;
            switch (token.v[0]) {
                case 'y': result += n <= 2 ? pad(date.getUTCFullYear() % 100, 2) : String(date.getUTCFullYear()); break;
                case 'e': result += String(date.getUTCFullYear()); break;
                case 'd': result += n === 1 ? String(date.getUTCDate()) : n === 2 ? pad(date.getUTCDate(), 2) : n === 3 ? locale.daysShort[date.getUTCDay()] : locale.days[date.getUTCDay()]; break;
                case 'h': { const h = twelveHour ? hour % 12 || 12 : hour; result += n === 1 ? String(h) : pad(h, 2); break; }
                case 's': result += n === 1 ? String(second) : pad(second, 2); break;
                case 'm':
                    if (n <= 2 && minutes(i)) result += n === 1 ? String(minute) : pad(minute, 2);
                    else {
                        const month = date.getUTCMonth();
                        result += n === 1 ? String(month + 1) : n === 2 ? pad(month + 1, 2) : n === 3 ? locale.monthsShort[month]
                            : n === 4 ? locale.months[month] : locale.months[month][0];
                    }
                    break;
            }
        }
    });
    return result;
}

function compileCode(code: string, date1904: boolean, locale: LocaleInfo): NumberFormatter | undefined {
    if (code.length > MAX_FORMAT_CODE) return undefined;
    const parsed = splitSections(code).slice(0, 3).map(tokenize);
    if (parsed.some(section => !section)) return undefined;
    const sections = parsed as Section[];
    // A lone text section ('@') does not format numbers.
    if (sections.length === 1 && sections[0].tokens.some(token => token.t === 'text')) return undefined;
    return value => {
        if (!Number.isFinite(value)) return undefined;
        let section = sections[0], autoMinus = value < 0;
        if (sections.length >= 2 && value < 0) { section = sections[1]; autoMinus = false; }
        else if (sections.length >= 3 && value === 0) section = sections[2];
        if (section.date) return value < 0 && autoMinus ? undefined : formatDate(section.tokens, Math.abs(value), date1904, locale);
        const text = formatNumber(section.tokens, Math.abs(value), locale);
        return text !== undefined && autoMinus && /[1-9]/.test(text) ? `-${text}` : text;
    };
}

/** Compiles a cell number format. Undefined means: show the stored value. */
export function compileNumberFormat(id: number, customCode: string | undefined, date1904: boolean, locale?: string): NumberFormatter | undefined {
    const info = localeInfo(locale);
    // Codes stored in the workbook take precedence over built-in defaults.
    if (customCode === undefined) {
        if (id === 14 || id === 22) {
            const time = compileCode('h:mm', date1904, info)!;
            return value => {
                if (!Number.isFinite(value)) return undefined;
                const day = info.shortDate.format(excelSerialToJSDate(id === 14 ? value : Math.floor(value), date1904));
                return id === 14 ? day : `${day} ${time(value)}`;
            };
        }
        const code = BUILTIN[id];
        return code === undefined ? undefined : compileCode(code, date1904, info);
    }
    return compileCode(customCode, date1904, info);
}

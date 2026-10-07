import { boundedNumber } from '../security';
import { Theme } from "../theme/types";
import { argbToHex, hexToRgb, rgbToHex } from "../utils";

export const getColor = (e: Element | undefined, themes: Theme[]): string => {
    const defaultColor = "#FF000000";
    // undefined
    if (e === undefined) return '';
    // indexed
    if (e.getAttribute('indexed')) {
        switch (e.getAttribute('indexed')) {
            case '0': return argbToHex('#FF000000');
            case '1': return argbToHex('#FFFFFFFF');
            case '2': return argbToHex('#FFFF0000');
            case '3': return argbToHex('#FF00FF00');
            case '4': return argbToHex('#FF0000FF');
            case '5': return argbToHex('#FFFFFF00');
            case '6': return argbToHex('#FFFF00FF');
            case '7': return argbToHex('#FF00FFFF');
            case '8': return argbToHex('#FF000000');
            case '9': return argbToHex('#FFFFFFFF');
            case '10': return argbToHex('#FFFF0000');
            case '11': return argbToHex('#FF00FF00');
            case '12': return argbToHex('#FF0000FF');
            case '13': return argbToHex('#FFFFFF00');
            case '14': return argbToHex('#FFFF00FF');
            case '15': return argbToHex('#FF00FFFF');
            case '16': return argbToHex('#FF800000');
            case '17': return argbToHex('#FF008000');
            case '18': return argbToHex('#FF000080');
            case '19': return argbToHex('#FF808000');
            case '20': return argbToHex('#FF800080');
            case '21': return argbToHex('#FF008080');
            case '22': return argbToHex('#FFC0C0C0');
            case '23': return argbToHex('#FF808080');
            case '24': return argbToHex('#FF9999FF');
            case '25': return argbToHex('#FF993366');
            case '26': return argbToHex('#FFFFFFCC');
            case '27': return argbToHex('#FFCCFFFF');
            case '28': return argbToHex('#FF660066');
            case '29': return argbToHex('#FFFF8080');
            case '30': return argbToHex('#FF0066CC');
            case '31': return argbToHex('#FFCCCCFF');
            case '32': return argbToHex('#FF000080');
            case '33': return argbToHex('#FFFF00FF');
            case '34': return argbToHex('#FFFFFF00');
            case '35': return argbToHex('#FF00FFFF');
            case '36': return argbToHex('#FF800080');
            case '37': return argbToHex('#FF800000');
            case '38': return argbToHex('#FF008080');
            case '39': return argbToHex('#FF0000FF');
            case '40': return argbToHex('#FF00CCFF');
            case '41': return argbToHex('#FFCCFFFF');
            case '42': return argbToHex('#FFCCFFCC');
            case '43': return argbToHex('#FFFFFF99');
            case '44': return argbToHex('#FF99CCFF');
            case '45': return argbToHex('#FFFF99CC');
            case '46': return argbToHex('#FFCC99FF');
            case '47': return argbToHex('#FFFFCC99');
            case '48': return argbToHex('#FF3366FF');
            case '49': return argbToHex('#FF33CCCC');
            case '50': return argbToHex('#FF99CC00');
            case '51': return argbToHex('#FFFFCC00');
            case '52': return argbToHex('#FFFF9900');
            case '53': return argbToHex('#FFFF6600');
            case '54': return argbToHex('#FF666699');
            case '55': return argbToHex('#FF969696');
            case '56': return argbToHex('#FF003366');
            case '57': return argbToHex('#FF339966');
            case '58': return argbToHex('#FF003300');
            case '59': return argbToHex('#FF333300');
            case '60': return argbToHex('#FF993300');
            case '61': return argbToHex('#FF993366');
            case '62': return argbToHex('#FF333399');
            case '63': return argbToHex('#FF333333');
            case '64': return argbToHex('#FF000000'); // Primer plano
            case '65': return argbToHex('#FFFFFFFF'); // Fondo

            default: return argbToHex(defaultColor);
        }
    }
    // rgb
    if (e.getAttribute('rgb')) {
        return argbToHex(e.getAttribute('rgb') ?? defaultColor);
    }
    // theme or tint
    if (e.getAttribute('theme') || e.getAttribute('tint')) {
        const t = themes.find(x => x.id === +(e.getAttribute('theme') ?? '0'));
        const themeArgb = t?.val ?? defaultColor;

        const rgb = hexToRgb(argbToHex(themeArgb));
        const tint = boundedNumber(e.getAttribute('tint') ?? 0, 'color tint', 1, -1);

        if (tint !== 0) {
            if (tint < 0) {
                const factor = 1.0 + tint; // darken
                rgb.r = Math.round(rgb.r * factor);
                rgb.g = Math.round(rgb.g * factor);
                rgb.b = Math.round(rgb.b * factor);
            } else { // tint > 0, lighten
                rgb.r = Math.round(rgb.r * (1.0 - tint) + 255 * tint);
                rgb.g = Math.round(rgb.g * (1.0 - tint) + 255 * tint);
                rgb.b = Math.round(rgb.b * (1.0 - tint) + 255 * tint);
            }
        }

        return rgbToHex(rgb.r, rgb.g, rgb.b);
    }
    
    return argbToHex(defaultColor);
    
};

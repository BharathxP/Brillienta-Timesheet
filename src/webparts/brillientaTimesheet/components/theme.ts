import { PartialTheme } from '@fluentui/react';
import { DM_SANS_WOFF2, MANROPE_WOFF2 } from '../assets/fontData';

export const BODY_FONT = "'DM Sans', 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif";
export const HEADING_FONT = "'Manrope', 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif";
export const NAVY = '#17232e';
export const LIME = '#caf578';

let fontsInjected = false;

/** Register the bundled fonts once per page. */
export function ensureFonts(): void {
  if (fontsInjected || typeof document === 'undefined') return;
  fontsInjected = true;
  const style = document.createElement('style');
  style.setAttribute('data-ed-fonts', '1');
  style.textContent =
    `@font-face{font-family:'DM Sans';font-style:normal;font-weight:400 700;font-display:swap;src:url(data:font/woff2;base64,${DM_SANS_WOFF2}) format('woff2');}` +
    `@font-face{font-family:'Manrope';font-style:normal;font-weight:500 800;font-display:swap;src:url(data:font/woff2;base64,${MANROPE_WOFF2}) format('woff2');}`;
  document.head.appendChild(style);
}

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '');
  const full = h.length === 3 ? h.split('').map((c) => c + c).join('') : h;
  const n = parseInt(full.substring(0, 6), 16);
  if (isNaN(n)) return [23, 35, 46];
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function mix(hex: string, target: [number, number, number], amount: number): string {
  const c = hexToRgb(hex);
  const out = c.map((v, i) => Math.round(v + (target[i] - v) * amount));
  return '#' + out.map((v) => (v < 16 ? '0' : '') + v.toString(16)).join('');
}

/** Fluent UI theme derived from the brand colour, matching the app's flat, navy-and-lime look. */
export function brandTheme(brand: string): PartialTheme {
  const white: [number, number, number] = [255, 255, 255];
  const black: [number, number, number] = [0, 0, 0];
  return {
    palette: {
      themePrimary: brand,
      themeLighterAlt: mix(brand, white, 0.95),
      themeLighter: mix(brand, white, 0.88),
      themeLight: mix(brand, white, 0.75),
      themeTertiary: mix(brand, white, 0.45),
      themeSecondary: mix(brand, white, 0.15),
      themeDarkAlt: mix(brand, white, 0.08),
      themeDark: mix(brand, black, 0.2),
      themeDarker: mix(brand, black, 0.4),
      neutralPrimary: NAVY,
      neutralSecondary: '#4b5862',
      neutralTertiary: '#8a959c',
      neutralTertiaryAlt: '#c8d0d2',
      neutralQuaternaryAlt: '#e3e8e9',
      neutralLight: '#eaeeef',
      neutralLighter: '#f2f5f5',
      neutralLighterAlt: '#f8faf9'
    },
    defaultFontStyle: { fontFamily: BODY_FONT },
    effects: { roundedCorner2: '7px', roundedCorner4: '9px', roundedCorner6: '12px' },
    semanticColors: {
      inputBorder: '#d3dadb',
      inputBorderHovered: '#9aa6a8',
      buttonBorder: '#d3dadb',
      bodyDivider: '#dfe5e6',
      menuItemBackgroundHovered: '#eef3e8'
    },
    components: {
      PrimaryButton: { styles: { root: { borderRadius: 7, fontWeight: 600 }, label: { fontWeight: 600 } } },
      DefaultButton: { styles: { root: { borderRadius: 7 }, label: { fontWeight: 600 } } }
    }
  };
}

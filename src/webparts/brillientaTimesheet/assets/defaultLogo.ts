// Default logo: a clock (white face, navy and lime hands) on the Brillienta green tile, the same tile style as the
// Employee Directory's badge. Used until an admin uploads their own logo (Admin › General › Branding).

export const BRAND_GREEN = '#038f62';
export const BRAND_NAVY = '#17232e';
export const BRAND_LIME = '#caf578';

/** The clock tile as SVG. */
export function clockTileSvg(background: string = BRAND_GREEN, face: string = '#ffffff', hands: string = BRAND_NAVY, accent: string = BRAND_LIME): string {
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120">` +
    `<rect width="120" height="120" rx="28" fill="${background}"/>` +
    `<circle cx="60" cy="62" r="34" fill="${face}"/>` +
    `<rect x="54" y="16" width="12" height="10" rx="3" fill="${accent}"/>` +
    `<path d="M60 62 V40" stroke="${hands}" stroke-width="7" stroke-linecap="round"/>` +
    `<path d="M60 62 L76 72" stroke="${hands}" stroke-width="7" stroke-linecap="round"/>` +
    `<circle cx="60" cy="62" r="6" fill="${accent}" stroke="${hands}" stroke-width="3"/>` +
    `</svg>`
  );
}

export const DEFAULT_LOGO = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(clockTileSvg());

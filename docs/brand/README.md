# Brillienta Timesheet logo

The clock: a white clock face with navy hands and a lime crown, on a green tile.

| File | Use |
| --- | --- |
| `brillienta-timesheet-logo.svg` | **The logo**: white clock, navy (`#17232e`) hands, lime (`#caf578`) crown, on a green (`#038f62`) rounded square |
| `brillienta-timesheet-logo-512.png` | The logo as a 512 px bitmap, transparent corners (marketplace listing, website) |
| `brillienta-timesheet-badge.svg` | The clock alone, outlined in green, where a tile does not fit, e.g. inline in text |

The package icons are drawn from the same shapes: `sharepoint/images/app-icon.png` (96 px, app catalog and the web part
picker), `teams/<web part id>_color.png` (192 px) and `_outline.png` (32 px, white on transparent). Regenerate them with
`node tools/icons/make-icons.cjs`, and these brand files with `node tools/icons/make-icons.cjs --brand`. If the
app icon changes, paste its new base64 into `iconImageUrl` in the web part manifest (`tests/teams.test.cjs` checks it).
Customers can still upload their own logo under Admin › General & branding.

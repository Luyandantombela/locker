# Cell Protector — Excel Add-in

A modern Excel add-in that lets you protect only the cells you care about — by selection or by formula — with optional password protection and editable-cell highlighting.

## Features

- **Lock Selected Range** — locks the cells you have highlighted; unlocks everything else
- **Lock All Formula Cells** — auto-detects every formula and locks it; unlocks the rest
- **Password Protection** — optionally protect the sheet with a password
- **Highlight Editable Cells** — applies a light green fill to all editable cells, stores original colors, and restores them on unprotect
- **Remove Protection** — detects sheet protection and removes it (with or without a password); restores highlighted cell colors automatically

## Tech Stack

- HTML · CSS · JavaScript
- [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office) (Microsoft Office Add-in API)
- No backend — fully static, hosted on GitHub Pages

## File Structure

```
cell-protector/
├── index.html       # Taskpane HTML (loaded inside Excel)
├── taskpane.css     # Styles — clean, green-accented, modern
├── taskpane.js      # All Office.js logic
├── manifest.xml     # Office Add-in manifest (points to GitHub Pages URL)
└── README.md
```

## GitHub Pages Hosting

This add-in is hosted at:

```
https://luyandantombela.github.io/locker/cell-protector/index.html
```

Enable GitHub Pages in your repo settings (Settings → Pages → Branch: main → `/root` or `/docs`).

## Sideloading the Add-in

### Excel on Windows / Mac

1. In Excel, go to **Insert → Add-ins → My Add-ins → Upload My Add-in**
2. Browse to `manifest.xml` and upload it
3. The **Cell Protector** button will appear on the **Home** tab

### Excel on the Web

1. Go to **Insert → Add-ins → Upload My Add-in**
2. Browse to and upload `manifest.xml`

## How to Use

1. **Protect by selection** — select any range in Excel, choose *Lock Selected Range*, then click **Protect Cells**
2. **Protect formulas** — choose *Lock All Formula Cells*, then click **Protect Cells** — the add-in finds and locks every formula automatically
3. **Password** — toggle *Enable Password Protection* and enter a password before clicking Protect
4. **Highlight** — check *Highlight Editable Cells* to visually mark which cells remain editable (colors are restored on unprotect)
5. **Unprotect** — click **Detect & Remove Protection**; if a password is required, enter it when prompted

## Future Extensibility

The code is structured for easy addition of:
- Lock specific columns
- Lock specific rows
- Lock by named ranges
- Protect multiple worksheets
- Protection templates

## License

MIT

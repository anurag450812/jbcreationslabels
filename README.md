# JB Creations - Shipping Label Tools

A desktop-capable label toolkit for sorting, cropping, counting, finding, and SKU to new3 automation for Flipkart, Amazon, and Meesho workflows.

## Features

- 📤 **Drag & Drop Upload**: Easy file upload with drag-and-drop support
- 🔍 **Smart Detection**: Automatically finds priority labels in PDF documents
- 📊 **Real-time Processing**: Shows progress while processing your files
- ⬇️ **Sorted Output**: Downloads a single PDF with priority labels at the beginning
- 🎨 **Beautiful UI**: Modern, responsive design that works on all devices
- 🏷️ **Platform Support**: Supports Flipkart, Amazon, and Meesho labels
- 📦 **SKU To New3 Automation**: Desktop-only workflow that ports the attached Python process into the app

## SKU To New3 Automation

The new tab is designed for the Windows desktop build of the app and mirrors the Python workflow:

- Drag and drop or browse `.xlsx`, `.xls`, `.csv`, and `.txt` order files
- Or scan a configured source folder for files modified today
- Extract SKUs using configurable column aliases
- Apply Meesho filtering by `Reason for Credit Entry` values
- Write combined SKUs to the `ORDERS` tab in Google Sheets using a stored service account JSON
- Wait for pivot refresh and export `STOCK ANALYSIS!C:D` to the configured `sku.csv`
- Clear the configured `new3` folder
- Copy SKU-named PDFs from a source PDF folder into `new3` using the exported quantity values
- Show per-platform counts, moved/skipped files, missing SKUs, and a detailed log

Settings editing is password-protected with `200274`. The desktop profile stores the imported Google service account JSON so the workflow can be restored after reinstalling the app.

## How to Use

1. Install dependencies with `npm install`
2. Start the Windows desktop app with `npm start`
3. Open the `SKU To New3 Automation` tab in the desktop app for the local automation workflow
4. Import your Google service account JSON from the tab
5. Unlock settings with password `200274` and configure your source, archive, CSV, and destination paths
6. Choose manual files or folder-scan mode
7. Click `Run SKU Automation`

For the PDF-only browser features, you can still open `index.html` directly, but the local folder automation requires the Electron desktop runtime.

## Desktop Runtime

This project now runs as an Electron app so the UI can keep its website-style workflow while still accessing local folders, files, and Google credentials.

- `npm start`: launch the desktop app locally
- `npm run build:win`: build a Windows portable package

The Netlify site can still act as the recovery/download page, but the SKU automation flow itself runs through the desktop build.

## Existing Label Tools

1. Open `index.html` in your web browser
2. (Optional) Select your platform (Flipkart/Amazon/Meesho) for custom cropping logic
3. Upload one or multiple PDF files containing shipping labels
4. Click "Process & Sort Labels"
5. Wait for processing to complete
6. Download your sorted PDF with priority labels at the beginning

## Priority Labels

The application sorts labels based on this priority list:
- 0baby series (0baby_1, 0baby_2, etc.)
- bd series (bd18, bd21, bd25, etc.)
- ch- prefixed series (ch-bd58, ch-fk21, etc.)
- fk series (fk04, etc.)
- gn series (gn31, gn35, etc.)
- hanuman series (hanuman01, hanuman02, hanuman03)
- jesus series (jesus08, jesus26, jesus27)
- And many more...

See the full list in the application's "View Priority Label List" section.

## Technology Stack

- **Frontend**: HTML5, CSS3, JavaScript (ES6+)
- **PDF Processing**: 
  - PDF.js - For reading and extracting text from PDFs
  - PDF-lib - For creating and manipulating PDFs
- **Styling**: Modern CSS with gradients and shadows

## Adding Custom Cropping Logic

The application has placeholders for platform-specific cropping logic. To add custom cropping:

1. Open `script.js`
2. Find the `croppingLogic` object at the bottom
3. Implement the `cropPage` function for each platform:

```javascript
flipkart: {
    cropPage: async (page) => {
        // Add your Flipkart-specific cropping logic here
        // Example: crop to specific dimensions
        return page;
    }
}
```

## File Structure

```
jbcreationslabels/
├── index.html                 # Main renderer HTML
├── styles.css                 # Styling
├── script.js                  # Renderer logic for all tabs
├── src/
│   ├── main.js                # Electron main process
│   ├── preload.js             # Safe desktop bridge for the renderer
│   └── sku-automation/
│       ├── config-store.js    # Desktop profile settings and JSON import
│       └── local-automation.js# Python-parity SKU automation logic
└── README.md                  # This file
```

## Browser Compatibility

- ✅ Chrome/Edge (Recommended)
- ✅ Firefox
- ✅ Safari
- ✅ Opera

## Notes

- The PDF tools still work in the browser, but the SKU automation tab requires the Electron desktop build
- Files are processed locally and never sent through a custom backend server
- Supports multiple PDF files at once
- Maintains original page quality

## Future Enhancements

- [ ] Implement platform-specific cropping logic
- [ ] Add support for more label formats
- [ ] Batch processing optimization
- [ ] Export sorting report
- [ ] Dark mode support

## Support

For issues or questions, please create an issue in the repository.

---

**© 2025 JB Creations. Built for efficient label management.**

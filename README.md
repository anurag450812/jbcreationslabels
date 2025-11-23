# JB Creations - Shipping Label Sorter

A web application to automatically sort shipping labels from Flipkart, Amazon, and Meesho based on priority codes.

## Features

- 📤 **Drag & Drop Upload**: Easy file upload with drag-and-drop support
- 🔍 **Smart Detection**: Automatically finds priority labels in PDF documents
- 📊 **Real-time Processing**: Shows progress while processing your files
- ⬇️ **Sorted Output**: Downloads a single PDF with priority labels at the beginning
- 🎨 **Beautiful UI**: Modern, responsive design that works on all devices
- 🏷️ **Platform Support**: Supports Flipkart, Amazon, and Meesho labels

## How to Use

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
├── index.html      # Main HTML file
├── styles.css      # Styling
├── script.js       # Application logic
└── README.md       # This file
```

## Browser Compatibility

- ✅ Chrome/Edge (Recommended)
- ✅ Firefox
- ✅ Safari
- ✅ Opera

## Notes

- The application runs entirely in the browser (no server required)
- Your files are processed locally and never uploaded to any server
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

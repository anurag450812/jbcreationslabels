import fs from 'fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

const pdfPath = 'C:\\Users\\anura\\OneDrive\\Desktop\\meesho mulit\\New folder\\Flipkart-Labels-15-Feb-2026-03-05-Copy.pdf';
const data = new Uint8Array(fs.readFileSync(pdfPath));

getDocument({data}).promise.then(async pdf => {
  console.log('Total pages:', pdf.numPages);
  
  // Analyze first 2 pages
  for (let pn = 1; pn <= Math.min(2, pdf.numPages); pn++) {
    const page = await pdf.getPage(pn);
    const vp = page.getViewport({scale:1});
    console.log('\n=== Page ' + pn + ' ===');
    console.log('Viewport:', vp.width, 'x', vp.height);
    
    const tc = await page.getTextContent();
    
    console.log('\nAll text items with positions:');
    for (const item of tc.items) {
      const s = item.str.trim();
      if (!s) continue;
      const y = item.transform ? item.transform[5] : 0;
      const x = item.transform ? item.transform[4] : 0;
      console.log('  Y=' + y.toFixed(1) + ' X=' + x.toFixed(1) + ' : "' + s.substring(0,80) + '"');
    }
    
    // Find dashed lines
    console.log('\nDashed/separator items:');
    for (const item of tc.items) {
      const s = item.str;
      if (s && s.length > 4) {
        const hasDash = (s.match(/-/g) || []).length > 4;
        if (hasDash) {
          const y = item.transform ? item.transform[5] : 0;
          console.log('  Dashes at Y=' + y.toFixed(1) + ' len=' + s.length + ' : "' + s.substring(0,60) + '"');
        }
      }
    }
  }
}).catch(e => console.error('Error:', e.message));

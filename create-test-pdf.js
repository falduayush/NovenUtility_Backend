const fs = require('fs-extra');
const path = require('path');
const puppeteer = require('puppeteer');

async function createTestPdf() {
  console.log('Creating test PDF file...');
  
  const tempDir = path.join(__dirname, 'temp');
  await fs.ensureDir(tempDir);
  
  const outputPath = path.join(tempDir, 'test-document.pdf');
  
  try {
    const browser = await puppeteer.launch({ 
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    
    // Create HTML content for the test PDF
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Test Document</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: #333;
            margin: 40px;
          }
          
          h1 {
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
          }
          
          h2 {
            color: #34495e;
            margin-top: 30px;
          }
          
          table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
          }
          
          th, td {
            border: 1px solid #ddd;
            padding: 8px 12px;
            text-align: left;
          }
          
          th {
            background-color: #f8f9fa;
            font-weight: bold;
          }
          
          .highlight {
            background-color: #fff3cd;
            padding: 10px;
            border-left: 4px solid #ffc107;
            margin: 20px 0;
          }
          
          ul {
            margin: 10px 0;
            padding-left: 20px;
          }
          
          li {
            margin: 5px 0;
          }
        </style>
      </head>
      <body>
        <h1>Sample Document for PDF to Word Conversion</h1>
        
        <p>This is a test document created to verify PDF to Word conversion functionality. It contains various elements that should be properly converted.</p>
        
        <h2>Document Features</h2>
        <ul>
          <li>Multiple headings and subheadings</li>
          <li>Formatted text with different styles</li>
          <li>Tables with structured data</li>
          <li>Lists and bullet points</li>
          <li>Highlighted sections</li>
        </ul>
        
        <h2>Sample Table</h2>
        <table>
          <thead>
            <tr>
              <th>Name</th>
              <th>Position</th>
              <th>Department</th>
              <th>Salary</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>John Doe</td>
              <td>Software Engineer</td>
              <td>Engineering</td>
              <td>$85,000</td>
            </tr>
            <tr>
              <td>Jane Smith</td>
              <td>Product Manager</td>
              <td>Product</td>
              <td>$95,000</td>
            </tr>
            <tr>
              <td>Mike Johnson</td>
              <td>Designer</td>
              <td>Design</td>
              <td>$75,000</td>
            </tr>
          </tbody>
        </table>
        
        <div class="highlight">
          <strong>Important Note:</strong> This document is designed to test the conversion capabilities of the PDF to Word tool. The conversion should preserve formatting, tables, and text structure.
        </div>
        
        <h2>Technical Requirements</h2>
        <p>For optimal conversion results, ensure that:</p>
        <ul>
          <li>LibreOffice is installed on the system</li>
          <li>The PDF contains text (not just images)</li>
          <li>The document structure is clear and well-formatted</li>
          <li>Tables and lists are properly structured</li>
        </ul>
        
        <h2>Conversion Quality</h2>
        <p>The quality of the conversion depends on several factors:</p>
        <ul>
          <li><strong>Text-based PDFs:</strong> Best results with high accuracy</li>
          <li><strong>Scanned documents:</strong> May require OCR processing</li>
          <li><strong>Complex layouts:</strong> May need manual adjustment</li>
          <li><strong>Images and graphics:</strong> Should be preserved as embedded objects</li>
        </ul>
        
        <p style="margin-top: 40px; text-align: center; color: #666; font-style: italic;">
          Generated on ${new Date().toLocaleDateString()} for testing purposes.
        </p>
      </body>
      </html>
    `;

    await page.setContent(htmlContent);
    
    // Generate PDF
    await page.pdf({
      path: outputPath,
      format: 'A4',
      margin: { top: '0.5in', right: '0.5in', bottom: '0.5in', left: '0.5in' },
      printBackground: true
    });

    await browser.close();
    
    const stats = fs.statSync(outputPath);
    console.log(`✅ Test PDF created successfully!`);
    console.log(`📁 Location: ${outputPath}`);
    console.log(`📄 Size: ${(stats.size / 1024).toFixed(2)} KB`);
    console.log(`\n💡 You can now use this PDF to test the PDF to Word conversion tool.`);
    
  } catch (error) {
    console.error('❌ Failed to create test PDF:', error.message);
  }
}

// Run the function
createTestPdf().catch(console.error);


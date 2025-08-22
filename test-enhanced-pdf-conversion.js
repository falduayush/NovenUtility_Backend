const fs = require('fs-extra');
const path = require('path');
const PdfToWordConverter = require('./utils/pdfToWordConverter');

async function testEnhancedPdfConversion() {
  console.log('=== Enhanced PDF to Word Conversion Test ===\n');
  
  // Test 1: Check LibreOffice availability
  console.log('1. Checking LibreOffice availability...');
  try {
    const libreOfficeAvailable = await PdfToWordConverter.isLibreOfficeAvailable();
    console.log(`   LibreOffice available: ${libreOfficeAvailable ? 'YES' : 'NO'}`);
    if (libreOfficeAvailable) {
      console.log('   ✅ LibreOffice will be used for best conversion quality');
    } else {
      console.log('   ⚠️  LibreOffice not found. Will use alternative methods.');
      console.log('   💡 Install LibreOffice for best results: https://www.libreoffice.org/download/');
    }
  } catch (error) {
    console.log(`   ❌ Error checking LibreOffice: ${error.message}`);
  }
  
  // Test 2: Create a test PDF with complex formatting
  console.log('\n2. Creating test PDF with complex formatting...');
  const tempDir = path.join(__dirname, 'temp');
  await fs.ensureDir(tempDir);
  
  const testPdfPath = path.join(tempDir, 'test-complex-document.pdf');
  
  try {
    const puppeteer = require('puppeteer');
    const browser = await puppeteer.launch({ 
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    
    // Create a complex HTML document with tables, headers, footers, etc.
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Complex Test Document</title>
        <style>
          @page {
            size: A4;
            margin: 1in;
          }
          
          body {
            font-family: Arial, sans-serif;
            font-size: 12px;
            line-height: 1.4;
            color: #333;
          }
          
          .header {
            text-align: center;
            border-bottom: 2px solid #333;
            padding-bottom: 10px;
            margin-bottom: 20px;
          }
          
          .footer {
            text-align: center;
            border-top: 1px solid #ccc;
            padding-top: 10px;
            margin-top: 20px;
            font-size: 10px;
            color: #666;
          }
          
          h1 {
            color: #2c3e50;
            font-size: 24px;
            margin-bottom: 10px;
          }
          
          h2 {
            color: #34495e;
            font-size: 18px;
            margin-top: 20px;
            margin-bottom: 10px;
          }
          
          table {
            border-collapse: collapse;
            width: 100%;
            margin: 15px 0;
            page-break-inside: avoid;
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
            margin: 15px 0;
          }
          
          ul {
            margin: 10px 0;
            padding-left: 20px;
          }
          
          li {
            margin: 5px 0;
          }
          
          .page-break {
            page-break-before: always;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>COMPANY REPORT</h1>
          <p>Generated on ${new Date().toLocaleDateString()}</p>
        </div>
        
        <h2>Executive Summary</h2>
        <p>This document contains various formatting elements to test PDF to Word conversion quality, including tables, headers, footers, and complex layouts.</p>
        
        <div class="highlight">
          <strong>Important:</strong> This document is designed to test the enhanced conversion capabilities of the PDF to Word tool.
        </div>
        
        <h2>Financial Data</h2>
        <table>
          <thead>
            <tr>
              <th>Department</th>
              <th>Budget</th>
              <th>Actual</th>
              <th>Variance</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Engineering</td>
              <td>$500,000</td>
              <td>$485,000</td>
              <td>-$15,000</td>
              <td>Under Budget</td>
            </tr>
            <tr>
              <td>Marketing</td>
              <td>$300,000</td>
              <td>$320,000</td>
              <td>+$20,000</td>
              <td>Over Budget</td>
            </tr>
            <tr>
              <td>Sales</td>
              <td>$750,000</td>
              <td>$780,000</td>
              <td>+$30,000</td>
              <td>Over Budget</td>
            </tr>
            <tr>
              <td>Operations</td>
              <td>$400,000</td>
              <td>$395,000</td>
              <td>-$5,000</td>
              <td>Under Budget</td>
            </tr>
          </tbody>
        </table>
        
        <h2>Key Performance Indicators</h2>
        <ul>
          <li><strong>Revenue Growth:</strong> 15% year-over-year</li>
          <li><strong>Customer Satisfaction:</strong> 4.5/5.0</li>
          <li><strong>Employee Retention:</strong> 92%</li>
          <li><strong>Market Share:</strong> 23%</li>
        </ul>
        
        <div class="page-break"></div>
        
        <h2>Detailed Analysis</h2>
        <p>This section contains detailed analysis of the company's performance across various metrics and departments.</p>
        
        <h3>Department Performance</h3>
        <table>
          <thead>
            <tr>
              <th>Department</th>
              <th>Efficiency</th>
              <th>Quality</th>
              <th>Innovation</th>
              <th>Overall Score</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Engineering</td>
              <td>95%</td>
              <td>98%</td>
              <td>92%</td>
              <td>95%</td>
            </tr>
            <tr>
              <td>Marketing</td>
              <td>88%</td>
              <td>85%</td>
              <td>90%</td>
              <td>88%</td>
            </tr>
            <tr>
              <td>Sales</td>
              <td>92%</td>
              <td>89%</td>
              <td>87%</td>
              <td>89%</td>
            </tr>
          </tbody>
        </table>
        
        <h3>Recommendations</h3>
        <ol>
          <li>Increase marketing budget allocation for Q4</li>
          <li>Implement new efficiency measures in operations</li>
          <li>Enhance employee training programs</li>
          <li>Expand market presence in emerging regions</li>
        </ol>
        
        <div class="footer">
          <p>Confidential - Internal Use Only</p>
          <p>Page 1 of 2 | Generated by FileFlow PDF to Word Converter</p>
        </div>
      </body>
      </html>
    `;

    await page.setContent(htmlContent);
    
    // Generate PDF
    await page.pdf({
      path: testPdfPath,
      format: 'A4',
      margin: { top: '0.5in', right: '0.5in', bottom: '0.5in', left: '0.5in' },
      printBackground: true,
      displayHeaderFooter: false
    });

    await browser.close();
    
    const stats = fs.statSync(testPdfPath);
    console.log(`   ✅ Complex test PDF created successfully!`);
    console.log(`   📁 Location: ${testPdfPath}`);
    console.log(`   📄 Size: ${(stats.size / 1024).toFixed(2)} KB`);
    
  } catch (error) {
    console.error('   ❌ Failed to create test PDF:', error.message);
    return;
  }
  
  // Test 3: Test enhanced PDF to Word conversion
  console.log('\n3. Testing enhanced PDF to Word conversion...');
  const outputDocxPath = path.join(tempDir, 'test-converted-document.docx');
  
  try {
    await PdfToWordConverter.convertPdfToWord(testPdfPath, outputDocxPath);
    
    const outputStats = fs.statSync(outputDocxPath);
    console.log(`   ✅ Enhanced conversion successful!`);
    console.log(`   📄 Output file size: ${(outputStats.size / 1024).toFixed(2)} KB`);
    console.log(`   📁 Saved to: ${outputDocxPath}`);
    
    if (outputStats.size < 1000) {
      console.log('   ⚠️  Output file seems very small, conversion might have issues.');
    }
    
  } catch (error) {
    console.log(`   ❌ Enhanced conversion failed: ${error.message}`);
  }
  
  // Summary
  console.log('\n=== Test Summary ===');
  console.log('📁 Check the following files in the temp directory:');
  console.log('   - test-complex-document.pdf (input PDF with complex formatting)');
  console.log('   - test-converted-document.docx (converted Word document)');
  
  console.log('\n💡 What to check in the converted document:');
  console.log('1. Headers and footers are preserved');
  console.log('2. Tables maintain their structure and formatting');
  console.log('3. Text formatting (bold, italic, etc.) is preserved');
  console.log('4. Page breaks are maintained');
  console.log('5. Lists and bullet points are properly formatted');
  console.log('6. Overall layout and spacing is similar to original');
  
  console.log('\n🔧 Conversion Methods Used:');
  console.log('1. LibreOffice (best for headers, footers, tables)');
  console.log('2. pdf2docx (good for table preservation)');
  console.log('3. Enhanced text extraction (fallback with formatting detection)');
  
  console.log('\n✅ Enhanced conversion test completed!');
}

// Run the test
testEnhancedPdfConversion().catch(console.error);


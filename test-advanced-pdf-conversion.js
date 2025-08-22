const fs = require('fs-extra');
const path = require('path');
const AdvancedPdfToWordConverter = require('./utils/advancedPdfToWordConverter');

async function testAdvancedPdfConversion() {
  console.log('=== Advanced PDF to Word Conversion Test ===\n');
  
  // Test 1: Check LibreOffice availability
  console.log('1. Checking LibreOffice availability...');
  try {
    const libreOfficeAvailable = await AdvancedPdfToWordConverter.isLibreOfficeAvailable();
    console.log(`   LibreOffice available: ${libreOfficeAvailable ? 'YES' : 'NO'}`);
    if (libreOfficeAvailable) {
      console.log('   ✅ LibreOffice will be used for maximum formatting preservation');
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
  
  const testPdfPath = path.join(tempDir, 'test-advanced-document.pdf');
  
  try {
    const puppeteer = require('puppeteer');
    const browser = await puppeteer.launch({ 
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    
    // Create a complex HTML document with advanced formatting
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Advanced Test Document</title>
        <style>
          @page {
            size: A4;
            margin: 1in;
          }
          
          body {
            font-family: 'Times New Roman', serif;
            font-size: 12px;
            line-height: 1.5;
            color: #000;
          }
          
          .header {
            text-align: center;
            border-bottom: 3px solid #000;
            padding-bottom: 15px;
            margin-bottom: 25px;
            font-family: Arial, sans-serif;
          }
          
          .footer {
            text-align: center;
            border-top: 2px solid #333;
            padding-top: 15px;
            margin-top: 25px;
            font-size: 10px;
            color: #666;
            font-family: Arial, sans-serif;
          }
          
          h1 {
            color: #1a1a1a;
            font-size: 28px;
            font-weight: bold;
            margin-bottom: 15px;
            text-transform: uppercase;
            letter-spacing: 2px;
          }
          
          h2 {
            color: #2c3e50;
            font-size: 20px;
            font-weight: bold;
            margin-top: 25px;
            margin-bottom: 15px;
            border-bottom: 1px solid #ccc;
            padding-bottom: 5px;
          }
          
          h3 {
            color: #34495e;
            font-size: 16px;
            font-weight: bold;
            margin-top: 20px;
            margin-bottom: 10px;
          }
          
          table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
            page-break-inside: avoid;
            font-family: Arial, sans-serif;
          }
          
          th, td {
            border: 2px solid #333;
            padding: 12px 15px;
            text-align: left;
            vertical-align: top;
          }
          
          th {
            background-color: #f8f9fa;
            font-weight: bold;
            font-size: 14px;
            color: #333;
          }
          
          td {
            font-size: 12px;
          }
          
          .highlight {
            background-color: #fff3cd;
            padding: 15px;
            border-left: 5px solid #ffc107;
            margin: 20px 0;
            font-weight: bold;
          }
          
          .important {
            background-color: #d4edda;
            padding: 10px;
            border: 2px solid #28a745;
            margin: 15px 0;
            font-weight: bold;
            color: #155724;
          }
          
          ul {
            margin: 15px 0;
            padding-left: 25px;
          }
          
          li {
            margin: 8px 0;
            line-height: 1.6;
          }
          
          ol {
            margin: 15px 0;
            padding-left: 25px;
          }
          
          .page-break {
            page-break-before: always;
          }
          
          .text-center {
            text-align: center;
          }
          
          .text-right {
            text-align: right;
          }
          
          .bold {
            font-weight: bold;
          }
          
          .italic {
            font-style: italic;
          }
          
          .underline {
            text-decoration: underline;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>ADVANCED TEST DOCUMENT</h1>
          <p class="bold">Formatting Preservation Test</p>
          <p class="italic">Generated on ${new Date().toLocaleDateString()}</p>
        </div>
        
        <h2>Executive Summary</h2>
        <p>This document is designed to test the <span class="bold">advanced formatting preservation</span> capabilities of the PDF to Word converter. It includes various formatting elements such as:</p>
        
        <div class="highlight">
          <strong>Key Features to Test:</strong>
          <ul>
            <li>Headers and footers preservation</li>
            <li>Font styles and sizes</li>
            <li>Table formatting and structure</li>
            <li>Text alignment and spacing</li>
            <li>Lists and bullet points</li>
            <li>Page breaks and layout</li>
          </ul>
        </div>
        
        <div class="important">
          <strong>Important:</strong> This document contains complex formatting that should be preserved during conversion.
        </div>
        
        <h2>Financial Performance Analysis</h2>
        <table>
          <thead>
            <tr>
              <th>Department</th>
              <th>Budget Allocation</th>
              <th>Actual Spending</th>
              <th>Variance</th>
              <th>Performance Status</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td class="bold">Engineering</td>
              <td class="text-right">$500,000</td>
              <td class="text-right">$485,000</td>
              <td class="text-right">-$15,000</td>
              <td><span class="bold" style="color: green;">Under Budget</span></td>
            </tr>
            <tr>
              <td class="bold">Marketing</td>
              <td class="text-right">$300,000</td>
              <td class="text-right">$320,000</td>
              <td class="text-right">+$20,000</td>
              <td><span class="bold" style="color: red;">Over Budget</span></td>
            </tr>
            <tr>
              <td class="bold">Sales</td>
              <td class="text-right">$750,000</td>
              <td class="text-right">$780,000</td>
              <td class="text-right">+$30,000</td>
              <td><span class="bold" style="color: red;">Over Budget</span></td>
            </tr>
            <tr>
              <td class="bold">Operations</td>
              <td class="text-right">$400,000</td>
              <td class="text-right">$395,000</td>
              <td class="text-right">-$5,000</td>
              <td><span class="bold" style="color: green;">Under Budget</span></td>
            </tr>
          </tbody>
        </table>
        
        <h2>Key Performance Indicators</h2>
        <ol>
          <li><strong>Revenue Growth:</strong> <span class="bold">15%</span> year-over-year</li>
          <li><strong>Customer Satisfaction:</strong> <span class="bold">4.5/5.0</span></li>
          <li><strong>Employee Retention:</strong> <span class="bold">92%</span></li>
          <li><strong>Market Share:</strong> <span class="bold">23%</span></li>
        </ol>
        
        <div class="page-break"></div>
        
        <h2>Detailed Analysis</h2>
        <p>This section contains detailed analysis of the company's performance across various metrics and departments.</p>
        
        <h3>Department Performance Metrics</h3>
        <table>
          <thead>
            <tr>
              <th>Department</th>
              <th>Efficiency Score</th>
              <th>Quality Rating</th>
              <th>Innovation Index</th>
              <th>Overall Performance</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td class="bold">Engineering</td>
              <td class="text-center">95%</td>
              <td class="text-center">98%</td>
              <td class="text-center">92%</td>
              <td class="text-center"><span class="bold">95%</span></td>
            </tr>
            <tr>
              <td class="bold">Marketing</td>
              <td class="text-center">88%</td>
              <td class="text-center">85%</td>
              <td class="text-center">90%</td>
              <td class="text-center"><span class="bold">88%</span></td>
            </tr>
            <tr>
              <td class="bold">Sales</td>
              <td class="text-center">92%</td>
              <td class="text-center">89%</td>
              <td class="text-center">87%</td>
              <td class="text-center"><span class="bold">89%</span></td>
            </tr>
          </tbody>
        </table>
        
        <h3>Strategic Recommendations</h3>
        <ul>
          <li><strong>Budget Optimization:</strong> Increase marketing budget allocation for Q4</li>
          <li><strong>Process Improvement:</strong> Implement new efficiency measures in operations</li>
          <li><strong>Employee Development:</strong> Enhance employee training programs</li>
          <li><strong>Market Expansion:</strong> Expand market presence in emerging regions</li>
        </ul>
        
        <div class="footer">
          <p class="bold">Confidential - Internal Use Only</p>
          <p>Page 1 of 2 | Generated by Advanced FileFlow PDF to Word Converter</p>
          <p class="italic">This document tests advanced formatting preservation capabilities</p>
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
    console.log(`   ✅ Advanced test PDF created successfully!`);
    console.log(`   📁 Location: ${testPdfPath}`);
    console.log(`   📄 Size: ${(stats.size / 1024).toFixed(2)} KB`);
    
  } catch (error) {
    console.error('   ❌ Failed to create test PDF:', error.message);
    return;
  }
  
  // Test 3: Test advanced PDF to Word conversion
  console.log('\n3. Testing advanced PDF to Word conversion...');
  const outputDocxPath = path.join(tempDir, 'test-advanced-converted-document.docx');
  
  try {
    await AdvancedPdfToWordConverter.convertPdfToWord(testPdfPath, outputDocxPath);
    
    const outputStats = fs.statSync(outputDocxPath);
    console.log(`   ✅ Advanced conversion successful!`);
    console.log(`   📄 Output file size: ${(outputStats.size / 1024).toFixed(2)} KB`);
    console.log(`   📁 Saved to: ${outputDocxPath}`);
    
    if (outputStats.size < 1000) {
      console.log('   ⚠️  Output file seems very small, conversion might have issues.');
    }
    
  } catch (error) {
    console.log(`   ❌ Advanced conversion failed: ${error.message}`);
  }
  
  // Summary
  console.log('\n=== Advanced Test Summary ===');
  console.log('📁 Check the following files in the temp directory:');
  console.log('   - test-advanced-document.pdf (input PDF with complex formatting)');
  console.log('   - test-advanced-converted-document.docx (converted Word document)');
  
  console.log('\n💡 What to check in the converted document:');
  console.log('1. Headers and footers are preserved exactly');
  console.log('2. Tables maintain their structure, borders, and formatting');
  console.log('3. Font styles (bold, italic, underline) are preserved');
  console.log('4. Font sizes and families are maintained');
  console.log('5. Text alignment (center, right, left) is preserved');
  console.log('6. Page breaks and layout are maintained');
  console.log('7. Lists and bullet points are properly formatted');
  console.log('8. Colors and highlighting are preserved');
  console.log('9. Overall spacing and margins are similar to original');
  
  console.log('\n🔧 Advanced Conversion Methods Used:');
  console.log('1. LibreOffice with multiple format options (best for formatting)');
  console.log('2. pdf2docx with enhanced settings (good for tables)');
  console.log('3. Enhanced pdf-parse with advanced text processing');
  console.log('4. Basic text extraction (fallback)');
  
  console.log('\n✅ Advanced conversion test completed!');
}

// Run the test
testAdvancedPdfConversion().catch(console.error);


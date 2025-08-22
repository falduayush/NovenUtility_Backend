const fs = require('fs-extra');
const path = require('path');
const PdfConverter = require('./utils/pdfConverter');
const DocxProcessor = require('./utils/docxProcessor');

async function testPdfConversion() {
  console.log('=== PDF Conversion Test ===\n');
  
  // Test 1: Check LibreOffice availability
  console.log('1. Checking LibreOffice availability...');
  try {
    const libreOfficeAvailable = await PdfConverter.isLibreOfficeAvailable();
    console.log(`   LibreOffice available: ${libreOfficeAvailable ? 'YES' : 'NO'}`);
    if (!libreOfficeAvailable) {
      console.log('   ⚠️  LibreOffice not found. PDF conversion will use HTML fallback (may lose formatting).');
      console.log('   💡 Install LibreOffice for best PDF conversion results.');
    }
  } catch (error) {
    console.log(`   ❌ Error checking LibreOffice: ${error.message}`);
  }
  
  // Test 2: Create a test DOCX file
  console.log('\n2. Creating test DOCX file...');
  const testContent = `Test Document

This is a test document with variables.

Company Name: {{CompanyName}}
Contact Person: {{ContactPerson}}
Date: {{Date}}

This document contains:
- Headers and formatting
- Variables in {{VariableName}} format
- Multiple paragraphs
- Basic structure

Please replace the variables above with actual values.`;
  
  const tempDir = path.join(__dirname, 'temp');
  await fs.ensureDir(tempDir);
  
  const testDocxPath = path.join(tempDir, 'test-template.docx');
  await DocxProcessor.createDocxFromText(testContent, testDocxPath);
  console.log(`   ✅ Test DOCX created: ${testDocxPath}`);
  
  // Test 3: Test PDF conversion
  console.log('\n3. Testing PDF conversion...');
  const testPdfPath = path.join(tempDir, 'test-output.pdf');
  
  try {
    await PdfConverter.convertDocxToPdfNative(testDocxPath, testPdfPath);
    
    const pdfStats = fs.statSync(testPdfPath);
    console.log(`   ✅ PDF conversion successful!`);
    console.log(`   📄 PDF file size: ${pdfStats.size} bytes`);
    console.log(`   📁 PDF saved to: ${testPdfPath}`);
    
    if (pdfStats.size < 1000) {
      console.log('   ⚠️  PDF file seems very small, conversion might have issues.');
    }
  } catch (error) {
    console.log(`   ❌ PDF conversion failed: ${error.message}`);
    console.log('   💡 Check if LibreOffice is properly installed.');
  }
  
  // Test 4: Test with variables
  console.log('\n4. Testing template processing with variables...');
  const variables = {
    CompanyName: 'Test Company Inc.',
    ContactPerson: 'John Doe',
    Date: new Date().toLocaleDateString()
  };
  
  try {
    const processedDocxPath = path.join(tempDir, 'test-processed.docx');
    await DocxProcessor.processTemplate(testDocxPath, variables, processedDocxPath);
    console.log('   ✅ Template processing successful!');
    
    const processedPdfPath = path.join(tempDir, 'test-processed.pdf');
    await PdfConverter.convertDocxToPdfNative(processedDocxPath, processedPdfPath);
    
    const processedPdfStats = fs.statSync(processedPdfPath);
    console.log(`   ✅ Processed PDF conversion successful!`);
    console.log(`   📄 Processed PDF file size: ${processedPdfStats.size} bytes`);
    console.log(`   📁 Processed PDF saved to: ${processedPdfPath}`);
    
  } catch (error) {
    console.log(`   ❌ Template processing failed: ${error.message}`);
  }
  
  // Summary
  console.log('\n=== Test Summary ===');
  console.log('📁 Check the following files in the temp directory:');
  console.log('   - test-template.docx (original template)');
  console.log('   - test-output.pdf (converted PDF)');
  console.log('   - test-processed.docx (template with variables)');
  console.log('   - test-processed.pdf (processed PDF)');
  
  console.log('\n💡 Recommendations:');
  console.log('1. Open both DOCX and PDF files to compare formatting');
  console.log('2. If PDF formatting is poor, install LibreOffice');
  console.log('3. If conversion fails, check LibreOffice installation');
  console.log('4. Restart the server after installing LibreOffice');
  
  console.log('\n✅ Test completed!');
}

// Run the test
testPdfConversion().catch(console.error);

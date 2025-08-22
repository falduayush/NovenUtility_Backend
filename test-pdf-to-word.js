const fs = require('fs-extra');
const path = require('path');
const { v4: uuidv4 } = require('uuid');

async function testPdfToWordConversion() {
  console.log('=== PDF to Word Conversion Test ===\n');
  
  // Test 1: Check if we can create a test PDF
  console.log('1. Creating test PDF file...');
  const tempDir = path.join(__dirname, 'temp');
  await fs.ensureDir(tempDir);
  
  // Create a simple test PDF using a basic approach
  const testPdfPath = path.join(tempDir, 'test-document.pdf');
  
  try {
    // For testing, we'll create a simple PDF using a basic method
    // In a real scenario, you'd have an actual PDF file
    console.log('   ⚠️  Please place a test PDF file at:', testPdfPath);
    console.log('   💡 You can create a simple PDF using any PDF creator or download a sample PDF');
    
    // Check if test file exists
    if (fs.existsSync(testPdfPath)) {
      const stats = fs.statSync(testPdfPath);
      console.log(`   ✅ Test PDF found: ${(stats.size / 1024).toFixed(2)} KB`);
    } else {
      console.log('   ❌ Test PDF not found. Please create or copy a PDF file to test with.');
      console.log('   📁 Expected location:', testPdfPath);
      return;
    }
  } catch (error) {
    console.log(`   ❌ Error checking test PDF: ${error.message}`);
    return;
  }
  
  // Test 2: Test PDF upload endpoint
  console.log('\n2. Testing PDF upload endpoint...');
  try {
    const FormData = require('form-data');
    const axios = require('axios');
    
    const form = new FormData();
    form.append('file', fs.createReadStream(testPdfPath));
    
    const uploadResponse = await axios.post('http://localhost:3001/api/upload-pdf', form, {
      headers: {
        ...form.getHeaders(),
      },
      timeout: 30000
    });
    
    if (uploadResponse.data.success) {
      console.log('   ✅ PDF upload successful');
      console.log('   📄 File ID:', uploadResponse.data.fileId);
      console.log('   📄 File Name:', uploadResponse.data.fileName);
      
      const fileId = uploadResponse.data.fileId;
      
      // Test 3: Test PDF to Word conversion
      console.log('\n3. Testing PDF to Word conversion...');
      try {
        const convertResponse = await axios.post('http://localhost:3001/api/convert-pdf-to-word', {
          fileId: fileId
        }, {
          timeout: 120000 // 2 minutes timeout for conversion
        });
        
        if (convertResponse.data.success) {
          console.log('   ✅ PDF to Word conversion successful');
          console.log('   📄 Output file:', convertResponse.data.fileName);
          console.log('   📥 Download URL:', convertResponse.data.downloadUrl);
          
          // Test 4: Test download
          console.log('\n4. Testing file download...');
          try {
            const downloadUrl = `http://localhost:3001${convertResponse.data.downloadUrl}`;
            const downloadResponse = await axios.get(downloadUrl, {
              responseType: 'stream',
              timeout: 30000
            });
            
            const outputPath = path.join(tempDir, convertResponse.data.fileName);
            const writer = fs.createWriteStream(outputPath);
            downloadResponse.data.pipe(writer);
            
            await new Promise((resolve, reject) => {
              writer.on('finish', resolve);
              writer.on('error', reject);
            });
            
            const downloadStats = fs.statSync(outputPath);
            console.log('   ✅ File download successful');
            console.log('   📄 Downloaded file size:', (downloadStats.size / 1024).toFixed(2), 'KB');
            console.log('   📁 Saved to:', outputPath);
            
          } catch (downloadError) {
            console.log('   ❌ File download failed:', downloadError.message);
          }
          
        } else {
          console.log('   ❌ PDF to Word conversion failed:', convertResponse.data.error);
        }
        
      } catch (convertError) {
        console.log('   ❌ PDF to Word conversion failed:', convertError.message);
        if (convertError.response) {
          console.log('   📄 Error details:', convertError.response.data);
        }
      }
      
    } else {
      console.log('   ❌ PDF upload failed:', uploadResponse.data.error);
    }
    
  } catch (uploadError) {
    console.log('   ❌ PDF upload failed:', uploadError.message);
    if (uploadError.response) {
      console.log('   📄 Error details:', uploadError.response.data);
    }
  }
  
  // Summary
  console.log('\n=== Test Summary ===');
  console.log('📁 Check the following files in the temp directory:');
  console.log('   - test-document.pdf (input PDF)');
  console.log('   - converted-*.docx (output Word document)');
  
  console.log('\n💡 Recommendations:');
  console.log('1. Make sure the backend server is running on port 3001');
  console.log('2. Install LibreOffice for best conversion results');
  console.log('3. Test with different types of PDFs (text, images, tables)');
  console.log('4. Check the converted Word document for formatting accuracy');
  
  console.log('\n✅ Test completed!');
}

// Check if required dependencies are available
async function checkDependencies() {
  console.log('Checking dependencies...\n');
  
  const requiredDeps = ['axios', 'form-data'];
  const missingDeps = [];
  
  for (const dep of requiredDeps) {
    try {
      require(dep);
      console.log(`✅ ${dep} - available`);
    } catch (error) {
      missingDeps.push(dep);
      console.log(`❌ ${dep} - missing`);
    }
  }
  
  if (missingDeps.length > 0) {
    console.log(`\n⚠️  Missing dependencies: ${missingDeps.join(', ')}`);
    console.log('💡 Install them with: npm install ' + missingDeps.join(' '));
    return false;
  }
  
  return true;
}

// Run the test
async function main() {
  const depsOk = await checkDependencies();
  if (depsOk) {
    await testPdfToWordConversion();
  }
}

main().catch(console.error);


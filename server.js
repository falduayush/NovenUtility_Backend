const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const { v4: uuidv4 } = require('uuid');
const DocxProcessor = require('./utils/docxProcessor');
const PdfConverter = require('./utils/pdfConverter');
const PdfToWordConverter = require('./utils/pdfToWordConverter');
const AdvancedPdfToWordConverter = require('./utils/advancedPdfToWordConverter');

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Create uploads and temp directories
const uploadsDir = path.join(__dirname, 'uploads');
const tempDir = path.join(__dirname, 'temp');

fs.ensureDirSync(uploadsDir);
fs.ensureDirSync(tempDir);

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    // Save with UUID and original extension only, avoid keeping original name
    const ext = path.extname(file.originalname) || '.docx'
    const uniqueName = `${uuidv4()}${ext}`;
    cb(null, uniqueName);
  }
});

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      cb(null, true);
    } else {
      cb(new Error('Only .docx files are allowed'), false);
    }
  },
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

// Configure multer for PDF uploads
const pdfStorage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const uniqueName = `${uuidv4()}-${file.originalname}`;
    cb(null, uniqueName);
  }
});

const uploadPdf = multer({
  storage: pdfStorage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new Error('Only .pdf files are allowed'), false);
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024 // 50MB limit
  }
});

// Store templates in memory (in production, use a database)
const templates = new Map();

// Simple persistence to disk for templates and their saved values
const dataDir = path.join(__dirname, 'data');
const templatesDataFile = path.join(dataDir, 'templates.json');
fs.ensureDirSync(dataDir);

function saveTemplatesToDisk() {
  try {
    // Convert Map values to an array of plain objects
    const serializable = Array.from(templates.values());
    fs.writeFileSync(templatesDataFile, JSON.stringify(serializable, null, 2));
  } catch (err) {
    console.error('Failed to save templates to disk:', err);
  }
}

function loadTemplatesFromDisk() {
  try {
    if (fs.existsSync(templatesDataFile)) {
      const raw = fs.readFileSync(templatesDataFile, 'utf8');
      const arr = JSON.parse(raw);
      templates.clear();
      arr.forEach((t) => {
        // Backward compatible: ensure fields exist
        const normalized = {
          id: t.id,
          name: t.name || '',
          originalFile: t.originalFile,
          variables: Array.isArray(t.variables) ? t.variables : [],
          originalText: t.originalText || '',
          createdAt: t.createdAt || new Date().toISOString(),
          // new optional field for saved default values
          savedValues: t.savedValues && typeof t.savedValues === 'object' ? t.savedValues : {}
        };
        templates.set(normalized.id, normalized);
      });
      console.log(`Loaded ${templates.size} templates from disk`);
    }
  } catch (err) {
    console.error('Failed to load templates from disk:', err);
  }
}

// Store uploaded PDF files for conversion (in production, use a database)
const uploadedPdfs = new Map();



// Step 1: Upload and parse DOCX file
app.post('/api/upload-template', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;
    const templateId = uuidv4();
    const originalBaseName = path.parse(req.file.originalname).name;

    // Extract text and variables from DOCX
    const { text, variables: variablesArray } = await DocxProcessor.extractTextAndVariables(filePath);

    // Store template information
    templates.set(templateId, {
      id: templateId,
      name: originalBaseName,
      originalFile: filePath,
      variables: variablesArray,
      originalText: text,
      createdAt: new Date().toISOString(),
      // optional persisted defaults per variable
      savedValues: {}
    });

    // persist to disk
    saveTemplatesToDisk();

    res.json({
      success: true,
      templateId: templateId,
      variables: variablesArray,
      message: `Template uploaded successfully. Found ${variablesArray.length} variables.`
    });

  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({ error: 'Failed to process template' });
  }
});

// Step 2: Get template variables
app.get('/api/template/:templateId', (req, res) => {
  try {
    const { templateId } = req.params;
    const template = templates.get(templateId);

    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    res.json({
      success: true,
      template: {
        id: template.id,
        name: template.name || '',
        variables: template.variables,
        createdAt: template.createdAt,
        savedValues: template.savedValues || {}
      }
    });

  } catch (error) {
    console.error('Get template error:', error);
    res.status(500).json({ error: 'Failed to get template' });
  }
});

// Save default values for a template's variables
app.post('/api/template/:templateId/values', (req, res) => {
  try {
    const { templateId } = req.params;
    const { values } = req.body;

    if (!values || typeof values !== 'object') {
      return res.status(400).json({ error: 'Invalid or missing values object' });
    }

    const template = templates.get(templateId);
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    template.savedValues = {
      ...(template.savedValues || {}),
      ...values
    };
    templates.set(templateId, template);
    saveTemplatesToDisk();

    res.json({
      success: true,
      savedValues: template.savedValues
    });

  } catch (error) {
    console.error('Save template values error:', error);
    res.status(500).json({ error: 'Failed to save template values' });
  }
});

// Step 3 & 4: Generate document with user values
app.post('/api/generate-document', async (req, res) => {
  try {
    const { templateId, variables, format = 'docx' } = req.body;

    if (!templateId || !variables) {
      return res.status(400).json({ error: 'Template ID and variables are required' });
    }

    const template = templates.get(templateId);
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    // Process template with variables
    let outputPath;
    let fileName;

    if (format === 'docx') {
      fileName = `generated-${templateId}.docx`;
      outputPath = path.join(tempDir, fileName);
      console.log('Generating DOCX file at:', outputPath);
      await DocxProcessor.processTemplate(template.originalFile, variables, outputPath);
      
      // Verify file was created
      if (!fs.existsSync(outputPath)) {
        throw new Error('Generated file was not created');
      }
      
      const stats = fs.statSync(outputPath);
      console.log('Generated file size:', stats.size);
      
      if (stats.size === 0) {
        throw new Error('Generated file is empty');
      }
    } else if (format === 'pdf') {
      // Render the DOCX with variables preserved (keeps tables/styles)
      const intermediateDocx = path.join(tempDir, `intermediate-${templateId}.docx`);
      console.log('Generating intermediate DOCX at:', intermediateDocx);
      await DocxProcessor.processTemplate(template.originalFile, variables, intermediateDocx);

      fileName = `generated-${templateId}.pdf`;
      outputPath = path.join(tempDir, fileName);

      // Check if LibreOffice is available for best PDF conversion
      const libreOfficeAvailable = await PdfConverter.isLibreOfficeAvailable();
      console.log('LibreOffice available:', libreOfficeAvailable);

      // Use improved PDF conversion with multiple fallback options
      console.log('Starting PDF conversion with multiple fallback options:', outputPath);
      try {
        await PdfConverter.convertDocxToPdfNative(intermediateDocx, outputPath);
        console.log('PDF conversion completed successfully');
      } catch (pdfErr) {
        console.error('All PDF conversion methods failed:', pdfErr.message);
        throw new Error(`PDF conversion failed: ${pdfErr.message}. Please ensure LibreOffice is installed for best results.`);
      }
      
      // Verify file was created
      if (!fs.existsSync(outputPath)) {
        throw new Error('Generated file was not created');
      }
      
      const stats = fs.statSync(outputPath);
      console.log('Generated file size:', stats.size);
      
      if (stats.size === 0) {
        throw new Error('Generated file is empty');
      }
    } else {
      return res.status(400).json({ error: 'Unsupported format. Use "docx" or "pdf"' });
    }

    // Create download URL
    const downloadUrl = `/api/download/${fileName}`;
    console.log('Download URL created:', downloadUrl);

    res.json({
      success: true,
      downloadUrl: downloadUrl,
      fileName: fileName,
      message: `Document generated successfully in ${format.toUpperCase()} format.`
    });

  } catch (error) {
    console.error('Generate document error:', error);
    res.status(500).json({ error: 'Failed to generate document' });
  }
});

// Step 5: Download generated file
app.get('/api/download/:fileName', (req, res) => {
  try {
    const { fileName } = req.params;
    const filePath = path.join(tempDir, fileName);
    
    console.log('Download request for file:', fileName);
    console.log('File path:', filePath);

    if (!fs.existsSync(filePath)) {
      console.error('File not found:', filePath);
      return res.status(404).json({ error: 'File not found' });
    }

    // Get file stats
    const stats = fs.statSync(filePath);
    console.log('File exists, size:', stats.size);
    
    if (stats.size === 0) {
      console.error('File is empty:', filePath);
      return res.status(500).json({ error: 'File is empty' });
    }

    // Set proper headers for file download
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    res.setHeader('Content-Length', stats.size);
    
    // Determine content type based on file extension
    const ext = path.extname(fileName).toLowerCase();
    if (ext === '.pdf') {
      res.setHeader('Content-Type', 'application/pdf');
    } else if (ext === '.docx') {
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    } else {
      res.setHeader('Content-Type', 'application/octet-stream');
    }
    
    console.log('Content-Type set to:', res.getHeader('Content-Type'));

    // Stream the file
    const fileStream = fs.createReadStream(filePath);
    fileStream.pipe(res);
    
    // Handle stream errors
    fileStream.on('error', (err) => {
      console.error('File stream error:', err);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Failed to stream file' });
      }
    });
    
    // Log when download starts
    fileStream.on('open', () => {
      console.log('File stream opened successfully');
    });
    
    // Log when download ends
    res.on('finish', () => {
      console.log('Download completed for:', fileName);
      // Don't clean up immediately - let files stay for debugging
      // setTimeout(() => {
      //   fs.remove(filePath).catch(console.error);
      // }, 10000); // Increased delay to ensure download completes
    });

  } catch (error) {
    console.error('Download error:', error);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Failed to download file' });
    }
  }
});

// List all templates
app.get('/api/templates', (req, res) => {
  try {
    const templatesList = Array.from(templates.values()).map(template => ({
      id: template.id,
      name: template.name || '',
      variables: template.variables,
      createdAt: template.createdAt
    }));

    res.json({
      success: true,
      templates: templatesList
    });

  } catch (error) {
    console.error('List templates error:', error);
    res.status(500).json({ error: 'Failed to list templates' });
  }
});

// Delete template
app.delete('/api/template/:templateId', async (req, res) => {
  try {
    const { templateId } = req.params;
    const template = templates.get(templateId);

    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    // Remove original file
    await fs.remove(template.originalFile);
    
    // Remove from memory
    templates.delete(templateId);

    // persist to disk
    saveTemplatesToDisk();

    res.json({
      success: true,
      message: 'Template deleted successfully'
    });

  } catch (error) {
    console.error('Delete template error:', error);
    res.status(500).json({ error: 'Failed to delete template' });
  }
});

// Get template statistics
app.get('/api/template/:templateId/stats', async (req, res) => {
  try {
    const { templateId } = req.params;
    const template = templates.get(templateId);

    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    const stats = await DocxProcessor.getFileStats(template.originalFile);

    res.json({
      success: true,
      stats: {
        ...stats,
        templateId: template.id,
        createdAt: template.createdAt
      }
    });

  } catch (error) {
    console.error('Get template stats error:', error);
    res.status(500).json({ error: 'Failed to get template statistics' });
  }
});

// Preview template content
app.get('/api/template/:templateId/preview', async (req, res) => {
  try {
    const { templateId } = req.params;
    const template = templates.get(templateId);

    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    const { text } = await DocxProcessor.extractTextAndVariables(template.originalFile);

    res.json({
      success: true,
      preview: text.substring(0, 500) + (text.length > 500 ? '...' : ''),
      fullContent: text
    });

  } catch (error) {
    console.error('Preview template error:', error);
    res.status(500).json({ error: 'Failed to preview template' });
  }
});

// Test download endpoint
app.get('/api/test-download', (req, res) => {
  try {
    const testContent = 'This is a test file for download functionality.';
    const testFileName = 'test-file.txt';
    const testFilePath = path.join(tempDir, testFileName);
    
    // Create a test file
    fs.writeFileSync(testFilePath, testContent);
    
    // Set headers
    res.setHeader('Content-Disposition', `attachment; filename="${testFileName}"`);
    res.setHeader('Content-Type', 'text/plain');
    res.setHeader('Content-Length', testContent.length);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    
    // Send the file
    res.sendFile(testFilePath, (err) => {
      if (err) {
        console.error('Test download error:', err);
      }
      // Clean up test file
      setTimeout(() => {
        fs.remove(testFilePath).catch(console.error);
      }, 5000);
    });
    
  } catch (error) {
    console.error('Test download error:', error);
    res.status(500).json({ error: 'Failed to create test download' });
  }
});

// Test DOCX generation endpoint
app.get('/api/test-docx', async (req, res) => {
  try {
    const testContent = 'This is a test DOCX file.\n\nIt contains multiple lines.\n\nGenerated for testing purposes.';
    const testFileName = 'test-document.docx';
    const testFilePath = path.join(tempDir, testFileName);
    
    // Create a test DOCX file using the DocxProcessor
    await DocxProcessor.createDocxFromText(testContent, testFilePath);
    
    // Set headers
    res.setHeader('Content-Disposition', `attachment; filename="${testFileName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    
    // Get file stats
    const stats = fs.statSync(testFilePath);
    res.setHeader('Content-Length', stats.size);
    
    // Send the file
    res.sendFile(testFilePath, (err) => {
      if (err) {
        console.error('Test DOCX download error:', err);
      }
      // Clean up test file
      setTimeout(() => {
        fs.remove(testFilePath).catch(console.error);
      }, 5000);
    });
    
  } catch (error) {
    console.error('Test DOCX error:', error);
    res.status(500).json({ error: 'Failed to create test DOCX' });
  }
});

// Health check endpoint
app.get('/api/health', async (req, res) => {
  try {
    const libreOfficeAvailable = await PdfConverter.isLibreOfficeAvailable();
    res.json({
      status: 'OK',
      timestamp: new Date().toISOString(),
      templatesCount: templates.size,
      libreOfficeAvailable: libreOfficeAvailable,
      pdfConversion: libreOfficeAvailable ? 'High fidelity (LibreOffice)' : 'Limited fidelity (HTML fallback)',
      recommendations: libreOfficeAvailable ? [] : [
        'Install LibreOffice for best PDF conversion results',
        'PDF output may lose formatting without LibreOffice'
      ]
    });
  } catch (error) {
    res.json({
      status: 'OK',
      timestamp: new Date().toISOString(),
      templatesCount: templates.size,
      libreOfficeAvailable: false,
      pdfConversion: 'Limited fidelity (HTML fallback)',
      recommendations: [
        'Install LibreOffice for best PDF conversion results',
        'PDF output may lose formatting without LibreOffice'
      ],
      error: error.message
    });
  }
});

// Debug endpoint to list temp files
app.get('/api/debug/files', (req, res) => {
  try {
    const files = fs.readdirSync(tempDir);
    const fileStats = files.map(file => {
      const filePath = path.join(tempDir, file);
      const stats = fs.statSync(filePath);
      return {
        name: file,
        size: stats.size,
        created: stats.birthtime,
        modified: stats.mtime
      };
    });
    
    res.json({
      success: true,
      files: fileStats,
      tempDir: tempDir
    });
  } catch (error) {
    console.error('Debug files error:', error);
    res.status(500).json({ error: 'Failed to list files' });
  }
});

// PDF to Word Conversion Endpoints

// Upload PDF file for conversion
app.post('/api/upload-pdf', uploadPdf.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;
    const fileId = uuidv4();
    const originalName = req.file.originalname;

    // Store file information
    uploadedPdfs.set(fileId, {
      id: fileId,
      originalFile: filePath,
      originalName: originalName,
      createdAt: new Date().toISOString()
    });

    res.json({
      success: true,
      fileId: fileId,
      fileName: originalName,
      message: 'PDF uploaded successfully for conversion.'
    });

  } catch (error) {
    console.error('PDF upload error:', error);
    res.status(500).json({ error: 'Failed to upload PDF' });
  }
});

// Convert PDF to Word
app.post('/api/convert-pdf-to-word', async (req, res) => {
  try {
    const { fileId } = req.body;

    if (!fileId) {
      return res.status(400).json({ error: 'File ID is required' });
    }

    const pdfFile = uploadedPdfs.get(fileId);
    if (!pdfFile) {
      return res.status(404).json({ error: 'PDF file not found' });
    }

    // Generate output filename
    const outputFileName = `converted-${fileId}.docx`;
    const outputPath = path.join(tempDir, outputFileName);

    console.log('Converting PDF to Word:', pdfFile.originalFile, '->', outputPath);

    // Use the advanced PDF to Word converter for maximum formatting preservation
    await AdvancedPdfToWordConverter.convertPdfToWord(pdfFile.originalFile, outputPath);

    // Verify file was created
    if (!fs.existsSync(outputPath)) {
      throw new Error('Converted file was not created');
    }
    
    const stats = fs.statSync(outputPath);
    console.log('Converted file size:', stats.size);
    
    if (stats.size === 0) {
      throw new Error('Converted file is empty');
    }

    // Create download URL
    const downloadUrl = `/api/download/${outputFileName}`;
    console.log('Download URL created:', downloadUrl);

    res.json({
      success: true,
      downloadUrl: downloadUrl,
      fileName: outputFileName,
      message: 'PDF converted to Word successfully.'
    });

  } catch (error) {
    console.error('PDF to Word conversion error:', error);
    res.status(500).json({ error: `Failed to convert PDF to Word: ${error.message}` });
  }
});

// Get uploaded PDF file info
app.get('/api/pdf/:fileId', (req, res) => {
  try {
    const { fileId } = req.params;
    const pdfFile = uploadedPdfs.get(fileId);

    if (!pdfFile) {
      return res.status(404).json({ error: 'PDF file not found' });
    }

    res.json({
      success: true,
      file: {
        id: pdfFile.id,
        originalName: pdfFile.originalName,
        createdAt: pdfFile.createdAt
      }
    });

  } catch (error) {
    console.error('Get PDF file error:', error);
    res.status(500).json({ error: 'Failed to get PDF file' });
  }
});

// List uploaded PDF files
app.get('/api/pdfs', (req, res) => {
  try {
    const pdfsList = Array.from(uploadedPdfs.values()).map(pdf => ({
      id: pdf.id,
      originalName: pdf.originalName,
      createdAt: pdf.createdAt
    }));

    res.json({
      success: true,
      pdfs: pdfsList
    });

  } catch (error) {
    console.error('List PDFs error:', error);
    res.status(500).json({ error: 'Failed to list PDF files' });
  }
});

// Delete uploaded PDF file
app.delete('/api/pdf/:fileId', async (req, res) => {
  try {
    const { fileId } = req.params;
    const pdfFile = uploadedPdfs.get(fileId);

    if (!pdfFile) {
      return res.status(404).json({ error: 'PDF file not found' });
    }

    // Remove original file
    await fs.remove(pdfFile.originalFile);
    
    // Remove from memory
    uploadedPdfs.delete(fileId);

    res.json({
      success: true,
      message: 'PDF file deleted successfully'
    });

  } catch (error) {
    console.error('Delete PDF file error:', error);
    res.status(500).json({ error: 'Failed to delete PDF file' });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Server error:', error);
  res.status(500).json({ error: 'Internal server error' });
});

// Start server
loadTemplatesFromDisk();

app.listen(PORT, () => {
  console.log(`Template Editor Backend running on port ${PORT}`);
  console.log(`Health check: http://localhost:${PORT}/api/health`);
});

# FileFlow Template Editor Backend

A Node.js backend for the FileFlow template editor that processes DOCX templates and generates documents in DOCX and PDF formats.

## Features

- Upload and parse DOCX templates with variables
- Generate documents with user-provided variable values
- Support for both DOCX and PDF output formats
- Template management (save, load, delete)
- High-fidelity PDF conversion with LibreOffice

## Installation

### Prerequisites

1. **Node.js** (v14 or higher)
2. **LibreOffice** (for best PDF conversion results)

### Installing LibreOffice

#### Windows
1. Download LibreOffice from [https://www.libreoffice.org/download/download/](https://www.libreoffice.org/download/download/)
2. Install with default settings
3. The backend will automatically detect LibreOffice in common installation paths

#### macOS
```bash
# Using Homebrew
brew install --cask libreoffice

# Or download from the official website
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt update
sudo apt install libreoffice
```

### Backend Setup

1. Install dependencies:
```bash
cd backend
npm install
```

2. Start the server:
```bash
npm start
```

The server will run on `http://localhost:3001`

## API Endpoints

### Health Check
- `GET /api/health` - Check server status and LibreOffice availability

### Template Management
- `POST /api/upload-template` - Upload a DOCX template
- `GET /api/templates` - List all saved templates
- `GET /api/template/:id` - Get template details
- `DELETE /api/template/:id` - Delete a template

### Document Generation
- `POST /api/generate-document` - Generate document with variables
- `GET /api/download/:filename` - Download generated file

## PDF Conversion

The backend uses multiple methods for PDF conversion, in order of preference:

1. **LibreOffice Command Line** (Best fidelity - preserves formatting, tables, headers, footers)
2. **LibreOffice Convert Library** (Good fidelity)
3. **HTML Conversion** (Fallback - may lose some formatting)

### Troubleshooting PDF Issues

#### PDF doesn't match DOCX formatting

**Problem**: Generated PDF looks different from the original DOCX template.

**Solutions**:
1. **Install LibreOffice**: This is the most important step for high-fidelity PDF conversion
2. **Check LibreOffice installation**: Use the health check endpoint to verify LibreOffice is detected
3. **Restart the server**: After installing LibreOffice, restart the backend server

#### LibreOffice not detected

**Problem**: The health check shows LibreOffice is not available.

**Solutions**:
1. **Verify installation**: Make sure LibreOffice is properly installed
2. **Check PATH**: Ensure LibreOffice is in your system PATH
3. **Windows users**: LibreOffice is typically installed in:
   - `C:\Program Files\LibreOffice\program\soffice.exe`
   - `C:\Program Files (x86)\LibreOffice\program\soffice.exe`

#### PDF conversion fails

**Problem**: All PDF conversion methods fail.

**Solutions**:
1. **Check LibreOffice**: Ensure LibreOffice is installed and accessible
2. **Check file permissions**: Ensure the temp directory is writable
3. **Check disk space**: Ensure there's enough disk space for temporary files
4. **Review logs**: Check the server console for detailed error messages

### Testing PDF Conversion

You can test if PDF conversion is working properly:

1. **Health Check**: Visit `http://localhost:3001/api/health` to see if LibreOffice is detected
2. **Test Generation**: Upload a simple template and try generating a PDF
3. **Compare Outputs**: Generate both DOCX and PDF versions and compare them

### Testing PDF to Word Conversion

You can test the PDF to Word conversion functionality:

1. **Create Test PDF**: Run `npm run create-test-pdf` to generate a sample PDF
2. **Test Conversion**: Run `npm run test-pdf-to-word` to test the full conversion process
3. **Manual Testing**: Use the frontend tool at `/tools/pdf-to-word` to upload and convert PDFs

## Development

### Running in Development Mode
```bash
npm run dev
```

### Testing
```bash
npm test
```

### Debug Endpoints

- `GET /api/debug/files` - List temporary files
- `GET /api/test-download` - Test file download functionality
- `GET /api/test-docx` - Test DOCX generation

### PDF to Word Conversion Endpoints

- `POST /api/upload-pdf` - Upload a PDF file for conversion
- `POST /api/convert-pdf-to-word` - Convert uploaded PDF to Word document
- `GET /api/pdf/:fileId` - Get uploaded PDF file information
- `GET /api/pdfs` - List all uploaded PDF files
- `DELETE /api/pdf/:fileId` - Delete uploaded PDF file

## File Structure

```
backend/
├── server.js              # Main server file
├── utils/
│   ├── docxProcessor.js   # DOCX processing utilities
│   └── pdfConverter.js    # PDF conversion utilities
├── uploads/               # Uploaded template files
├── temp/                  # Temporary generated files
└── package.json
```

## Environment Variables

- `PORT` - Server port (default: 3001)
- `SOFFICE_BIN` - Path to LibreOffice executable (optional)
- `LIBREOFFICE_BIN` - Alternative path to LibreOffice executable (optional)

## Troubleshooting

### Common Issues

1. **"LibreOffice not found"**: Install LibreOffice and restart the server
2. **"PDF conversion failed"**: Check if LibreOffice is properly installed
3. **"File not found"**: Check if the temp directory exists and is writable
4. **"Permission denied"**: Check file permissions on uploads and temp directories

### Logs

The server provides detailed logging for debugging:
- Template upload and processing
- PDF conversion attempts and results
- File operations and errors

Check the console output for detailed error messages when issues occur.




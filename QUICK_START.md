# Quick Start Guide - Template Editor Backend

## 🚀 Get Started in 3 Steps

### 1. Install Dependencies
```bash
cd backend
npm install
```

### 2. Start the Server
```bash
# Development mode (with auto-restart)
npm run dev

# Production mode
npm start
```

The server will start on `http://localhost:3001`

### 3. Test the API
```bash
npm test
```

## 📋 API Workflow

### Step 1: Upload Template
```bash
curl -X POST http://localhost:3001/api/upload-template \
  -F "file=@your-template.docx"
```

**Response:**
```json
{
  "success": true,
  "templateId": "uuid-here",
  "variables": ["Name", "Company", "Date"],
  "message": "Template uploaded successfully. Found 3 variables."
}
```

### Step 2: Generate Document
```bash
curl -X POST http://localhost:3001/api/generate-document \
  -H "Content-Type: application/json" \
  -d '{
    "templateId": "uuid-here",
    "variables": {
      "Name": "John Doe",
      "Company": "TechCorp",
      "Date": "2024-01-15"
    },
    "format": "docx"
  }'
```

**Response:**
```json
{
  "success": true,
  "downloadUrl": "/api/download/generated-uuid.docx",
  "fileName": "generated-uuid.docx",
  "message": "Document generated successfully in DOCX format."
}
```

### Step 3: Download File
```bash
curl -O http://localhost:3001/api/download/generated-uuid.docx
```

## 🔧 Template Format

Create a DOCX file with variables in this format:
```
Dear {{ RecipientName }},

Your meeting is scheduled for {{ MeetingDate }} at {{ MeetingLocation }}.

Best regards,
{{ SenderName }}
```

## 📁 File Structure
```
backend/
├── server.js              # Main server
├── utils/
│   ├── docxProcessor.js   # DOCX processing
│   └── pdfConverter.js    # PDF conversion
├── uploads/               # Uploaded files
├── temp/                  # Generated files
└── sample-template.txt    # Example template
```

## 🎯 Key Features

- ✅ DOCX file upload and parsing
- ✅ Automatic variable detection (`{{ VariableName }}`)
- ✅ Variable replacement with user values
- ✅ DOCX and PDF output formats
- ✅ Secure file download
- ✅ Template management
- ✅ File statistics and preview
- ✅ Automatic cleanup

## 🔍 Health Check
```bash
curl http://localhost:3001/api/health
```

## 📝 Sample Template

Use the content in `sample-template.txt` to create a DOCX file for testing.

## 🛠️ Development

- **Auto-restart**: `npm run dev`
- **Manual restart**: `npm start`
- **Test API**: `npm test`
- **Port**: 3001 (configurable via PORT env var)

## 🚨 Troubleshooting

1. **Port already in use**: Change PORT environment variable
2. **File upload fails**: Ensure file is .docx format and < 10MB
3. **PDF generation slow**: First run may take longer due to Puppeteer setup
4. **Template not found**: Check templateId is correct

## 📞 Support

Check the full README.md for detailed API documentation and examples.





const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);
const DocxProcessor = require('./docxProcessor');

class PdfToWordConverter {
  /**
   * Convert PDF to Word with enhanced formatting preservation
   * @param {string} inputPdfPath - Path to input PDF file
   * @param {string} outputDocxPath - Path for output DOCX file
   * @returns {Promise<string>} - Path to generated DOCX file
   */
  static async convertPdfToWord(inputPdfPath, outputDocxPath) {
    console.log('Starting enhanced PDF to Word conversion...');
    
    // Method 1: Try LibreOffice with enhanced options (best for headers, footers, formatting)
    try {
      console.log('Attempting LibreOffice conversion with enhanced formatting...');
      return await this.convertWithLibreOffice(inputPdfPath, outputDocxPath);
    } catch (error) {
      console.log('LibreOffice conversion failed:', error.message);
    }
    
    // Method 2: Try pdf2docx with enhanced formatting preservation
    try {
      console.log('Attempting pdf2docx conversion with enhanced formatting...');
      return await this.convertWithPdf2Docx(inputPdfPath, outputDocxPath);
    } catch (error) {
      console.log('pdf2docx conversion failed:', error.message);
    }
    
    // Method 3: Try alternative parsing methods
    try {
      console.log('Attempting alternative parsing methods...');
      return await this.convertWithAlternativeParsing(inputPdfPath, outputDocxPath);
    } catch (error) {
      console.log('Alternative parsing failed:', error.message);
    }
    
    // Method 4: Enhanced text extraction with structure preservation (fallback)
    console.log('Falling back to enhanced text extraction...');
    return await this.convertWithEnhancedTextExtraction(inputPdfPath, outputDocxPath);
  }

  /**
   * Convert using LibreOffice with enhanced options
   */
  static async convertWithLibreOffice(inputPdfPath, outputDocxPath) {
    const outputDir = path.dirname(outputDocxPath);
    
    // Try multiple LibreOffice conversion methods for best results
    const conversionMethods = [
      // Method 1: MS Word 2007 XML with best compatibility
      `soffice --headless --convert-to docx:"MS Word 2007 XML" --outdir "${outputDir}" "${inputPdfPath}"`,
      // Method 2: Office Open XML Text with better formatting
      `soffice --headless --convert-to docx:"Office Open XML Text" --outdir "${outputDir}" "${inputPdfPath}"`,
      // Method 3: Standard DOCX format
      `soffice --headless --convert-to docx --outdir "${outputDir}" "${inputPdfPath}"`
    ];
    
    for (let i = 0; i < conversionMethods.length; i++) {
      try {
        const command = conversionMethods[i];
        console.log(`Attempting LibreOffice conversion method ${i + 1}:`, command);
        
        const { stdout, stderr } = await execAsync(command, { timeout: 120000 });
        
        if (stderr) {
          console.log('LibreOffice stderr:', stderr);
        }
        
        console.log('LibreOffice stdout:', stdout);
        
        // LibreOffice creates the DOCX with the same name but .docx extension
        const expectedDocxPath = inputPdfPath.replace('.pdf', '.docx');
        
        if (fs.existsSync(expectedDocxPath)) {
          // Move to the desired output path
          await fs.move(expectedDocxPath, outputDocxPath, { overwrite: true });
          console.log(`LibreOffice conversion method ${i + 1} successful`);
          return outputDocxPath;
        }
      } catch (error) {
        console.log(`LibreOffice conversion method ${i + 1} failed:`, error.message);
        if (i === conversionMethods.length - 1) {
          throw error; // Re-throw if all methods failed
        }
      }
    }
    
    throw new Error('All LibreOffice conversion methods failed');
  }

  /**
   * Convert using pdf2docx with enhanced formatting preservation
   */
  static async convertWithPdf2Docx(inputPdfPath, outputDocxPath) {
    const { Converter } = require('pdf2docx');
    const converter = new Converter(inputPdfPath);
    
    // Enhanced conversion options for better formatting preservation
    const options = {
      start: 0, // Start from first page
      end: undefined, // Convert all pages
      pages: undefined, // Convert all pages
      zoom: 1.0, // No zoom
      debug: false,
      // Enhanced table settings
      table_settings: {
        vertical_strategy: 'text',
        horizontal_strategy: 'text',
        min_section_height: 20,
        connected_border_tolerance: 3,
        margin: 0.1
      },
      // Enhanced text settings
      text_settings: {
        font_size_mapping: true, // Preserve font sizes
        font_family_mapping: true, // Preserve font families
        line_spacing: true, // Preserve line spacing
        text_align: true, // Preserve text alignment
        text_color: true, // Preserve text colors
        text_bold: true, // Preserve bold text
        text_italic: true, // Preserve italic text
        text_underline: true // Preserve underlined text
      },
      // Enhanced layout settings
      layout_settings: {
        page_margin: true, // Preserve page margins
        page_size: true, // Preserve page size
        header_footer: true, // Preserve headers and footers
        page_break: true, // Preserve page breaks
        section_break: true // Preserve section breaks
      }
    };
    
    await converter.convert(outputDocxPath, options);
    console.log('pdf2docx conversion successful with enhanced settings');
    return outputDocxPath;
  }

  /**
   * Convert using enhanced text extraction with structure preservation
   */
  static async convertWithEnhancedTextExtraction(inputPdfPath, outputDocxPath) {
    const pdf = require('pdf-parse');
    const dataBuffer = await fs.readFile(inputPdfPath);
    const data = await pdf(dataBuffer);
    
    // Enhanced text processing to preserve structure
    const enhancedText = this.enhanceExtractedText(data.text);
    
    // Create a DOCX with enhanced formatting
    await DocxProcessor.createDocxFromText(enhancedText, outputDocxPath);
    console.log('Enhanced text extraction conversion successful');
    return outputDocxPath;
  }

  /**
   * Try alternative PDF parsing for better formatting preservation
   */
  static async convertWithAlternativeParsing(inputPdfPath, outputDocxPath) {
    try {
      // Try using pdf2pic for better text extraction with positioning
      const { fromPath } = require('pdf2pic');
      const options = {
        density: 300,
        saveFilename: "page",
        savePath: path.dirname(outputDocxPath),
        format: "png",
        width: 2480,
        height: 3508
      };
      
      const convert = fromPath(inputPdfPath, options);
      const pageData = await convert(1); // Convert first page
      
      // For now, fall back to enhanced text extraction
      return await this.convertWithEnhancedTextExtraction(inputPdfPath, outputDocxPath);
    } catch (error) {
      console.log('Alternative parsing failed, using enhanced text extraction:', error.message);
      return await this.convertWithEnhancedTextExtraction(inputPdfPath, outputDocxPath);
    }
  }

  /**
   * Enhance extracted text to better preserve formatting and structure
   */
  static enhanceExtractedText(text) {
    if (!text) return '';
    
    let enhancedText = text;
    
    // Preserve line breaks and paragraph structure
    enhancedText = enhancedText
      .replace(/\n\s*\n/g, '\n\n') // Normalize multiple line breaks
      .replace(/\r\n/g, '\n') // Normalize line endings
      .trim();
    
    // Split into lines for processing
    const lines = enhancedText.split('\n');
    const processedLines = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      // Skip empty lines
      if (!line) {
        processedLines.push('');
        continue;
      }
      
      // Try to identify and preserve table structures
      if (this.isTableRow(line)) {
        processedLines.push(this.formatTableRow(line));
      }
      // Try to identify headers
      else if (this.isHeader(line)) {
        processedLines.push(`# ${line}`);
      }
      // Try to identify numbered lists
      else if (this.isNumberedList(line)) {
        processedLines.push(`## ${line}`);
      }
      // Try to identify bullet points
      else if (this.isBulletPoint(line)) {
        processedLines.push(`### ${line}`);
      }
      // Regular text line
      else {
        processedLines.push(line);
      }
    }
    
    enhancedText = processedLines.join('\n');
    
    return enhancedText;
  }

  /**
   * Check if a line might be a table row
   */
  static isTableRow(line) {
    // Check for multiple spaces or tabs that might indicate table structure
    const spaceCount = (line.match(/\s{2,}/g) || []).length;
    const tabCount = (line.match(/\t/g) || []).length;
    
    // If there are multiple spaces or tabs, it might be a table
    return spaceCount > 0 || tabCount > 0;
  }

  /**
   * Format a table row for better preservation
   */
  static formatTableRow(line) {
    // Replace multiple spaces with tabs for better table structure
    return line.replace(/\s{2,}/g, '\t');
  }

  /**
   * Check if a line might be a header
   */
  static isHeader(line) {
    // Check if line is all caps and not too long
    return line === line.toUpperCase() && line.length > 3 && line.length < 100;
  }

  /**
   * Check if a line is a numbered list item
   */
  static isNumberedList(line) {
    return /^\d+\.\s/.test(line);
  }

  /**
   * Check if a line is a bullet point
   */
  static isBulletPoint(line) {
    return /^[•\-\*]\s/.test(line);
  }

  /**
   * Check if LibreOffice is available
   */
  static async isLibreOfficeAvailable() {
    try {
      const { stdout } = await execAsync('soffice --version', { timeout: 5000 });
      console.log('LibreOffice found:', stdout.trim());
      return true;
    } catch (error) {
      console.log('LibreOffice not found via command line');
      
      // Check common installation paths on Windows
      if (process.platform === 'win32') {
        const commonPaths = [
          'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
          'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
        ];
        
        for (const path of commonPaths) {
          if (fs.existsSync(path)) {
            console.log('LibreOffice found at:', path);
            return true;
          }
        }
      }
      
      return false;
    }
  }
}

module.exports = PdfToWordConverter;

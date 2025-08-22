const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);
const DocxProcessor = require('./docxProcessor');

class AdvancedPdfToWordConverter {
  /**
   * Convert PDF to Word with maximum formatting preservation
   * @param {string} inputPdfPath - Path to input PDF file
   * @param {string} outputDocxPath - Path for output DOCX file
   * @returns {Promise<string>} - Path to generated DOCX file
   */
  static async convertPdfToWord(inputPdfPath, outputDocxPath) {
    console.log('Starting advanced PDF to Word conversion...');
    
    // Method 1: Try LibreOffice with multiple conversion formats
    try {
      console.log('Attempting LibreOffice conversion with multiple formats...');
      return await this.convertWithLibreOfficeAdvanced(inputPdfPath, outputDocxPath);
    } catch (error) {
      console.log('LibreOffice conversion failed:', error.message);
    }
    
    // Method 2: Try pdf2docx with enhanced settings
    try {
      console.log('Attempting pdf2docx with enhanced settings...');
      return await this.convertWithPdf2DocxAdvanced(inputPdfPath, outputDocxPath);
    } catch (error) {
      console.log('pdf2docx conversion failed:', error.message);
    }
    
    // Method 3: Try using pdf-parse with enhanced text processing
    try {
      console.log('Attempting enhanced pdf-parse conversion...');
      return await this.convertWithPdfParseAdvanced(inputPdfPath, outputDocxPath);
    } catch (error) {
      console.log('Enhanced pdf-parse conversion failed:', error.message);
    }
    
    // Method 4: Fallback to basic text extraction
    console.log('Falling back to basic text extraction...');
    return await this.convertWithBasicExtraction(inputPdfPath, outputDocxPath);
  }

  /**
   * Convert using LibreOffice with multiple advanced methods
   */
  static async convertWithLibreOfficeAdvanced(inputPdfPath, outputDocxPath) {
    const outputDir = path.dirname(outputDocxPath);
    
    // Try multiple LibreOffice conversion methods in order of preference
    const conversionMethods = [
      // Method 1: Office Open XML Text (best for formatting)
      `soffice --headless --convert-to docx:"Office Open XML Text" --outdir "${outputDir}" "${inputPdfPath}"`,
      // Method 2: MS Word 2007 XML (good compatibility)
      `soffice --headless --convert-to docx:"MS Word 2007 XML" --outdir "${outputDir}" "${inputPdfPath}"`,
      // Method 3: Standard DOCX format
      `soffice --headless --convert-to docx --outdir "${outputDir}" "${inputPdfPath}"`,
      // Method 4: RTF format (fallback)
      `soffice --headless --convert-to rtf --outdir "${outputDir}" "${inputPdfPath}"`
    ];
    
    for (let i = 0; i < conversionMethods.length; i++) {
      try {
        const command = conversionMethods[i];
        console.log(`Attempting LibreOffice method ${i + 1}:`, command);
        
        const { stdout, stderr } = await execAsync(command, { timeout: 180000 }); // 3 minutes timeout
        
        if (stderr) {
          console.log('LibreOffice stderr:', stderr);
        }
        
        console.log('LibreOffice stdout:', stdout);
        
        // Check for different output file extensions
        const possibleExtensions = ['.docx', '.rtf'];
        let convertedFile = null;
        
        for (const ext of possibleExtensions) {
          const expectedPath = inputPdfPath.replace('.pdf', ext);
          if (fs.existsSync(expectedPath)) {
            convertedFile = expectedPath;
            break;
          }
        }
        
        if (convertedFile) {
          // Move to the desired output path
          await fs.move(convertedFile, outputDocxPath, { overwrite: true });
          console.log(`LibreOffice method ${i + 1} successful`);
          return outputDocxPath;
        }
      } catch (error) {
        console.log(`LibreOffice method ${i + 1} failed:`, error.message);
        if (i === conversionMethods.length - 1) {
          throw error;
        }
      }
    }
    
    throw new Error('All LibreOffice conversion methods failed');
  }

  /**
   * Convert using pdf2docx with advanced settings
   */
  static async convertWithPdf2DocxAdvanced(inputPdfPath, outputDocxPath) {
    try {
      const { Converter } = require('pdf2docx');
      const converter = new Converter(inputPdfPath);
      
      // Advanced conversion options for maximum formatting preservation
      const options = {
        start: 0,
        end: undefined,
        pages: undefined,
        zoom: 1.0,
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
          font_size_mapping: true,
          font_family_mapping: true,
          line_spacing: true,
          text_align: true,
          text_color: true,
          text_bold: true,
          text_italic: true,
          text_underline: true
        },
        // Enhanced layout settings
        layout_settings: {
          page_margin: true,
          page_size: true,
          header_footer: true,
          page_break: true,
          section_break: true
        }
      };
      
      await converter.convert(outputDocxPath, options);
      console.log('pdf2docx advanced conversion successful');
      return outputDocxPath;
    } catch (error) {
      throw new Error(`pdf2docx conversion failed: ${error.message}`);
    }
  }

  /**
   * Convert using enhanced pdf-parse with better text processing
   */
  static async convertWithPdfParseAdvanced(inputPdfPath, outputDocxPath) {
    try {
      const pdf = require('pdf-parse');
      const dataBuffer = await fs.readFile(inputPdfPath);
      
      // Enhanced pdf-parse options
      const options = {
        normalizeWhitespace: false,
        disableCombineTextItems: false
      };
      
      const data = await pdf(dataBuffer, options);
      
      // Advanced text processing to preserve formatting
      const enhancedText = this.processTextWithAdvancedFormatting(data.text, data);
      
      // Create DOCX with advanced formatting
      await DocxProcessor.createDocxFromTextAdvanced(enhancedText, outputDocxPath);
      console.log('Enhanced pdf-parse conversion successful');
      return outputDocxPath;
    } catch (error) {
      throw new Error(`Enhanced pdf-parse conversion failed: ${error.message}`);
    }
  }

  /**
   * Convert using basic text extraction (fallback)
   */
  static async convertWithBasicExtraction(inputPdfPath, outputDocxPath) {
    try {
      const pdf = require('pdf-parse');
      const dataBuffer = await fs.readFile(inputPdfPath);
      const data = await pdf(dataBuffer);
      
      // Basic text processing
      const processedText = this.processTextBasic(data.text);
      
      // Create basic DOCX
      await DocxProcessor.createDocxFromText(processedText, outputDocxPath);
      console.log('Basic text extraction conversion successful');
      return outputDocxPath;
    } catch (error) {
      throw new Error(`Basic text extraction failed: ${error.message}`);
    }
  }

  /**
   * Advanced text processing with formatting preservation
   */
  static processTextWithAdvancedFormatting(text, pdfData) {
    if (!text) return '';
    
    let processedText = text;
    
    // Preserve line breaks and paragraph structure
    processedText = processedText
      .replace(/\n\s*\n/g, '\n\n')
      .replace(/\r\n/g, '\n')
      .trim();
    
    // Split into lines for processing
    const lines = processedText.split('\n');
    const processedLines = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      if (!line) {
        processedLines.push('');
        continue;
      }
      
      // Enhanced formatting detection
      if (this.isHeader(line)) {
        processedLines.push(`# ${line}`);
      } else if (this.isSubHeader(line)) {
        processedLines.push(`## ${line}`);
      } else if (this.isTableRow(line)) {
        processedLines.push(this.formatTableRow(line));
      } else if (this.isNumberedList(line)) {
        processedLines.push(`## ${line}`);
      } else if (this.isBulletPoint(line)) {
        processedLines.push(`### ${line}`);
      } else if (this.isEmphasizedText(line)) {
        processedLines.push(`**${line}**`);
      } else {
        processedLines.push(line);
      }
    }
    
    return processedLines.join('\n');
  }

  /**
   * Basic text processing
   */
  static processTextBasic(text) {
    if (!text) return '';
    
    return text
      .replace(/\n\s*\n/g, '\n\n')
      .replace(/\r\n/g, '\n')
      .trim();
  }

  /**
   * Check if line is a header
   */
  static isHeader(line) {
    return line === line.toUpperCase() && line.length > 3 && line.length < 100;
  }

  /**
   * Check if line is a sub-header
   */
  static isSubHeader(line) {
    return line.length > 5 && line.length < 80 && 
           (line.endsWith(':') || /^[A-Z][a-z]/.test(line));
  }

  /**
   * Check if line is a table row
   */
  static isTableRow(line) {
    const spaceCount = (line.match(/\s{2,}/g) || []).length;
    const tabCount = (line.match(/\t/g) || []).length;
    return spaceCount > 0 || tabCount > 0;
  }

  /**
   * Format table row
   */
  static formatTableRow(line) {
    return line.replace(/\s{2,}/g, '\t');
  }

  /**
   * Check if line is a numbered list
   */
  static isNumberedList(line) {
    return /^\d+\.\s/.test(line);
  }

  /**
   * Check if line is a bullet point
   */
  static isBulletPoint(line) {
    return /^[•\-\*]\s/.test(line);
  }

  /**
   * Check if line is emphasized text
   */
  static isEmphasizedText(line) {
    return line.length > 10 && line.length < 200 && 
           (line.includes('**') || line.includes('__') || 
            line === line.toUpperCase() || /^[A-Z]/.test(line));
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

module.exports = AdvancedPdfToWordConverter;


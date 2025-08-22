const puppeteer = require('puppeteer');
const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);
let libreConvert;

class PdfConverter {
  /**
   * Convert text content to PDF
   * @param {string} content - Text content to convert
   * @param {string} outputPath - Path where to save the PDF
   * @param {Object} options - PDF generation options
   * @returns {string} - Path to generated PDF
   */
  static async convertTextToPdf(content, outputPath, options = {}) {
    const {
      format = 'A4',
      margin = { top: '1in', right: '1in', bottom: '1in', left: '1in' },
      fontSize = '14px',
      fontFamily = 'Arial, sans-serif',
      lineHeight = '1.6',
      title = 'Document'
    } = options;

    try {
      const browser = await puppeteer.launch({ 
        headless: 'new',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
      });
      
      const page = await browser.newPage();
      
      // Create HTML content with better formatting
      const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <title>${title}</title>
          <style>
            @page {
              size: ${format};
              margin: ${margin.top} ${margin.right} ${margin.bottom} ${margin.left};
            }
            
            body {
              font-family: ${fontFamily};
              font-size: ${fontSize};
              line-height: ${lineHeight};
              color: #333;
              margin: 0;
              padding: 20px;
            }
            
            .content {
              white-space: normal;
              word-wrap: break-word;
            }
            
            .content p {
              margin: 0.5em 0;
              text-align: justify;
            }
            
            .content h1, .content h2, .content h3, .content h4, .content h5, .content h6 {
              margin: 1em 0 0.5em 0;
              page-break-after: avoid;
            }
            
            .page-break {
              page-break-before: always;
            }
            
            h1, h2, h3, h4, h5, h6 {
              margin-top: 1.5em;
              margin-bottom: 0.5em;
              font-weight: bold;
            }
            
            h1 { font-size: 1.8em; }
            h2 { font-size: 1.5em; }
            h3 { font-size: 1.3em; }
            h4 { font-size: 1.1em; }
            
            p {
              margin: 0.5em 0;
            }
          </style>
        </head>
        <body>
          <div class="content">
            ${this.formatContentForHtml(content)}
          </div>
        </body>
        </html>
      `;

      await page.setContent(htmlContent);
      
      // Generate PDF
      await page.pdf({
        path: outputPath,
        format: format,
        margin: margin,
        printBackground: true,
        displayHeaderFooter: false
      });

      await browser.close();
      
      return outputPath;
    } catch (error) {
      throw new Error(`Failed to convert to PDF: ${error.message}`);
    }
  }

  /**
   * Convert HTML content to PDF
   * @param {string} html - Full or partial HTML content
   * @param {string} outputPath - Path where to save the PDF
   * @param {Object} options - PDF generation options
   * @returns {string} - Path to generated PDF
   */
  static async convertHtmlToPdf(html, outputPath, options = {}) {
    const {
      format = 'A4',
      margin = { top: '1in', right: '1in', bottom: '1in', left: '1in' },
      title = 'Document',
      customCss = ''
    } = options;

    try {
      const browser = await puppeteer.launch({
        headless: 'new',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
      });
      const page = await browser.newPage();

      const htmlContent = `
        <!DOCTYPE html>
        <html>
          <head>
            <meta charset="UTF-8">
            <title>${title}</title>
            <style>
              @page { size: ${format}; margin: ${margin.top} ${margin.right} ${margin.bottom} ${margin.left}; }
              body { font-family: Arial, sans-serif; color: #333; margin: 0; }
              .content { padding: 20px; }
              table { border-collapse: collapse; width: 100%; margin: 12px 0; }
              th, td { border: 1px solid #999; padding: 6px 8px; }
              h1, h2, h3, h4, h5, h6 { page-break-after: avoid; }
              ${customCss}
            </style>
          </head>
          <body>
            <div class="content">${html}</div>
          </body>
        </html>`;

      await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
      await page.pdf({
        path: outputPath,
        format,
        margin,
        printBackground: true
      });
      await browser.close();
      return outputPath;
    } catch (error) {
      throw new Error(`Failed to convert HTML to PDF: ${error.message}`);
    }
  }

  /**
   * Check if LibreOffice is available on the system
   * @returns {boolean} - True if LibreOffice is available
   */
  static async isLibreOfficeAvailable() {
    try {
      // Try to find soffice command
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

  /**
   * Convert DOCX to PDF using LibreOffice command line (best fidelity)
   * @param {string} inputDocxPath
   * @param {string} outputPdfPath
   */
  static async convertDocxToPdfLibreOffice(inputDocxPath, outputPdfPath) {
    try {
      const inputDir = path.dirname(inputDocxPath);
      const outputDir = path.dirname(outputPdfPath);
      
      // Use LibreOffice command line for best results
      const command = `soffice --headless --convert-to pdf --outdir "${outputDir}" "${inputDocxPath}"`;
      console.log('Executing LibreOffice command:', command);
      
      const { stdout, stderr } = await execAsync(command, { timeout: 30000 });
      
      if (stderr) {
        console.log('LibreOffice stderr:', stderr);
      }
      
      console.log('LibreOffice stdout:', stdout);
      
      // LibreOffice creates the PDF with the same name but .pdf extension
      const expectedPdfPath = inputDocxPath.replace('.docx', '.pdf');
      
      if (fs.existsSync(expectedPdfPath)) {
        // Move to the desired output path
        await fs.move(expectedPdfPath, outputPdfPath, { overwrite: true });
        return outputPdfPath;
      } else {
        throw new Error('LibreOffice did not create the expected PDF file');
      }
    } catch (error) {
      throw new Error(`LibreOffice conversion failed: ${error.message}`);
    }
  }

  /**
   * Convert DOCX to PDF using libreoffice-convert library
   * @param {string} inputDocxPath
   * @param {string} outputPdfPath
   */
  static async convertDocxToPdfLibreConvert(inputDocxPath, outputPdfPath) {
    try {
      if (!libreConvert) {
        libreConvert = require('libreoffice-convert');
      }

      const docxBuffer = await fs.readFile(inputDocxPath);
      const pdfBuffer = await new Promise((resolve, reject) => {
        libreConvert.convert(docxBuffer, '.pdf', undefined, (err, done) => {
          if (err) return reject(err);
          resolve(done);
        });
      });
      await fs.writeFile(outputPdfPath, pdfBuffer);
      return outputPdfPath;
    } catch (error) {
      throw new Error(`LibreOffice convert library failed: ${error.message}`);
    }
  }

  /**
   * Convert DOCX to PDF with multiple fallback options
   * @param {string} inputDocxPath
   * @param {string} outputPdfPath
   */
  static async convertDocxToPdfNative(inputDocxPath, outputPdfPath) {
    console.log('Starting DOCX to PDF conversion with multiple fallback options...');
    
    // Method 1: Try LibreOffice command line (best fidelity)
    try {
      console.log('Attempting LibreOffice command line conversion...');
      return await this.convertDocxToPdfLibreOffice(inputDocxPath, outputPdfPath);
    } catch (error) {
      console.log('LibreOffice command line failed:', error.message);
    }
    
    // Method 2: Try libreoffice-convert library
    try {
      console.log('Attempting libreoffice-convert library conversion...');
      return await this.convertDocxToPdfLibreConvert(inputDocxPath, outputPdfPath);
    } catch (error) {
      console.log('LibreOffice convert library failed:', error.message);
    }
    
    // Method 3: Fallback to HTML conversion (loses formatting but works)
    console.log('Falling back to HTML conversion (formatting may be lost)...');
    const mammoth = require('mammoth');
    const { value: html, messages } = await mammoth.convertToHtml({ path: inputDocxPath }, {
      styleMap: [
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "table => table",
        "tr => tr",
        "td => td",
        "th => th"
      ]
    });
    
    if (messages && messages.length) {
      console.log('Mammoth conversion messages:', messages);
    }
    
    return await this.convertHtmlToPdf(html, outputPdfPath, {
      customCss: `
        table { border-collapse: collapse; width: 100%; margin: 12px 0; }
        th, td { border: 1px solid #999; padding: 6px 8px; text-align: left; }
        th { background-color: #f5f5f5; font-weight: bold; }
      `
    });
  }

  /**
   * Format content for HTML display
   * @param {string} content - Raw text content
   * @returns {string} - Formatted HTML content
   */
  static formatContentForHtml(content) {
    console.log('Formatting content for HTML:', content);
    const lines = content.split('\n');
    console.log('Split into lines:', lines);
    
    const formattedLines = [];
    let inParagraph = false;
    let paragraphContent = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      // Handle headings
      if (line.startsWith('#')) {
        // Close any open paragraph
        if (inParagraph && paragraphContent.length > 0) {
          formattedLines.push(`<p>${paragraphContent.join(' ')}</p>`);
          paragraphContent = [];
          inParagraph = false;
        }
        
        const level = line.match(/^#+/)[0].length;
        const text = line.replace(/^#+\s*/, '');
        formattedLines.push(`<h${level}>${text}</h${level}>`);
        continue;
      }
      
      // Handle empty lines
      if (line === '') {
        // Close any open paragraph
        if (inParagraph && paragraphContent.length > 0) {
          formattedLines.push(`<p>${paragraphContent.join(' ')}</p>`);
          paragraphContent = [];
          inParagraph = false;
        }
        // Don't add <br> for empty lines - let CSS handle spacing
        continue;
      }
      
      // Handle regular content
      if (!inParagraph) {
        inParagraph = true;
        paragraphContent = [line];
      } else {
        paragraphContent.push(line);
      }
    }
    
    // Close any remaining paragraph
    if (inParagraph && paragraphContent.length > 0) {
      formattedLines.push(`<p>${paragraphContent.join(' ')}</p>`);
    }
    
    const result = formattedLines.join('\n');
    console.log('Formatted HTML result:', result);
    return result;
  }

  /**
   * Clean content by removing excessive whitespace
   * @param {string} content - Raw content
   * @returns {string} - Cleaned content
   */
  static cleanContent(content) {
    return content
      .replace(/\r\n/g, '\n') // Normalize line endings
      .replace(/\r/g, '\n')   // Handle old Mac line endings
      .replace(/\n\s*\n\s*\n/g, '\n\n') // Remove multiple consecutive empty lines
      .replace(/[ \t]+/g, ' ') // Replace multiple spaces/tabs with single space
      .trim(); // Remove leading/trailing whitespace
  }

  /**
   * Convert DOCX content to PDF
   * @param {string} docxContent - Content extracted from DOCX
   * @param {string} outputPath - Path for PDF output
   * @param {Object} options - PDF options
   * @returns {string} - Path to generated PDF
   */
  static async convertDocxContentToPdf(docxContent, outputPath, options = {}) {
    console.log('PDF Converter - Original content:', docxContent);
    const cleanedContent = this.cleanContent(docxContent);
    console.log('PDF Converter - Cleaned content:', cleanedContent);
    return this.convertTextToPdf(cleanedContent, outputPath, {
      title: 'Document',
      ...options
    });
  }

  /**
   * Generate PDF with custom styling
   * @param {string} content - Text content
   * @param {string} outputPath - Output path
   * @param {Object} styling - Custom styling options
   * @returns {string} - Path to generated PDF
   */
  static async generateStyledPdf(content, outputPath, styling = {}) {
    const {
      backgroundColor = '#ffffff',
      textColor = '#333333',
      accentColor = '#007bff',
      showHeader = true,
      showFooter = true,
      customCss = ''
    } = styling;

    const browser = await puppeteer.launch({ headless: 'new' });
    const page = await browser.newPage();
    
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: ${textColor};
            background-color: ${backgroundColor};
            margin: 0;
            padding: 40px;
          }
          
          .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
          }
          
          .header {
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 3px solid ${accentColor};
            ${!showHeader ? 'display: none;' : ''}
          }
          
          .content {
            white-space: normal;
            word-wrap: break-word;
          }
          
          .content p {
            margin: 0.5em 0;
            text-align: justify;
          }
          
          .content h1, .content h2, .content h3, .content h4, .content h5, .content h6 {
            margin: 1em 0 0.5em 0;
            page-break-after: avoid;
          }
          
          .footer {
            text-align: center;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            font-size: 12px;
            color: #666;
            ${!showFooter ? 'display: none;' : ''}
          }
          
          ${customCss}
        </style>
      </head>
      <body>
        <div class="container">
          ${showHeader ? `
            <div class="header">
              <h1 style="color: ${accentColor};">Document Template</h1>
              <p>Generated on ${new Date().toLocaleDateString()}</p>
            </div>
          ` : ''}
          
          <div class="content">
            ${this.formatContentForHtml(content)}
          </div>
          
          ${showFooter ? `
            <div class="footer">
              <p>Generated by FileFlow Template Editor</p>
            </div>
          ` : ''}
        </div>
      </body>
      </html>
    `;

    await page.setContent(htmlContent);
    
    await page.pdf({
      path: outputPath,
      format: 'A4',
      margin: { top: '0.5in', right: '0.5in', bottom: '0.5in', left: '0.5in' },
      printBackground: true
    });

    await browser.close();
    
    return outputPath;
  }
}

module.exports = PdfConverter;


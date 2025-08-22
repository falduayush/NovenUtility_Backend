const mammoth = require('mammoth');
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require('docx');
const fs = require('fs-extra');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

class DocxProcessor {
  /**
   * Extract text and variables from DOCX file
   * @param {string} filePath - Path to the DOCX file
   * @returns {Object} - Object containing text and variables
   */
  static async extractTextAndVariables(filePath) {
    try {
      const result = await mammoth.extractRawText({ path: filePath });
      const text = result.value;

      // Collect variables from main text + headers/footers
      const variableRegex = /\{\{\s*([^}]+)\s*\}\}/g;
      const variables = new Set();
      let match;

      // From body text
      while ((match = variableRegex.exec(text)) !== null) {
        variables.add(match[1].trim());
      }

      // From headers and footers (scan XML and strip tags first)
      try {
        const zip = new PizZip(await fs.readFile(filePath, 'binary'));
        const headerFooterFiles = zip.file(/word\/(header\d+|footer\d+)\.xml/);
        headerFooterFiles.forEach(f => {
          const xml = f.asText();
          const plain = xml.replace(/<[^>]+>/g, '');
          let m;
          while ((m = variableRegex.exec(plain)) !== null) {
            variables.add(m[1].trim());
          }
        });
      } catch (_) {
        // Best-effort; ignore extraction errors here
      }

      return {
        text: text,
        variables: Array.from(variables),
        messages: result.messages
      };
    } catch (error) {
      throw new Error(`Failed to extract text from DOCX: ${error.message}`);
    }
  }

  /**
   * Replace variables in text content
   * @param {string} text - Original text content
   * @param {Object} variables - Object with variable names as keys and values
   * @returns {string} - Text with replaced variables
   */
  static replaceVariables(text, variables) {
    let processedText = text;

    Object.keys(variables).forEach(variableName => {
      const regex = new RegExp(`\\{\\{\\s*${variableName}\\s*\\}\\}`, 'g');
      processedText = processedText.replace(regex, variables[variableName] || '');
    });

    return processedText;
  }

  /**
   * Create DOCX document from text content with enhanced formatting
   * @param {string} content - Text content
   * @param {string} outputPath - Path where to save the DOCX file
   */
  static async createDocxFromText(content, outputPath) {
    try {
      // Split content into paragraphs
      const paragraphs = content.split('\n').filter(line => line.trim() !== '');

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              margin: {
                top: 1440, // 1 inch
                right: 1440,
                bottom: 1440,
                left: 1440
              }
            }
          },
          children: paragraphs.map(paragraph => {
            // Check if it's a heading (starts with #)
            if (paragraph.startsWith('#')) {
              const level = paragraph.match(/^#+/)[0].length;
              const text = paragraph.replace(/^#+\s*/, '');
              
              return new Paragraph({
                text: text,
                heading: level === 1 ? HeadingLevel.HEADING_1 : 
                        level === 2 ? HeadingLevel.HEADING_2 : 
                        level === 3 ? HeadingLevel.HEADING_3 : 
                        HeadingLevel.HEADING_4,
                spacing: {
                  before: 240, // 12pt before
                  after: 120   // 6pt after
                }
              });
            } 
            // Check if it might be a table row (contains multiple spaces or tabs)
            else if (paragraph.includes('  ') || paragraph.includes('\t')) {
              // Try to create a table from this line
              const cells = paragraph.split(/\s{2,}|\t/).filter(cell => cell.trim() !== '');
              if (cells.length > 1) {
                return new Paragraph({
                  children: [
                    new TextRun({
                      text: paragraph,
                      size: 24,
                      font: 'Courier New' // Use monospace font for table-like content
                    })
                  ],
                  spacing: {
                    before: 120,
                    after: 120
                  }
                });
              }
            }
            // Check if it's a numbered list
            else if (/^\d+\.\s/.test(paragraph)) {
              return new Paragraph({
                children: [
                  new TextRun({
                    text: paragraph,
                    size: 24
                  })
                ],
                numbering: {
                  reference: "default-numbering",
                  level: 0
                },
                spacing: {
                  before: 60,
                  after: 60
                }
              });
            }
            // Check if it's a bullet point
            else if (/^[•\-\*]\s/.test(paragraph)) {
              return new Paragraph({
                children: [
                  new TextRun({
                    text: paragraph,
                    size: 24
                  })
                ],
                bullet: {
                  level: 0
                },
                spacing: {
                  before: 60,
                  after: 60
                }
              });
            }
            // Regular paragraph
            else {
              return new Paragraph({
                children: [
                  new TextRun({
                    text: paragraph,
                    size: 24
                  })
                ],
                spacing: {
                  before: 120,
                  after: 120
                }
              });
            }
          })
        }]
      });

      const buffer = await Packer.toBuffer(doc);
      await fs.writeFile(outputPath, buffer);

      return outputPath;
    } catch (error) {
      throw new Error(`Failed to create DOCX: ${error.message}`);
    }
  }

  /**
   * Create DOCX document from text content with advanced formatting
   * @param {string} content - Text content
   * @param {string} outputPath - Path where to save the DOCX file
   */
  static async createDocxFromTextAdvanced(content, outputPath) {
    try {
      // Split content into paragraphs
      const paragraphs = content.split('\n').filter(line => line.trim() !== '');

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              margin: {
                top: 1440, // 1 inch
                right: 1440,
                bottom: 1440,
                left: 1440
              },
              size: {
                width: 11906, // A4 width
                height: 16838  // A4 height
              }
            }
          },
          children: paragraphs.map(paragraph => {
            // Check if it's a heading (starts with #)
            if (paragraph.startsWith('#')) {
              const level = paragraph.match(/^#+/)[0].length;
              const text = paragraph.replace(/^#+\s*/, '');
              
              return new Paragraph({
                text: text,
                heading: level === 1 ? HeadingLevel.HEADING_1 : 
                        level === 2 ? HeadingLevel.HEADING_2 : 
                        level === 3 ? HeadingLevel.HEADING_3 : 
                        HeadingLevel.HEADING_4,
                spacing: {
                  before: 240, // 12pt before
                  after: 120   // 6pt after
                },
                alignment: AlignmentType.CENTER
              });
            } 
            // Check if it's emphasized text (starts with **)
            else if (paragraph.startsWith('**') && paragraph.endsWith('**')) {
              const text = paragraph.replace(/\*\*/g, '');
              return new Paragraph({
                children: [
                  new TextRun({
                    text: text,
                    size: 28, // Larger font
                    bold: true
                  })
                ],
                spacing: {
                  before: 120,
                  after: 120
                },
                alignment: AlignmentType.CENTER
              });
            }
            // Check if it might be a table row (contains multiple spaces or tabs)
            else if (paragraph.includes('  ') || paragraph.includes('\t')) {
              // Try to create a table from this line
              const cells = paragraph.split(/\s{2,}|\t/).filter(cell => cell.trim() !== '');
              if (cells.length > 1) {
                return new Paragraph({
                  children: [
                    new TextRun({
                      text: paragraph,
                      size: 24,
                      font: 'Courier New' // Use monospace font for table-like content
                    })
                  ],
                  spacing: {
                    before: 120,
                    after: 120
                  }
                });
              }
            }
            // Check if it's a numbered list
            else if (/^\d+\.\s/.test(paragraph)) {
              return new Paragraph({
                children: [
                  new TextRun({
                    text: paragraph,
                    size: 24
                  })
                ],
                numbering: {
                  reference: "default-numbering",
                  level: 0
                },
                spacing: {
                  before: 60,
                  after: 60
                }
              });
            }
            // Check if it's a bullet point
            else if (/^[•\-\*]\s/.test(paragraph)) {
              return new Paragraph({
                children: [
                  new TextRun({
                    text: paragraph,
                    size: 24
                  })
                ],
                bullet: {
                  level: 0
                },
                spacing: {
                  before: 60,
                  after: 60
                }
              });
            }
            // Regular paragraph
            else {
              return new Paragraph({
                children: [
                  new TextRun({
                    text: paragraph,
                    size: 24
                  })
                ],
                spacing: {
                  before: 120,
                  after: 120
                }
              });
            }
          })
        }]
      });

      const buffer = await Packer.toBuffer(doc);
      await fs.writeFile(outputPath, buffer);

      return outputPath;
    } catch (error) {
      throw new Error(`Failed to create advanced DOCX: ${error.message}`);
    }
  }

  /**
   * Process DOCX template with variables
   * @param {string} templatePath - Path to template DOCX file
   * @param {Object} variables - Variables to replace
   * @param {string} outputPath - Path for output file
   * @returns {string} - Path to generated file
   */
  static async processTemplate(templatePath, variables, outputPath) {
    try {
      // Read the existing DOCX so we can edit in-place and preserve formatting/tables
      const content = await fs.readFile(templatePath, 'binary');

      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '{{', end: '}}' }
      });

      // Ensure undefined variables don't break rendering
      const safeVariables = Object.fromEntries(
        Object.entries(variables || {}).map(([key, value]) => [key, value == null ? '' : value])
      );

      // Use modern API: pass data directly to render
      doc.render(safeVariables);

      const buffer = doc.getZip().generate({ type: 'nodebuffer' });
      await fs.writeFile(outputPath, buffer);
      return outputPath;
    } catch (error) {
      // Add more context from docxtemplater if available
      if (error && error.properties && error.properties.errors) {
        const explanation = error.properties.errors
          .map(e => `${e.properties && e.properties.explanation ? e.properties.explanation : e.message}`)
          .join('; ');
        throw new Error(`Failed to process template: ${explanation}`);
      }
      throw new Error(`Failed to process template: ${error.message}`);
    }
  }

  /**
   * Validate DOCX file
   * @param {string} filePath - Path to DOCX file
   * @returns {boolean} - True if valid DOCX file
   */
  static async validateDocxFile(filePath) {
    try {
      const result = await mammoth.extractRawText({ path: filePath });
      return result.value.length > 0;
    } catch (error) {
      return false;
    }
  }

  /**
   * Get file statistics
   * @param {string} filePath - Path to DOCX file
   * @returns {Object} - File statistics
   */
  static async getFileStats(filePath) {
    try {
      const result = await mammoth.extractRawText({ path: filePath });
      const text = result.value;
      
      const wordCount = text.split(/\s+/).filter(word => word.length > 0).length;
      const characterCount = text.length;
      const lineCount = text.split('\n').length;
      
      // Count variables
      const variableRegex = /\{\{\s*([^}]+)\s*\}\}/g;
      const variables = new Set();
      let match;

      while ((match = variableRegex.exec(text)) !== null) {
        variables.add(match[1].trim());
      }

      return {
        wordCount,
        characterCount,
        lineCount,
        variableCount: variables.size,
        variables: Array.from(variables)
      };
    } catch (error) {
      throw new Error(`Failed to get file stats: ${error.message}`);
    }
  }
}

module.exports = DocxProcessor;



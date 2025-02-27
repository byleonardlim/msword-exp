// Commands.js - Contains document structure and manipulation functions

Office.onReady(() => {
    // Register global command functions
    window.structureFunctions = {
        groupParagraphsByHeader,
        highlightParagraphGroup,
        formatSection,
        addCommentToSection,
        findHeadingsByLevel, // New function for finding headings by level
        navigateToHeading,    // New function for navigating to a heading
        getDocumentOutline    // New function to get the full document structure
    };
});

/**
 * Groups paragraphs under their respective headers
 * @returns {Promise<Object>} Structure object with headers and their paragraphs
 */
async function groupParagraphsByHeader() {
    return await Word.run(async (context) => {
        // Get all paragraphs
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load(['text', 'style', 'font']);
        await context.sync();
        
        const structure = [];
        let currentHeader = null;
        let currentLevel = 0;
        
        // Process each paragraph
        for (let i = 0; i < paragraphs.items.length; i++) {
            const paragraph = paragraphs.items[i];
            const style = paragraph.style;
            
            // Check if it's a heading (Office uses styles like "Heading 1", "Heading 2", etc.)
            if (style.includes('Heading')) {
                // Extract level from style name (Heading 1 -> level 1)
                const level = parseInt(style.replace(/\D/g, '')) || 1;
                
                // Create a new header object
                currentHeader = {
                    text: paragraph.text,
                    level: level,
                    paragraphs: [],
                    index: i
                };
                
                currentLevel = level;
                structure.push(currentHeader);
            } 
            else if (paragraph.text.trim() !== '') {
                // Regular paragraph - assign to current header if one exists
                if (currentHeader) {
                    currentHeader.paragraphs.push({
                        text: paragraph.text,
                        index: i
                    });
                } else {
                    // Paragraph without a header - treat as its own section
                    structure.push({
                        text: paragraph.text,
                        level: 0,
                        paragraphs: [],
                        index: i
                    });
                }
            }
        }
        
        return structure;
    });
}

/**
 * Highlight paragraphs that belong to a specific header
 * @param {number} headerIndex - Index of the header in the document
 * @param {string} highlightColor - Color to use for highlighting (e.g. 'yellow')
 */
async function highlightParagraphGroup(headerIndex, highlightColor = 'yellow') {
    await Word.run(async (context) => {
        // Get structure to find all paragraphs under the header
        const structure = await groupParagraphsByHeader();
        const headerItem = structure.find(item => item.index === headerIndex);
        
        if (!headerItem) {
            throw new Error(`Header with index ${headerIndex} not found`);
        }
        
        // Get all paragraphs
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        
        // Highlight the header
        paragraphs.items[headerItem.index].font.highlightColor = highlightColor;
        
        // Highlight all paragraphs under this header
        headerItem.paragraphs.forEach(para => {
            paragraphs.items[para.index].font.highlightColor = highlightColor;
        });
        
        await context.sync();
    });
}

/**
 * Format a section (header and its paragraphs)
 * @param {number} headerIndex - Index of the header in the document
 * @param {object} formatting - Formatting to apply { bold, italic, underline, color, etc. }
 */
async function formatSection(headerIndex, formatting = {}) {
    await Word.run(async (context) => {
        // Get structure to find all paragraphs under the header
        const structure = await groupParagraphsByHeader();
        const headerItem = structure.find(item => item.index === headerIndex);
        
        if (!headerItem) {
            throw new Error(`Header with index ${headerIndex} not found`);
        }
        
        // Get all paragraphs
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load(['items', 'font']);
        await context.sync();
        
        // Apply formatting to the header and its paragraphs
        const paraIndexes = [headerItem.index, ...headerItem.paragraphs.map(p => p.index)];
        
        paraIndexes.forEach(index => {
            const font = paragraphs.items[index].font;
            
            // Apply each formatting property if specified
            if (formatting.bold !== undefined) font.bold = formatting.bold;
            if (formatting.italic !== undefined) font.italic = formatting.italic;
            if (formatting.underline !== undefined) font.underline = formatting.underline;
            if (formatting.color) font.color = formatting.color;
            if (formatting.size) font.size = formatting.size;
            if (formatting.font) font.name = formatting.font;
        });
        
        await context.sync();
    });
}

/**
 * Add a comment to a section (header and optionally its paragraphs)
 * @param {number} headerIndex - Index of the header in the document
 * @param {string} commentText - The comment text to add
 * @param {boolean} includeChildParagraphs - Whether to comment on child paragraphs as well
 */
async function addCommentToSection(headerIndex, commentText, includeChildParagraphs = false) {
    await Word.run(async (context) => {
        // Get structure to find all paragraphs under the header
        const structure = await groupParagraphsByHeader();
        const headerItem = structure.find(item => item.index === headerIndex);
        
        if (!headerItem) {
            throw new Error(`Header with index ${headerIndex} not found`);
        }
        
        // Get all paragraphs
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        
        // Add comment to the header
        paragraphs.items[headerItem.index].insertComment(commentText);
        
        // Add comments to child paragraphs if requested
        if (includeChildParagraphs && headerItem.paragraphs.length > 0) {
            headerItem.paragraphs.forEach(para => {
                paragraphs.items[para.index].insertComment(commentText);
            });
        }
        
        await context.sync();
    });
}

/**
 * Find headings by level
 * @param {number} level - The heading level to find (1 for Heading 1, 2 for Heading 2, etc.)
 * @returns {Promise<Array>} - Array of heading paragraphs
 */
async function findHeadingsByLevel(level) {
    return await Word.run(async (context) => {
        // Get all paragraphs
        const allParagraphs = context.document.body.paragraphs;
        allParagraphs.load(["text", "style", "id"]);
        await context.sync();
        
        const headings = [];
        const targetStyle = `Heading ${level}`;
        
        for (let i = 0; i < allParagraphs.items.length; i++) {
            if (allParagraphs.items[i].style === targetStyle) {
                headings.push({
                    id: allParagraphs.items[i].id,
                    text: allParagraphs.items[i].text,
                    index: i
                });
            }
        }
        
        return headings;
    });
}

/**
 * Navigate to a specific heading
 * @param {string} headingText - The text of the heading to navigate to
 * @param {boolean} exactMatch - Whether to require an exact match (true) or partial match (false)
 * @returns {Promise<boolean>} - True if heading was found and selected, false otherwise
 */
async function navigateToHeading(headingText, exactMatch = false) {
    return await Word.run(async (context) => {
        // Search for the heading
        const searchOptions = {
            matchCase: false,
            matchWholeWord: exactMatch
        };
        
        const searchResults = context.document.body.search(headingText, searchOptions);
        searchResults.load(["text", "style", "font"]);
        await context.sync();
        
        // Find the first result that's a heading
        let headingFound = false;
        for (let i = 0; i < searchResults.items.length; i++) {
            if (searchResults.items[i].style.includes("Heading")) {
                searchResults.items[i].select();
                headingFound = true;
                break;
            }
        }
        
        await context.sync();
        return headingFound;
    });
}

/**
 * Get document outline based on headings
 * @returns {Promise<Array>} - Array of headings with their level and text
 */
async function getDocumentOutline() {
    return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load(["text", "style", "id"]);
        await context.sync();
        
        const outline = [];
        for (let i = 0; i < paragraphs.items.length; i++) {
            const para = paragraphs.items[i];
            if (para.style.includes("Heading")) {
                const level = parseInt(para.style.replace(/\D/g, ''));
                outline.push({
                    id: para.id,
                    text: para.text,
                    level: level,
                    index: i
                });
            }
        }
        
        return outline;
    });
}
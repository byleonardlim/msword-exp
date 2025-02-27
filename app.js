// Office.js initialization
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Initialize event handlers
        document.getElementById('saveApiKey').onclick = saveApiKey;
        document.getElementById('suggestChanges').onclick = suggestChanges;
        document.getElementById('fixGrammar').onclick = fixGrammar;
        document.getElementById('summarizeSelection').onclick = summarizeSelection;
        
        // Load API key if saved
        loadApiKey();
        
        // Load document structure
        loadDocumentStructure();
        
        // Enable live suggestions as user types
        setupLiveSuggestions();
    }
});

// API Key Management using LocalStorage
function saveApiKey() {
    const apiKey = document.getElementById('apiKey').value;
    if (apiKey) {
        localStorage.setItem('openai_api_key', apiKey);
        document.getElementById('apiKeyStatus').innerText = 'API Key saved successfully!';
        document.getElementById('apiKeyStatus').style.color = 'green';
    } else {
        document.getElementById('apiKeyStatus').innerText = 'Please enter an API Key';
        document.getElementById('apiKeyStatus').style.color = 'red';
    }
}

function loadApiKey() {
    const apiKey = localStorage.getItem('openai_api_key');
    if (apiKey) {
        document.getElementById('apiKey').value = apiKey;
        document.getElementById('apiKeyStatus').innerText = 'API Key loaded from storage';
        document.getElementById('apiKeyStatus').style.color = 'green';
    }
}

function getApiKey() {
    return localStorage.getItem('openai_api_key');
}

// Document Structure Navigation
async function loadDocumentStructure() {
    await Word.run(async (context) => {
        // Get all headings in the document
        const headingsParagraphs = context.document.body.paragraphs;
        headingsParagraphs.load('text, style');
        
        await context.sync();
        
        const headingsTree = document.getElementById('headingsTree');
        headingsTree.innerHTML = '';
        
        // Process paragraphs to find headings
        for (let i = 0; i < headingsParagraphs.items.length; i++) {
            const paragraph = headingsParagraphs.items[i];
            const style = paragraph.style;
            
            // Check if paragraph is a heading (Heading 1-6)
            if (style && style.includes('Heading')) {
                const headingLevel = parseInt(style.replace('Heading ', ''));
                if (!isNaN(headingLevel) && headingLevel >= 1 && headingLevel <= 6) {
                    // Create heading item in the tree
                    const headingItem = document.createElement('div');
                    headingItem.className = `heading-item heading-h${headingLevel}`;
                    headingItem.innerText = paragraph.text;
                    headingItem.dataset.paragraphIndex = i;
                    
                    // Add click event to navigate to the heading
                    headingItem.onclick = () => navigateToHeading(i);
                    
                    headingsTree.appendChild(headingItem);
                }
            }
        }
    }).catch(handleError);
}

async function navigateToHeading(paragraphIndex) {
    await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('text');
        
        await context.sync();
        
        if (paragraphIndex < paragraphs.items.length) {
            // Select the paragraph
            paragraphs.items[paragraphIndex].select();
            
            // Find all paragraphs until the next heading or end of document
            let endIndex = paragraphIndex + 1;
            while (endIndex < paragraphs.items.length) {
                const nextParagraph = paragraphs.items[endIndex];
                const style = nextParagraph.style;
                
                if (style && style.includes('Heading')) {
                    break;
                }
                
                endIndex++;
            }
            
            // Select the range from heading to next heading (or end)
            if (endIndex > paragraphIndex + 1) {
                const range = paragraphs.items[paragraphIndex].getRange();
                const rangeEnd = paragraphs.items[endIndex - 1].getRange('End');
                
                range.expandTo(rangeEnd);
                range.select();
            }
        }
        
        await context.sync();
    }).catch(handleError);
}

// OpenAI Integration
async function callOpenAI(prompt, model = 'gpt-4-turbo') {
    const apiKey = getApiKey();
    if (!apiKey) {
        document.getElementById('aiSuggestions').innerHTML = 
            '<div class="suggestion-item">Please enter your OpenAI API Key first</div>';
        return null;
    }
    
    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: model,
                messages: [
                    {
                        role: 'system',
                        content: 'You are a text processing tool. Respond only with the exact output requested without any explanations, introductions, or additional text. Do not use phrases like "Here is" or "Here\'s". Never explain your reasoning or add notes. Just return the exact result.'
                    },
                    {
                        role: 'user',
                        content: prompt
                    }
                ],
                temperature: 0.3 // Lower temperature for more precise outputs
            })
        });
        
        if (!response.ok) {
            throw new Error(`API error: ${response.status}`);
        }
        
        const data = await response.json();
        return data.choices[0].message.content.trim();
    } catch (error) {
        document.getElementById('aiSuggestions').innerHTML = 
            `<div class="suggestion-item">Error calling OpenAI API: ${error.message}</div>`;
        return null;
    }
}

// AI Feature: Suggest Changes
async function suggestChanges() {
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        
        const text = selection.text;
        if (!text || text.trim() === '') {
            document.getElementById('aiSuggestions').innerHTML = 
                '<div class="suggestion-item">Please select some text first</div>';
            return;
        }
        
        document.getElementById('aiSuggestions').innerHTML = 
            '<div class="suggestion-item">Analyzing your text...</div>';
        
        const prompt = `Improve the following text:

${text}

Return a numbered list with exactly 3 specific improvements. Format each point as "1. [Issue]: [Suggestion]" without any introduction or conclusion.`;
        
        const suggestions = await callOpenAI(prompt);
        
        if (suggestions) {
            // Format the suggestions
            document.getElementById('aiSuggestions').innerHTML = 
                `<div class="suggestion-item">${suggestions.replace(/\n/g, '<br>')}</div>
                <button class="apply-button" id="applyChanges">Apply Suggestions</button>`;
                
            // Handle apply button
            document.getElementById('applyChanges').onclick = async () => {
                const improvedPrompt = `Rewrite this text with the improvements:

Original text: ${text}

Suggested improvements: ${suggestions}

Return ONLY the improved version with no explanation:`;
                
                const improvedText = await callOpenAI(improvedPrompt);
                
                if (improvedText) {
                    await Word.run(async (context) => {
                        const selection = context.document.getSelection();
                        selection.insertText(improvedText, 'Replace');
                        await context.sync();
                    }).catch(handleError);
                }
            };
        }
    }).catch(handleError);
}

// AI Feature: Fix Grammar
async function fixGrammar() {
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        
        const text = selection.text;
        if (!text || text.trim() === '') {
            document.getElementById('aiSuggestions').innerHTML = 
                '<div class="suggestion-item">Please select some text first</div>';
            return;
        }
        
        document.getElementById('aiSuggestions').innerHTML = 
            '<div class="suggestion-item">Fixing grammar and spelling...</div>';
        
        const prompt = `Fix grammar, spelling, and punctuation in this text:

${text}

Return ONLY the corrected text:`;
        
        const correctedText = await callOpenAI(prompt);
        
        if (correctedText) {
            document.getElementById('aiSuggestions').innerHTML = 
                `<div class="suggestion-item">Corrected text:<br><br>${correctedText.replace(/\n/g, '<br>')}</div>
                <button class="apply-button" id="applyCorrections">Apply Corrections</button>`;
                
            document.getElementById('applyCorrections').onclick = async () => {
                await Word.run(async (context) => {
                    const selection = context.document.getSelection();
                    selection.insertText(correctedText, 'Replace');
                    await context.sync();
                }).catch(handleError);
            };
        }
    }).catch(handleError);
}

// AI Feature: Summarize Selection
async function summarizeSelection() {
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        
        const text = selection.text;
        if (!text || text.trim() === '') {
            document.getElementById('aiSuggestions').innerHTML = 
                '<div class="suggestion-item">Please select some text first</div>';
            return;
        }
        
        document.getElementById('aiSuggestions').innerHTML = 
            '<div class="suggestion-item">Generating summary...</div>';
        
        const prompt = `Summarize this text in 2-3 concise sentences:

${text}

Return ONLY the summary:`;
        
        const summary = await callOpenAI(prompt);
        
        if (summary) {
            document.getElementById('aiSuggestions').innerHTML = 
                `<div class="suggestion-item">Summary:<br><br>${summary.replace(/\n/g, '<br>')}</div>
                <button class="apply-button" id="insertSummary">Replace with Summary</button>`;
                
            document.getElementById('insertSummary').onclick = async () => {
                await Word.run(async (context) => {
                    // Replace selected text with summary
                    const selection = context.document.getSelection();
                    selection.insertText(summary, 'Replace');
                    await context.sync();
                }).catch(handleError);
            };
        }
    }).catch(handleError);
}

// Live Suggestions Setup
let suggestionTimeout = null;
async function setupLiveSuggestions() {
    await Word.run(async (context) => {
        // Set up an event handler for document changes
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            onSelectionChanged,
            function(result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error('Failed to add selection changed handler:', result.error.message);
                }
            }
        );
    }).catch(handleError);
}

function onSelectionChanged(eventArgs) {
    // Clear any pending suggestion request
    if (suggestionTimeout) {
        clearTimeout(suggestionTimeout);
    }
    
    // Set a delay to avoid making too many API calls
    suggestionTimeout = setTimeout(async () => {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            
            // Only suggest if the selection is a reasonable size
            const text = selection.text;
            if (text && text.trim() !== '' && text.length < 500) {
                // Get quick suggestions without showing UI feedback
                const prompt = `Suggest one improvement for this text fragment: "${text}"

Return ONLY a single brief suggestion without any introduction or explanation.`;
                
                const quickSuggestion = await callOpenAI(prompt, 'gpt-3.5-turbo');
                
                if (quickSuggestion) {
                    document.getElementById('aiSuggestions').innerHTML = 
                        `<div class="suggestion-item">${quickSuggestion}</div>`;
                }
            }
        }).catch((error) => {
            // Silently handle errors for live suggestions
            console.error('Live suggestion error:', error);
        });
    }, 2000); // 2-second delay
}

// Error handling
function handleError(error) {
    console.error('Error:', error);
    document.getElementById('aiSuggestions').innerHTML = 
        `<div class="suggestion-item">Error: ${error.message}</div>`;
}
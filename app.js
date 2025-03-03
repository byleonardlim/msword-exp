// Handle URL parameters for context menu actions
function handleUrlParameters() {
    try {
        // Parse URL parameters
        const urlParams = new URLSearchParams(window.location.search);
        const action = urlParams.get('action');
        const text = urlParams.get('text');
        
        // If we have both action and text parameters
        if (action && text) {
            // Hide the main UI and show a processing message
            document.getElementById('container').style.display = 'none';
            
            // Create a simple UI for processing context menu actions
            const processingDiv = document.createElement('div');
            processingDiv.id = 'processingDiv';
            processingDiv.style.padding = '20px';
            processingDiv.style.textAlign = 'center';
            
            // Add processing message
            const message = document.createElement('p');
            message.innerText = `Processing ${action === 'fixGrammar' ? 'grammar fix' : 'summarization'}...`;
            processingDiv.appendChild(message);
            
            // Add to body
            document.body.appendChild(processingDiv);
            
            // Perform the requested action
            if (action === 'fixGrammar') {
                processGrammarFix(text);
            } else if (action === 'summarize') {
                processSummarization(text);
            }
        }
    } catch (error) {
        console.error('Error handling URL parameters:', error);
    }
}

// Process grammar fix from context menu
async function processGrammarFix(text) {
    try {
        const processingDiv = document.getElementById('processingDiv');
        processingDiv.innerHTML = '<p>Fixing grammar and spelling...</p>';
        
        const prompt = `Fix grammar, spelling, and punctuation in this text:

${text}

Return ONLY the corrected text:`;
        
        const correctedText = await callOpenAI(prompt);
        
        if (correctedText) {
            // Show the result
            processingDiv.innerHTML = `
                <p>Grammar fixed successfully:</p>
                <div style="text-align: left; background-color: #f5f5f5; padding: 10px; margin: 10px; border-left: 3px solid #0078d4;">
                    ${correctedText.replace(/\n/g, '<br>')}
                </div>
                <div style="margin-top: 15px;">
                    <button id="applyButton" style="background-color: #0078d4; color: white; border: none; padding: 8px 16px; cursor: pointer; margin-right: 10px;">Apply Changes</button>
                    <button id="cancelButton" style="background-color: #f3f2f1; border: 1px solid #8a8886; padding: 8px 16px; cursor: pointer;">Cancel</button>
                </div>
            `;
            
            // Add event handlers
            document.getElementById('applyButton').addEventListener('click', function() {
                // Send the corrected text back to the parent window
                Office.context.ui.messageParent(JSON.stringify({
                    correctedText: correctedText
                }));
            });
            
            document.getElementById('cancelButton').addEventListener('click', function() {
                Office.context.ui.messageParent('{"cancelled": true}');
            });
        } else {
            processingDiv.innerHTML = `
                <p>Error fixing grammar. Please try again.</p>
                <button id="closeButton" style="background-color: #f3f2f1; border: 1px solid #8a8886; padding: 8px 16px; cursor: pointer; margin-top: 10px;">Close</button>
            `;
            
            document.getElementById('closeButton').addEventListener('click', function() {
                Office.context.ui.messageParent('{"cancelled": true}');
            });
        }
    } catch (error) {
        console.error('Error processing grammar fix:', error);
    }
}

// Process summarization from context menu
async function processSummarization(text) {
    try {
        const processingDiv = document.getElementById('processingDiv');
        processingDiv.innerHTML = '<p>Generating summary...</p>';
        
        const prompt = `Summarize this text in 2-3 concise sentences:

${text}

Return ONLY the summary:`;
        
        const summary = await callOpenAI(prompt);
        
        if (summary) {
            // Show the result
            processingDiv.innerHTML = `
                <p>Summary generated successfully:</p>
                <div style="text-align: left; background-color: #f5f5f5; padding: 10px; margin: 10px; border-left: 3px solid #0078d4;">
                    ${summary.replace(/\n/g, '<br>')}
                </div>
                <div style="margin-top: 15px;">
                    <button id="applyButton" style="background-color: #0078d4; color: white; border: none; padding: 8px 16px; cursor: pointer; margin-right: 10px;">Replace with Summary</button>
                    <button id="cancelButton" style="background-color: #f3f2f1; border: 1px solid #8a8886; padding: 8px 16px; cursor: pointer;">Cancel</button>
                </div>
            `;
            
            // Add event handlers
            document.getElementById('applyButton').addEventListener('click', function() {
                // Send the summary text back to the parent window
                Office.context.ui.messageParent(JSON.stringify({
                    summaryText: summary
                }));
            });
            
            document.getElementById('cancelButton').addEventListener('click', function() {
                Office.context.ui.messageParent('{"cancelled": true}');
            });
        } else {
            processingDiv.innerHTML = `
                <p>Error generating summary. Please try again.</p>
                <button id="closeButton" style="background-color: #f3f2f1; border: 1px solid #8a8886; padding: 8px 16px; cursor: pointer; margin-top: 10px;">Close</button>
            `;
            
            document.getElementById('closeButton').addEventListener('click', function() {
                Office.context.ui.messageParent('{"cancelled": true}');
            });
        }
    } catch (error) {
        console.error('Error processing summarization:', error);
    }
}// Office.js initialization
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Initialize event handlers
        document.getElementById('saveApiKey').onclick = saveApiKey;
        document.getElementById('fixGrammar').onclick = fixGrammar;
        document.getElementById('summarizeSelection').onclick = summarizeSelection;
        
        // Load API key if saved
        loadApiKey();
        
        // Load document structure
        loadDocumentStructure();
        
        // Setup document change tracking
        setupDocumentChangeTracking();
        
        // Enable live suggestions as user types
        setupLiveSuggestions();
        
        // Add button to insert document canvas selection buttons
        addInsertSelectionButtonsButton();
        
        // Check if we need to perform an action based on URL parameters
        // (for context menu integration)
        handleUrlParameters();
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
    try {
        await Word.run(async (context) => {
            // Get all headings in the document
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load(['text', 'style', 'styleBuiltIn']);
            
            await context.sync();
            
            const headingsTree = document.getElementById('headingsTree');
            headingsTree.innerHTML = '';
            
            // Store heading info with their levels and indexes
            const headings = [];
            
            // Process paragraphs to find headings
            for (let i = 0; i < paragraphs.items.length; i++) {
                const paragraph = paragraphs.items[i];
                
                // Check if paragraph is a heading using styleBuiltIn property
                if (paragraph.styleBuiltIn && 
                    (paragraph.styleBuiltIn >= Word.Style.heading1 && 
                     paragraph.styleBuiltIn <= Word.Style.heading6)) {
                    
                    // Calculate heading level (1-6)
                    const level = paragraph.styleBuiltIn - Word.Style.heading1 + 1;
                    
                    headings.push({
                        index: i,
                        text: paragraph.text,
                        level: level,
                        isExpanded: true // Default to expanded
                    });
                }
            }
            
            // If no headings found, show a message
            if (headings.length === 0) {
                headingsTree.innerHTML = '<div class="no-headings">No headings found in document. Add headings to enable navigation.</div>';
                return;
            }
            
            // Build the heading tree UI
            for (let i = 0; i < headings.length; i++) {
                const heading = headings[i];
                
                // Create container for heading and toggle button
                const headingContainer = document.createElement('div');
                headingContainer.className = 'heading-container';
                
                // Add toggle button
                const toggleButton = document.createElement('button');
                toggleButton.className = 'toggle-button';
                toggleButton.innerHTML = heading.isExpanded ? '−' : '+'; // Minus or plus sign
                toggleButton.title = heading.isExpanded ? 'Collapse section' : 'Expand section';
                toggleButton.addEventListener('click', (e) => {
                    e.stopPropagation(); // Prevent triggering the heading click
                    toggleSection(heading.index, heading.level, headings, i);
                });
                
                // Create heading item
                const headingItem = document.createElement('div');
                headingItem.className = `heading-item heading-h${heading.level}`;
                headingItem.innerText = heading.text || '[Empty Heading]';
                
                // Store paragraph index as data attribute
                headingItem.dataset.paragraphIndex = heading.index;
                headingItem.dataset.headingIndex = i;
                
                // Add click handler to select content
                headingItem.addEventListener('click', () => {
                    navigateToHeading(heading.index, heading.level, headings);
                });
                
                // Add elements to container
                headingContainer.appendChild(toggleButton);
                headingContainer.appendChild(headingItem);
                headingsTree.appendChild(headingContainer);
            }
            
            // Add refresh button
            const refreshButton = document.createElement('button');
            refreshButton.className = 'ms-Button';
            refreshButton.innerHTML = '<span class="ms-Button-label">Refresh Structure</span>';
            refreshButton.onclick = loadDocumentStructure;
            
            // Add the refresh button to the top of the headings container
            headingsTree.parentElement.insertBefore(refreshButton, headingsTree);
        });
    } catch (error) {
        console.error('Error loading document structure:', error);
        document.getElementById('headingsTree').innerHTML = 
            `<div class="error">Error loading document structure: ${error.message}</div>`;
    }
}

// Toggle section visibility (expand/collapse)
async function toggleSection(paragraphIndex, headingLevel, headings, headingArrayIndex) {
    try {
        // Toggle the expanded state
        const isExpanded = !headings[headingArrayIndex].isExpanded;
        headings[headingArrayIndex].isExpanded = isExpanded;
        
        // Update the toggle button
        const toggleButton = document.querySelector(`.heading-container:nth-child(${headingArrayIndex + 1}) .toggle-button`);
        if (toggleButton) {
            toggleButton.innerHTML = isExpanded ? '−' : '+';
            toggleButton.title = isExpanded ? 'Collapse section' : 'Expand section';
        }
        
        await Word.run(async (context) => {
            // Get all paragraphs
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load(['text', 'style', 'styleBuiltIn']);
            
            await context.sync();
            
            // Find the section's end index
            let endIndex = paragraphs.items.length - 1;
            
            // Find the next heading of same or higher level
            for (let i = headingArrayIndex + 1; i < headings.length; i++) {
                if (headings[i].level <= headingLevel) {
                    endIndex = headings[i].index - 1;
                    break;
                }
            }
            
            // Get the paragraphs to show/hide (exclude the heading itself)
            const startIndex = paragraphIndex + 1;
            
            if (startIndex <= endIndex) {
                // Style to use for hiding/showing
                const hiddenStyle = context.document.styles.add("HiddenContent");
                hiddenStyle.font.color = "white"; // Make text white (invisible on white background)
                hiddenStyle.font.size = 1; // Very small size
                hiddenStyle.hidden = true; // Use Word's hidden text feature
                
                // Apply or remove the style to each paragraph in the section
                for (let i = startIndex; i <= endIndex; i++) {
                    if (i < paragraphs.items.length) {
                        const paragraph = paragraphs.items[i];
                        
                        // Skip headings within this section (we don't want to hide them)
                        const isHeading = paragraph.styleBuiltIn && 
                                         (paragraph.styleBuiltIn >= Word.Style.heading1 && 
                                          paragraph.styleBuiltIn <= Word.Style.heading6);
                        
                        if (!isHeading) {
                            if (!isExpanded) {
                                // Hide paragraph
                                paragraph.style = "HiddenContent";
                            } else {
                                // Restore paragraph (use Normal style if no specific style)
                                paragraph.style = "Normal";
                            }
                        }
                    }
                }
            }
            
            await context.sync();
        });
    } catch (error) {
        console.error('Error toggling section:', error);
    }
}

async function navigateToHeading(paragraphIndex, headingLevel, headings) {
    await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('text');
        
        await context.sync();
        
        if (paragraphIndex < paragraphs.items.length) {
            // Find the ending paragraph
            let endIndex = -1;
            
            // Find the next heading of same or higher level
            for (const heading of headings) {
                if (heading.index > paragraphIndex && heading.level <= headingLevel) {
                    endIndex = heading.index - 1;
                    break;
                }
            }
            
            // If no next heading found, select to the end of the document
            if (endIndex === -1) {
                endIndex = paragraphs.items.length - 1;
            }
            
            // Select the range from heading to next heading (or end)
            const startParagraph = paragraphs.items[paragraphIndex];
            startParagraph.select();
            
            // If we need to select more than one paragraph
            if (endIndex > paragraphIndex) {
                // Get the range of the first paragraph
                const range = startParagraph.getRange();
                
                // Get the range of the last paragraph
                const endParagraph = paragraphs.items[endIndex];
                const endRange = endParagraph.getRange('End');
                
                // Expand the selection from start to end
                range.expandTo(endRange);
                range.select();
            }
            
            // Scroll the selected content into view
            startParagraph.getRange().scrollIntoView();
            
            // Update UI to highlight selected heading
            updateSelectedHeadingUI(paragraphIndex);
        }
        
        await context.sync();
    }).catch(handleError);
}

// Alternative approach using content controls
async function toggleSectionWithContentControls(paragraphIndex, headingLevel, headings, headingArrayIndex) {
    try {
        // Toggle the expanded state
        const isExpanded = !headings[headingArrayIndex].isExpanded;
        headings[headingArrayIndex].isExpanded = isExpanded;
        
        // Update the toggle button
        const toggleButton = document.querySelector(`.heading-container:nth-child(${headingArrayIndex + 1}) .toggle-button`);
        if (toggleButton) {
            toggleButton.innerHTML = isExpanded ? '−' : '+';
            toggleButton.title = isExpanded ? 'Collapse section' : 'Expand section';
        }
        
        await Word.run(async (context) => {
            // Get the section content control if it exists
            const contentControls = context.document.contentControls;
            contentControls.load("items");
            await context.sync();
            
            // Look for existing content control for this section
            let sectionControl = null;
            const controlTag = `section-${paragraphIndex}`;
            
            for (let i = 0; i < contentControls.items.length; i++) {
                if (contentControls.items[i].tag === controlTag) {
                    sectionControl = contentControls.items[i];
                    break;
                }
            }
            
            // If we don't have a content control for this section yet, create one
            if (!sectionControl && !isExpanded) {
                // Get all paragraphs
                const paragraphs = context.document.body.paragraphs;
                paragraphs.load(['text', 'style', 'styleBuiltIn']);
                
                await context.sync();
                
                // Find the section's end index
                let endIndex = paragraphs.items.length - 1;
                
                // Find the next heading of same or higher level
                for (let i = headingArrayIndex + 1; i < headings.length; i++) {
                    if (headings[i].level <= headingLevel) {
                        endIndex = headings[i].index - 1;
                        break;
                    }
                }
                
                // Get the paragraphs to wrap (exclude the heading itself)
                const startIndex = paragraphIndex + 1;
                
                if (startIndex <= endIndex) {
                    // Create a range from start to end paragraphs
                    const startParagraph = paragraphs.items[startIndex];
                    const endParagraph = paragraphs.items[endIndex];
                    
                    const startRange = startParagraph.getRange("Start");
                    const endRange = endParagraph.getRange("End");
                    
                    startRange.expandTo(endRange);
                    
                    // Insert a content control around the section
                    sectionControl = startRange.insertContentControl();
                    sectionControl.tag = controlTag;
                    sectionControl.title = "Section Content";
                    sectionControl.appearance = "BoundingBox";
                    
                    // Collapse the control
                    sectionControl.appearance = "Hidden";
                }
            } 
            // If we have a control and need to expand it
            else if (sectionControl && isExpanded) {
                // Show the content control
                sectionControl.appearance = "BoundingBox";
                
                // Optional: Delete the content control but keep its contents
                sectionControl.delete(true);
            }
            
            await context.sync();
        });
    } catch (error) {
        console.error('Error toggling section with content controls:', error);
    }
}

// Update the UI to highlight the selected heading
function updateSelectedHeadingUI(selectedIndex) {
    // Remove existing selection highlight
    const headingItems = document.querySelectorAll('.heading-item');
    headingItems.forEach(item => {
        item.classList.remove('selected');
    });
    
    // Add highlight to the selected heading
    const selectedItem = document.querySelector(`.heading-item[data-paragraphIndex="${selectedIndex}"]`);
    if (selectedItem) {
        selectedItem.classList.add('selected');
        selectedItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}

// Setup document change tracking to refresh structure
function setupDocumentChangeTracking() {
    // Debounce function to limit updates
    function debounce(func, wait) {
        let timeout;
        return function(...args) {
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(this, args), wait);
        };
    }
    
    // Listen for document changes that might affect headings
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        debounce(() => {
            // Only refresh occasionally to avoid performance issues
            if (Math.random() < 0.1) { // 10% chance to refresh
                loadDocumentStructure();
            }
        }, 2000)
    );
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
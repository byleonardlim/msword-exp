Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        // Initialize event handlers
        document.getElementById("insert-content-control").onclick = insertContentControlAroundSelection;
        document.getElementById("add-critique").onclick = addCritiqueToSelection;
        document.getElementById("wrap-paragraphs").onclick = wrapParagraphsInContentControls;
        document.getElementById("highlight-controls").onclick = highlightAllContentControls;
        document.getElementById("list-controls").onclick = listAllContentControls;
        
        // Set up document selection change event
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, updateSelectedText);
        
        // Initial critique count
        window.critiqueCount = 0;
        window.critiques = [];
        
        // Initial call to update the task pane
        updateSelectedText();
        updateCritiqueSummary();
    }
});

// Track selected text and display it in the task pane with enhanced highlighting
function updateSelectedText() {
    Word.run(async function(context) {
        // Get the current selection
        const selection = context.document.getSelection();
        selection.load("text");
        
        await context.sync();
        
        // Update the task pane with the selected text
        const selectedTextElement = document.getElementById("selected-text");
        if (selection.text.trim() === "") {
            selectedTextElement.innerHTML = "<p>No text currently selected. Highlight text in the document to see it here.</p>";
        } else {
            // Truncate very long selections for display
            let displayText = selection.text;
            if (displayText.length > 300) {
                displayText = displayText.substring(0, 300) + "...";
            }
            
            // Show highlighted text with prominent styling and edit button
            selectedTextElement.innerHTML = `
                <div class="highlighted-container">
                    <p class="selected-content">${escapeHtml(displayText)}</p>
                    <div class="highlight-actions">
                        <button id="quick-wrap-button" class="highlight-action-button">Wrap for Editing</button>
                    </div>
                </div>
            `;
            
            // Add event listener to the quick wrap button
            document.getElementById("quick-wrap-button").addEventListener("click", function() {
                wrapSelectionForEditing();
            });
        }
    }).catch(function(error) {
        console.log("Error: " + error);
        selectedTextElement.innerHTML = `<p class='error'>Error displaying selection: ${error.message}</p>`;
    });
}

// Helper function to safely display text in HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Insert a content control around the current selection
function insertContentControlAroundSelection() {
    Word.run(async function(context) {
        // Get the current selection
        const range = context.document.getSelection();
        range.load("text");
        
        await context.sync();
        
        // Check if there is text selected
        if (range.text.trim() === "") {
            document.getElementById("selected-text").innerHTML = 
                "<p class='error'>Please select some text first.</p>";
            return;
        }
        
        // Create a content control around the selection
        const contentControl = range.insertContentControl();
        
        // Configure the content control
        contentControl.title = "Paragraph Control";
        contentControl.tag = "paragraph-" + new Date().getTime(); // Unique tag
        contentControl.appearance = Word.ContentControlAppearance.boundingBox;
        contentControl.color = "blue"; // Set the color of the content control
        
        await context.sync();
        
        document.getElementById("selected-text").innerHTML = 
            `<p class='success'>Content control added successfully around: "${range.text.substring(0, 50)}${range.text.length > 50 ? '...' : ''}"</p>`;
            
        // Update the control list
        listAllContentControls();
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error adding content control: ${error.message}</p>`;
    });
}

// Wrap the current selection for editing with better visual feedback
function wrapSelectionForEditing() {
    Word.run(async function(context) {
        // Get the current selection
        const range = context.document.getSelection();
        range.load("text");
        
        await context.sync();
        
        // Check if there is text selected
        if (range.text.trim() === "") {
            document.getElementById("selected-text").innerHTML = 
                "<p class='error'>Please select some text first.</p>";
            return;
        }
        
        // Create a content control around the selection specifically for editing
        const contentControl = range.insertContentControl();
        
        // Configure the content control with edit-specific properties
        contentControl.title = "Edit Selection";
        contentControl.tag = "edit-" + new Date().getTime(); // Unique tag
        contentControl.appearance = Word.ContentControlAppearance.boundingBox;
        contentControl.color = "purple"; // Purple to indicate editing mode
        
        // Highlight the text for better visibility during editing
        const contentRange = contentControl.getRange();
        contentRange.font.highlightColor = "#E6E6FA"; // Light purple highlight
        
        await context.sync();
        
        // Create a feedback message with guidance for the user
        const textPreview = range.text.substring(0, 50) + (range.text.length > 50 ? '...' : '');
        
        document.getElementById("selected-text").innerHTML = `
            <div class="edit-feedback">
                <p class='success'>Text wrapped and ready for editing!</p>
                <p class="edit-preview">"${escapeHtml(textPreview)}"</p>
                <p class="edit-instruction">You can now edit this text directly in the document. The purple box indicates the editable region.</p>
            </div>
        `;
            
        // Update the control list to show the new editable content
        listAllContentControls();
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error preparing text for editing: ${error.message}</p>`;
    });
}

// Add a critique to the selected text
function addCritiqueToSelection() {
    Word.run(async function(context) {
        // Get the current selection
        const range = context.document.getSelection();
        range.load("text");
        
        await context.sync();
        
        // Check if there is text selected
        if (range.text.trim() === "") {
            document.getElementById("selected-text").innerHTML = 
                "<p class='error'>Please select some text first.</p>";
            return;
        }
        
        // Create a content control around the selection
        const contentControl = range.insertContentControl();
        
        // Increment critique count
        window.critiqueCount++;
        const critiqueId = window.critiqueCount;
        
        // Configure the content control
        contentControl.title = `Critique #${critiqueId}`;
        contentControl.tag = `critique-${critiqueId}`;
        contentControl.appearance = Word.ContentControlAppearance.tags;
        contentControl.color = "red"; // Set the color of the content control
        
        // Store the critique
        window.critiques.push({
            id: critiqueId,
            text: range.text,
            timestamp: new Date().toLocaleString()
        });
        
        await context.sync();
        
        // Update UI
        document.getElementById("selected-text").innerHTML = 
            `<p class='success'>Critique #${critiqueId} added for: "${range.text.substring(0, 50)}${range.text.length > 50 ? '...' : ''}"</p>`;
            
        // Update critique summary and control list
        updateCritiqueSummary();
        listAllContentControls();
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error adding critique: ${error.message}</p>`;
    });
}

// Update the critique summary in the task pane
function updateCritiqueSummary() {
    const summaryElement = document.getElementById("critique-summary");
    
    if (window.critiques.length === 0) {
        summaryElement.innerHTML = "<p>No critiques added yet.</p>";
        return;
    }
    
    let html = `<p class="critique-count">Total critiques: ${window.critiques.length}</p>`;
    html += "<ul class='critique-list'>";
    
    window.critiques.forEach(critique => {
        let displayText = critique.text;
        if (displayText.length > 100) {
            displayText = displayText.substring(0, 100) + "...";
        }
        
        html += `<li class="critique-item">
            <div class="critique-header">
                <span class="critique-id">Critique #${critique.id}</span>
                <span class="critique-timestamp">${critique.timestamp}</span>
            </div>
            <div class="critique-content">${escapeHtml(displayText)}</div>
            <button class="goto-critique" data-id="${critique.id}">Go to this critique</button>
        </li>`;
    });
    
    html += "</ul>";
    summaryElement.innerHTML = html;
    
    // Add event listeners to "Go to" buttons
    const gotoButtons = document.getElementsByClassName("goto-critique");
    for (let i = 0; i < gotoButtons.length; i++) {
        gotoButtons[i].addEventListener("click", function() {
            navigateToCritique(this.getAttribute("data-id"));
        });
    }
}

// Navigate to a specific critique in the document
function navigateToCritique(id) {
    Word.run(async function(context) {
        // Get all content controls
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        contentControls.load("items/tag");
        
        await context.sync();
        
        // Find the critique content control with the matching tag
        const tag = `critique-${id}`;
        let found = false;
        
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Select this content control
                contentControls.items[i].select();
                found = true;
                break;
            }
        }
        
        await context.sync();
        
        if (!found) {
            document.getElementById("selected-text").innerHTML = 
                `<p class='error'>Critique #${id} not found. It may have been removed.</p>`;
            
            // Remove from our list if not found in document
            window.critiques = window.critiques.filter(critique => critique.id != id);
            updateCritiqueSummary();
        }
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error navigating to critique: ${error.message}</p>`;
    });
}

// Wrap each paragraph in the document with a content control
function wrapParagraphsInContentControls() {
    Word.run(async function(context) {
        // Get all paragraphs in the document
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        
        await context.sync();
        
        let count = 0;
        
        // Loop through each paragraph and wrap it in a content control
        for (let i = 0; i < paragraphs.items.length; i++) {
            if (paragraphs.items[i].text.trim() !== "") {
                const contentControl = paragraphs.items[i].insertContentControl();
                contentControl.title = "Paragraph " + (i + 1);
                contentControl.tag = "para-" + i;
                contentControl.appearance = Word.ContentControlAppearance.boundingBox;
                contentControl.color = "green";
                count++;
            }
        }
        
        await context.sync();
        
        document.getElementById("selected-text").innerHTML = 
            `<p class='success'>${count} paragraphs wrapped in content controls.</p>`;
            
        // Update the control list
        listAllContentControls();
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error wrapping paragraphs: ${error.message}</p>`;
    });
}

// Highlight all content controls in the document
function highlightAllContentControls() {
    Word.run(async function(context) {
        // Get all content controls in the document
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        
        await context.sync();
        
        // Loop through each content control and highlight it
        for (let i = 0; i < contentControls.items.length; i++) {
            const contentControl = contentControls.items[i];
            
            // Get the range of the content control
            const range = contentControl.getRange();
            
            // Apply highlighting
            range.font.highlightColor = "#FFFF00"; // Yellow highlighting
        }
        
        await context.sync();
        
        document.getElementById("selected-text").innerHTML = 
            `<p class='success'>${contentControls.items.length} content controls highlighted.</p>`;
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error highlighting content controls: ${error.message}</p>`;
    });
}

// List all content controls in the document
function listAllContentControls() {
    Word.run(async function(context) {
        // Get all content controls in the document
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        contentControls.load("items/tag");
        contentControls.load("items/title");
        contentControls.load("items/text");
        contentControls.load("items/appearance");
        contentControls.load("items/color");
        
        await context.sync();
        
        // Clear the previous list
        document.getElementById("control-list").innerHTML = "";
        
        // Create a list of content controls
        const listElement = document.createElement("ul");
        listElement.className = "control-list";
        
        if (contentControls.items.length === 0) {
            document.getElementById("control-list").innerHTML = "<p>No content controls found.</p>";
        } else {
            // Group content controls by type
            const paragraphControls = [];
            const critiqueControls = [];
            const otherControls = [];
            
            for (let i = 0; i < contentControls.items.length; i++) {
                const contentControl = contentControls.items[i];
                
                if (contentControl.tag && contentControl.tag.startsWith("critique-")) {
                    critiqueControls.push(contentControl);
                } else if (contentControl.tag && contentControl.tag.startsWith("para-")) {
                    paragraphControls.push(contentControl);
                } else {
                    otherControls.push(contentControl);
                }
            }
            
            // Function to create list items for each control group
            const createListItems = (controls, groupTitle) => {
                if (controls.length === 0) return "";
                
                let html = `<li class="control-group"><h3>${groupTitle} (${controls.length})</h3><ul class="control-sublist">`;
                
                for (const control of controls) {
                    // Truncate long text for display
                    const displayText = control.text.length > 50 
                        ? control.text.substring(0, 50) + "..." 
                        : control.text;
                    
                    html += `<li class="control-item">
                        <div class="control-header">
                            <span class="control-title">${control.title || "No title"}</span>
                            <span class="control-tag">${control.tag || "No tag"}</span>
                        </div>
                        <div class="control-text">${escapeHtml(displayText)}</div>
                        <div class="control-actions">
                            <button class="goto-control" data-tag="${control.tag}">Go to</button>
                            <button class="remove-control" data-tag="${control.tag}">Remove</button>
                            <button class="edit-control" data-tag="${control.tag}">Edit</button>
                        </div>
                    </li>`;
                }
                
                html += "</ul></li>";
                return html;
            };
            
            // Add each group to the list
            listElement.innerHTML = 
                createListItems(critiqueControls, "Critiques") +
                createListItems(paragraphControls, "Paragraphs") +
                createListItems(otherControls, "Other Controls");
            
            document.getElementById("control-list").appendChild(listElement);
            
            // Add event listeners to buttons
            const addButtonListeners = (className, handler) => {
                const buttons = document.getElementsByClassName(className);
                for (let i = 0; i < buttons.length; i++) {
                    buttons[i].addEventListener("click", function() {
                        handler(this.getAttribute("data-tag"));
                    });
                }
            };
            
            addButtonListeners("goto-control", gotoContentControl);
            addButtonListeners("remove-control", removeContentControl);
            addButtonListeners("edit-control", editContentControl);
        }
        
        await context.sync();
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("control-list").innerHTML = 
            `<p class='error'>Error listing content controls: ${error.message}</p>`;
    });
}

// Navigate to a content control
function gotoContentControl(tag) {
    Word.run(async function(context) {
        // Get all content controls
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        contentControls.load("items/tag");
        
        await context.sync();
        
        // Find the content control with the matching tag
        let found = false;
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Select this content control
                contentControls.items[i].select();
                found = true;
                break;
            }
        }
        
        await context.sync();
        
        if (!found) {
            document.getElementById("selected-text").innerHTML = 
                `<p class='error'>Control with tag "${tag}" not found. It may have been removed.</p>`;
            
            // If it was a critique that's not found, remove it from our list
            if (tag.startsWith("critique-")) {
                const id = tag.replace("critique-", "");
                window.critiques = window.critiques.filter(critique => critique.id != id);
                updateCritiqueSummary();
            }
            
            // Update the control list
            listAllContentControls();
        }
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error navigating to content control: ${error.message}</p>`;
    });
}

// Remove a content control by tag
function removeContentControl(tag) {
    Word.run(async function(context) {
        // Get all content controls in the document
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        contentControls.load("items/tag");
        
        await context.sync();
        
        // Find the content control with the matching tag
        let found = false;
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Remove the content control but keep its content
                contentControls.items[i].delete(true);
                found = true;
                break;
            }
        }
        
        await context.sync();
        
        if (found) {
            document.getElementById("selected-text").innerHTML = 
                `<p class='success'>Content control with tag "${tag}" removed.</p>`;
                
            // If it was a critique that was removed, update our list
            if (tag.startsWith("critique-")) {
                const id = tag.replace("critique-", "");
                window.critiques = window.critiques.filter(critique => critique.id != id);
                updateCritiqueSummary();
            }
            
            // Update the control list
            listAllContentControls();
        } else {
            document.getElementById("selected-text").innerHTML = 
                `<p class='error'>Control with tag "${tag}" not found.</p>`;
        }
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error removing content control: ${error.message}</p>`;
    });
}

// Edit a content control by tag
function editContentControl(tag) {
    Word.run(async function(context) {
        // Get all content controls in the document
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        contentControls.load("items/tag");
        
        await context.sync();
        
        // Find the content control with the matching tag
        let found = false;
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Get the content control
                const contentControl = contentControls.items[i];
                
                // Select it for editing
                contentControl.select();
                
                // Customize the content control appearance to indicate editing mode
                contentControl.appearance = Word.ContentControlAppearance.boundingBox;
                contentControl.color = "purple";
                
                found = true;
                break;
            }
        }
        
        await context.sync();
        
        if (found) {
            document.getElementById("selected-text").innerHTML = 
                `<p class='success'>Now editing content control with tag "${tag}". Make your changes in the document.</p>`;
        } else {
            document.getElementById("selected-text").innerHTML = 
                `<p class='error'>Control with tag "${tag}" not found.</p>`;
        }
    }).catch(function(error) {
        console.log("Error: " + error);
        document.getElementById("selected-text").innerHTML = 
            `<p class='error'>Error editing content control: ${error.message}</p>`;
    });
}

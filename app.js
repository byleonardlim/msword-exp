Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        // Initialize event handlers
        document.getElementById("insert-content-control").onclick = insertContentControlAroundSelection;
        document.getElementById("wrap-paragraphs").onclick = wrapParagraphsInContentControls;
        document.getElementById("highlight-controls").onclick = highlightAllContentControls;
        document.getElementById("list-controls").onclick = listAllContentControls;
    }
});

// Insert a content control around the current selection
function insertContentControlAroundSelection() {
    Word.run(async function(context) {
        // Get the current selection
        const range = context.document.getSelection();
        
        // Create a content control around the selection
        const contentControl = range.insertContentControl();
        
        // Configure the content control
        contentControl.title = "Paragraph Control";
        contentControl.tag = "paragraph-" + new Date().getTime(); // Unique tag
        contentControl.appearance = Word.ContentControlAppearance.tags;
        contentControl.color = "blue"; // Set the color of the content control
        
        await context.sync();
        console.log("Content control inserted successfully.");
    }).catch(function(error) {
        console.log("Error: " + error);
    });
}

// Wrap each paragraph in the document with a content control
function wrapParagraphsInContentControls() {
    Word.run(async function(context) {
        // Get all paragraphs in the document
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        
        await context.sync();
        
        // Loop through each paragraph and wrap it in a content control
        for (let i = 0; i < paragraphs.items.length; i++) {
            if (paragraphs.items[i].text.trim() !== "") {
                const contentControl = paragraphs.items[i].insertContentControl();
                contentControl.title = "Paragraph " + (i + 1);
                contentControl.tag = "para-" + i;
                contentControl.appearance = Word.ContentControlAppearance.boundingBox;
                contentControl.color = "green";
            }
        }
        
        await context.sync();
        console.log("All paragraphs wrapped in content controls.");
    }).catch(function(error) {
        console.log("Error: " + error);
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
        console.log("All content controls highlighted.");
    }).catch(function(error) {
        console.log("Error: " + error);
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
        
        await context.sync();
        
        // Clear the previous list
        document.getElementById("control-list").innerHTML = "";
        
        // Create a list of content controls
        const listElement = document.createElement("ul");
        
        if (contentControls.items.length === 0) {
            document.getElementById("control-list").innerHTML = "<p>No content controls found.</p>";
        } else {
            for (let i = 0; i < contentControls.items.length; i++) {
                const contentControl = contentControls.items[i];
                
                // Create a list item for each content control
                const listItem = document.createElement("li");
                listItem.innerHTML = `
                    <strong>Title:</strong> ${contentControl.title || "No title"} | 
                    <strong>Tag:</strong> ${contentControl.tag || "No tag"} | 
                    <strong>Text:</strong> ${contentControl.text.substring(0, 30)}${contentControl.text.length > 30 ? "..." : ""}
                    <button class="remove-control" data-tag="${contentControl.tag}">Remove</button>
                    <button class="edit-control" data-tag="${contentControl.tag}">Edit</button>
                `;
                listElement.appendChild(listItem);
            }
            
            document.getElementById("control-list").appendChild(listElement);
            
            // Add event listeners to the remove and edit buttons
            const removeButtons = document.getElementsByClassName("remove-control");
            for (let i = 0; i < removeButtons.length; i++) {
                removeButtons[i].addEventListener("click", function() {
                    removeContentControl(this.getAttribute("data-tag"));
                });
            }
            
            const editButtons = document.getElementsByClassName("edit-control");
            for (let i = 0; i < editButtons.length; i++) {
                editButtons[i].addEventListener("click", function() {
                    editContentControl(this.getAttribute("data-tag"));
                });
            }
        }
        
        await context.sync();
    }).catch(function(error) {
        console.log("Error: " + error);
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
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Remove the content control but keep its content
                contentControls.items[i].delete(true);
                break;
            }
        }
        
        await context.sync();
        
        // Refresh the list of content controls
        listAllContentControls();
    }).catch(function(error) {
        console.log("Error: " + error);
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
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Get the content control
                const contentControl = contentControls.items[i];
                
                // Customize the content control appearance
                contentControl.appearance = Word.ContentControlAppearance.tags;
                contentControl.color = "red";
                
                // Set focus to this content control
                contentControl.select();
                break;
            }
        }
        
        await context.sync();
    }).catch(function(error) {
        console.log("Error: " + error);
    });
}

// Add a new paragraph with a content control
function addParagraphWithContentControl() {
    Word.run(async function(context) {
        // Add a new paragraph at the end of the document
        const paragraph = context.document.body.insertParagraph("New paragraph with content control", Word.InsertLocation.end);
        
        // Create a content control around the paragraph
        const contentControl = paragraph.insertContentControl();
        contentControl.title = "New Paragraph";
        contentControl.tag = "new-para-" + new Date().getTime();
        contentControl.appearance = Word.ContentControlAppearance.boundingBox;
        contentControl.color = "purple";
        
        // Make the content control placeholder text editable
        contentControl.placeholderText = "Edit this text...";
        
        await context.sync();
        console.log("New paragraph with content control added.");
    }).catch(function(error) {
        console.log("Error: " + error);
    });
}

// Advanced: Lock a content control to prevent editing
function lockContentControl(tag) {
    Word.run(async function(context) {
        // Get all content controls
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        contentControls.load("items/tag");
        
        await context.sync();
        
        // Find the content control with the matching tag
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Get the content control
                const contentControl = contentControls.items[i];
                
                // Lock the content control
                contentControl.cannotEdit = true;
                contentControl.cannotDelete = true;
                
                // Change appearance to indicate it's locked
                contentControl.appearance = Word.ContentControlAppearance.tags;
                contentControl.color = "gray";
                break;
            }
        }
        
        await context.sync();
        console.log("Content control locked.");
    }).catch(function(error) {
        console.log("Error: " + error);
    });
}

// Advanced: Change text formatting within a content control
function formatContentControlText(tag, formatting) {
    Word.run(async function(context) {
        // Get all content controls
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        contentControls.load("items/tag");
        
        await context.sync();
        
        // Find the content control with the matching tag
        for (let i = 0; i < contentControls.items.length; i++) {
            if (contentControls.items[i].tag === tag) {
                // Get the range of the content control
                const range = contentControls.items[i].getRange();
                
                // Apply formatting
                if (formatting.bold) range.font.bold = true;
                if (formatting.italic) range.font.italic = true;
                if (formatting.color) range.font.color = formatting.color;
                if (formatting.size) range.font.size = formatting.size;
                
                break;
            }
        }
        
        await context.sync();
        console.log("Content control text formatted.");
    }).catch(function(error) {
        console.log("Error: " + error);
    });
}
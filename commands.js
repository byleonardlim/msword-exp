// Commands for context menu and ribbon buttons
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        // Initialize event handlers
    };

    // Expose functions to the global Office object
    Office.actions.associate("fixGrammarContextMenu", fixGrammarContextMenu);
    Office.actions.associate("summarizeContextMenu", summarizeContextMenu);

    // Context menu function for Fix Grammar
    function fixGrammarContextMenu(event) {
        // Get the selected text
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const selectedText = asyncResult.value;
                
                // If there's selected text, process it
                if (selectedText && selectedText.trim() !== '') {
                    // Show the task pane
                    Office.context.ui.displayTaskpaneAsync().then(function() {
                        // Use the same AI action as in the task pane
                        Office.actions.invoke("FIX_GRAMMAR", {
                            text: selectedText
                        }).then(function(result) {
                            if (result.correctedText) {
                                // Replace the selected text with the corrected version
                                Office.context.document.setSelectedDataAsync(
                                    result.correctedText,
                                    { coercionType: Office.CoercionType.Text },
                                    function (result) {
                                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                                            console.log('Text replaced successfully');
                                        } else {
                                            console.error('Error replacing text:', result.error);
                                        }
                                    }
                                );
                            }
                        }).catch(function(error) {
                            console.error('Error processing grammar fix:', error);
                        });
                    });
                } else {
                    // Notify if no text is selected
                    Office.context.ui.displayDialogAsync(
                        'https://localhost:3000/notification.html?message=Please%20select%20some%20text%20first',
                        { height: 30, width: 40, displayInIframe: true },
                        function() {}
                    );
                }
            }
        });
        
        event.completed();
    }

    // Context menu function for Summarize
    function summarizeContextMenu(event) {
        // Get the selected text
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const selectedText = asyncResult.value;
                
                // If there's selected text, process it
                if (selectedText && selectedText.trim() !== '') {
                    // Show the task pane
                    Office.context.ui.displayTaskpaneAsync().then(function() {
                        // Send message to taskpane to process the summarization
                        Office.actions.invoke("SUMMARIZE_TEXT", {
                            text: selectedText
                        }).then(function(result) {
                            if (result.summaryText) {
                                // Replace the selected text with the summary
                                Office.context.document.setSelectedDataAsync(
                                    result.summaryText,
                                    { coercionType: Office.CoercionType.Text },
                                    function (result) {
                                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                                            console.log('Text replaced with summary successfully');
                                        } else {
                                            console.error('Error replacing text:', result.error);
                                        }
                                    }
                                );
                            }
                        }).catch(function(error) {
                            console.error('Error processing summary:', error);
                        });
                    });
                } else {
                    // Notify if no text is selected
                    Office.context.ui.displayDialogAsync(
                        'https://localhost:3000/notification.html?message=Please%20select%20some%20text%20first',
                        { height: 30, width: 40, displayInIframe: true },
                        function() {}
                    );
                }
            }
        });
        
        event.completed();
    }
})();
// Configuration file for external API connections
// Uses local storage to retrieve API keys for development/prototype purposes

// Load API configuration from local storage
function getApiConfiguration() {
    return {
        // OpenAI API Configuration
        openai: {
            apiKey: localStorage.getItem('openai_api_key') || '',
            model: localStorage.getItem('openai_model') || 'gpt-4o',
            endpoint: "https://api.openai.com/v1/chat/completions",
            maxTokens: 2000,
            temperature: 0.3
        },
        
        // Microsoft Graph API Configuration (if needed for future expansion)
        graph: {
            clientId: localStorage.getItem('graph_client_id') || '',
            authority: "https://login.microsoftonline.com/common",
            redirectUri: "https://byleonardlim.github.io/msword-exp/taskpane.html",
            scopes: ["Files.ReadWrite"]
        }
    };
}

// Function to check if APIs are configured
function isApiConfigured() {
    const config = getApiConfiguration();
    return config.openai.apiKey !== '';
}

// Function to open settings dialog
function openSettingsDialog() {
    // Display a dialog for setting API keys
    Office.context.ui.displayDialogAsync(
        'https://byleonardlim.github.io/msword-exp/settings.html',
        { height: 60, width: 30, displayInIframe: true },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Failed to open settings dialog: ' + result.error.message);
            }
        }
    );
}

// Export configurations
Office.onReady(() => {
    // Make configurations available globally
    window.getApiConfiguration = getApiConfiguration;
    window.isApiConfigured = isApiConfigured;
    window.openSettingsDialog = openSettingsDialog;
});
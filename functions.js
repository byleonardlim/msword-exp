// Functions.js - Contains the AI integration functionality

/**
 * Analyzes document content and returns spelling corrections and tone suggestions
 * @param {string} text - The document text to analyze
 * @returns {Promise<object>} - Analysis results with spelling and tone suggestions
 */
async function analyzeDocumentContent(text) {
    try {
        // Check if API is configured
        if (!window.isApiConfigured()) {
            console.log("API not configured, returning mock data");
            return getMockAnalysisResults(text);
        }
        
        // Get API configuration from localStorage
        const config = window.getApiConfiguration().openai;
        
        // Prepare the AI prompt
        const systemPrompt = `
You are an AI assistant that helps with document editing in Microsoft Word. 
Analyze the provided text for:
1. Spelling errors: Identify misspelled words and suggest corrections
2. Tone suggestions: Convert formal language to smart casual tone

Format your response as a JSON object with this exact structure:
{
  "spelling": [
    { "word": "misspelled word", "suggestion": "correct spelling" },
    ...
  ],
  "tone": [
    { "original": "original formal text", "suggestion": "suggested casual text" },
    ...
  ]
}`;

        // Call the OpenAI API
        const response = await fetch(config.endpoint, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${config.apiKey}`
            },
            body: JSON.stringify({
                model: config.model,
                messages: [
                    { role: "system", content: systemPrompt },
                    { role: "user", content: `Please analyze this text and provide spelling corrections and tone suggestions to make it more smart casual: "${text}"` }
                ],
                max_tokens: config.maxTokens,
                temperature: config.temperature,
                response_format: { type: "json_object" }
            })
        });

        if (!response.ok) {
            const errorData = await response.json();
            console.error("API request failed:", errorData);
            throw new Error(`API request failed with status ${response.status}`);
        }

        const data = await response.json();
        console.log("Analysis results:", data);
        
        // Parse the response content from the API
        try {
            const content = JSON.parse(data.choices[0].message.content);
            return {
                spelling: content.spelling || [],
                tone: content.tone || []
            };
        } catch (parseError) {
            console.error("Error parsing API response:", parseError);
            throw new Error("Failed to parse the AI response");
        }
    } catch (error) {
        console.error("Error analyzing document with AI:", error);
        
        // Return mock data for demonstration purposes when API fails
        console.warn("Falling back to mock implementation");
        return getMockAnalysisResults(text);
    }
}

/**
 * Provides mock analysis results for demonstration or testing
 * @param {string} text - The document text to analyze
 * @returns {object} - Mock analysis results
 */
function getMockAnalysisResults(text) {
    // Extract some common spelling errors
    const spellingErrors = [];
    const commonMisspellings = {
        "teh": "the",
        "recieve": "receive",
        "seperate": "separate",
        "definately": "definitely",
        "accomodate": "accommodate",
        "occured": "occurred",
        "thier": "their",
        "wierd": "weird",
        "alot": "a lot",
        "tommorrow": "tomorrow"
    };
    
    // Check for common misspellings in the text
    Object.keys(commonMisspellings).forEach(misspelling => {
        if (text.toLowerCase().includes(misspelling.toLowerCase())) {
            spellingErrors.push({
                word: misspelling,
                suggestion: commonMisspellings[misspelling]
            });
        }
    });
    
    // Generate some tone suggestions based on common formal phrases
    const tonePatterns = [
        {
            pattern: /it is (essential|imperative|mandatory|required)/i,
            original: "It is essential",
            suggestion: "It's important"
        },
        {
            pattern: /please be advised/i,
            original: "Please be advised",
            suggestion: "Just so you know"
        },
        {
            pattern: /as per our (discussion|conversation|agreement)/i,
            original: "As per our discussion",
            suggestion: "Based on what we talked about"
        },
        {
            pattern: /in accordance with/i,
            original: "In accordance with",
            suggestion: "Following"
        },
        {
            pattern: /we regret to inform you/i,
            original: "We regret to inform you",
            suggestion: "Unfortunately"
        },
        {
            pattern: /pursuant to/i,
            original: "Pursuant to",
            suggestion: "According to"
        },
        {
            pattern: /at your earliest convenience/i,
            original: "At your earliest convenience",
            suggestion: "When you have a chance"
        },
        {
            pattern: /kindly/i,
            original: "Kindly",
            suggestion: "Please"
        }
    ];
    
    const toneSuggestions = [];
    
    // Check for formal tone patterns
    tonePatterns.forEach(pattern => {
        const matches = text.match(pattern.pattern);
        if (matches) {
            // Get the matching sentence
            const sentences = text.split(/[.!?]+/);
            const matchingSentence = sentences.find(s => s.match(pattern.pattern));
            
            if (matchingSentence) {
                toneSuggestions.push({
                    original: matchingSentence.trim(),
                    suggestion: matchingSentence.replace(pattern.pattern, pattern.suggestion).trim()
                });
            }
        }
    });
    
    return {
        spelling: spellingErrors,
        tone: toneSuggestions
    };
}

/**
 * Suggests improvements to a paragraph based on the writing style
 * @param {string} paragraph - The paragraph text
 * @param {string} style - The desired writing style (e.g., "casual", "professional")
 * @returns {Promise<string>} - Improved paragraph
 */
async function suggestImprovement(paragraph, style = "smart casual") {
    try {
        // Check if API is configured
        if (!window.isApiConfigured()) {
            console.log("API not configured for improvement suggestions");
            return paragraph; // Return original if API not configured
        }
        
        // Get API configuration from localStorage
        const config = window.getApiConfiguration().openai;
        
        const response = await fetch(config.endpoint, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${config.apiKey}`
            },
            body: JSON.stringify({
                model: config.model,
                messages: [
                    {
                        role: "system",
                        content: `You are an AI assistant that helps improve writing. Rewrite the provided text in a ${style} tone while preserving the meaning.`
                    },
                    {
                        role: "user",
                        content: paragraph
                    }
                ],
                max_tokens: 1000,
                temperature: 0.7
            })
        });

        if (!response.ok) {
            throw new Error(`API request failed with status ${response.status}`);
        }

        const data = await response.json();
        return data.choices[0].message.content;
    } catch (error) {
        console.error("Error suggesting improvements:", error);
        return paragraph; // Return original if error
    }
}

// Export the functions for use in other modules
Office.onReady(() => {
    // Make functions available globally
    window.analyzeDocumentContent = analyzeDocumentContent;
    window.suggestImprovement = suggestImprovement;
});
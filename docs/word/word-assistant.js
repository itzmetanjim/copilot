/**
 * Odin Word Assistant - Main JavaScript
 * Implements all Word-specific AI features
 */

// Global variables
let currentSection = null;
let documentContent = '';

// Initialize when Office is ready
Office.onReady((info) => {
    console.log('Office.js ready for Word');
    initializeWordAssistant();
});

/**
 * Initialize the Word assistant
 */
async function initializeWordAssistant() {
    try {
        // Initialize AI service
        const configLoaded = await aiService.initialize();
        if (!configLoaded) {
            showConfigurationHelp();
            return;
        }

        // Setup UI components
        setupUI();
        
        // Setup provider selector
        const providerContainer = document.getElementById('provider-selector-container');
        if (providerContainer) {
            const selector = odinUI.createProviderSelector(aiService.currentProvider);
            providerContainer.appendChild(selector);
        }

        console.log('Word Assistant initialized successfully');
    } catch (error) {
        console.error('Failed to initialize Word Assistant:', error);
        aiService.showError('Failed to initialize. Please check your configuration.');
    }
}

/**
 * Setup UI components and event listeners
 */
function setupUI() {
    // Create action buttons for each response area
    const sections = ['draft', 'summarize', 'rewrite', 'format'];
    sections.forEach(section => {
        const actions = odinUI.createActionButtons(`${section}-response`, ['insert', 'copy', 'regenerate']);
        const container = document.getElementById(`${section}-response-actions`);
        if (container) {
            container.replaceWith(actions);
        }
    });

    // Setup keyboard shortcuts
    document.addEventListener('keydown', handleKeyboardShortcuts);
}

/**
 * Handle keyboard shortcuts
 */
function handleKeyboardShortcuts(event) {
    // Ctrl+Shift+O: Open Odin
    if (event.ctrlKey && event.shiftKey && event.key === 'O') {
        event.preventDefault();
        showDraftSection();
    }
    
    // Ctrl+Shift+S: Summarize
    if (event.ctrlKey && event.shiftKey && event.key === 'S') {
        event.preventDefault();
        showSummarizeSection();
    }
    
    // Ctrl+Shift+R: Rewrite
    if (event.ctrlKey && event.shiftKey && event.key === 'R') {
        event.preventDefault();
        showRewriteSection();
    }
}

/**
 * Show configuration help
 */
function showConfigurationHelp() {
    const helpContent = `
        <div class="odin-alert odin-alert-warning">
            <h3>🔧 Configuration Required</h3>
            <p>Welcome to Odin AI Assistant! To get started, you need to configure your AI provider API keys.</p>
            
            <div style="margin: 16px 0; padding: 12px; background: #f8f9fa; border-radius: 6px;">
                <p><strong>Quick Setup:</strong></p>
                <ol style="margin: 8px 0; padding-left: 20px;">
                    <li>Click the <strong>⚙️ Settings</strong> button (top-right corner)</li>
                    <li>Go to the <strong>🔑 API Keys</strong> tab</li>
                    <li>Add your API key for at least one provider:
                        <ul style="margin: 8px 0;">
                            <li><strong>Google Gemini</strong> - Get free API key at <a href="https://makersuite.google.com/app/apikey" target="_blank">Google AI Studio</a></li>
                            <li><strong>Azure OpenAI</strong> - Use your Azure OpenAI service credentials</li>
                            <li><strong>Cerebras</strong> - Get API key at <a href="https://cloud.cerebras.ai" target="_blank">Cerebras Cloud</a></li>
                        </ul>
                    </li>
                    <li>Save your settings and start using AI features!</li>
                </ol>
            </div>
            
            <p><strong>✨ Pro Tip:</strong> You can configure multiple providers and switch between them. The AI Assistant works entirely client-side - your API keys are stored securely in your browser's local storage.</p>
            
            <button class="odin-btn odin-btn-primary" onclick="showSettings()" style="margin-top: 12px;">
                ⚙️ Open Settings
            </button>
            
            <p style="margin-top: 16px; font-size: 14px; color: #666;">
                Need help? Visit the <a href="https://github.com/itzmetanjim/copilot" target="_blank">GitHub repository</a> for detailed setup instructions.
            </p>
        </div>
    `;
    
    document.querySelector('.odin-container').innerHTML = helpContent;
}

// === SECTION MANAGEMENT ===

/**
 * Show the draft section
 */
function showDraftSection() {
    hideAllSections();
    document.getElementById('draft-section').style.display = 'block';
    currentSection = 'draft';
    document.getElementById('draft-prompt').focus();
}

/**
 * Show the summarize section
 */
function showSummarizeSection() {
    hideAllSections();
    document.getElementById('summarize-section').style.display = 'block';
    currentSection = 'summarize';
}

/**
 * Show the rewrite section
 */
function showRewriteSection() {
    hideAllSections();
    document.getElementById('rewrite-section').style.display = 'block';
    currentSection = 'rewrite';
}

/**
 * Show the format section
 */
function showFormatSection() {
    hideAllSections();
    document.getElementById('format-section').style.display = 'block';
    currentSection = 'format';
}

/**
 * Hide all feature sections
 */
function hideAllSections() {
    const sections = ['draft-section', 'summarize-section', 'rewrite-section', 'format-section', 'settings-panel'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            section.style.display = 'none';
        }
    });
    currentSection = null;
}

// === DRAFT FEATURES ===

/**
 * Generate content based on prompt
 */
async function generateDraft() {
    const prompt = document.getElementById('draft-prompt').value.trim();
    const contentType = document.getElementById('draft-type').value;
    const tone = document.getElementById('draft-tone').value;
    
    if (!prompt) {
        aiService.showError('Please enter a prompt for content generation.');
        return;
    }
    
    try {
        const settings = odinUI.getSettings();
        const enhancedPrompt = `Create a ${contentType} with a ${tone} tone. ${prompt}`;
        
        const response = await aiService.generateText(enhancedPrompt, {
            maxTokens: settings.maxTokens,
            temperature: settings.temperature
        });
        
        odinUI.showResponse('draft-response', response, true);
        aiService.showSuccess('Content generated successfully!');
    } catch (error) {
        console.error('Draft generation error:', error);
        aiService.showError('Failed to generate content. Please try again.');
    }
}

// === SUMMARIZE FEATURES ===

/**
 * Summarize the entire document
 */
async function summarizeDocument() {
    try {
        const content = await getDocumentContent();
        if (!content.trim()) {
            aiService.showError('Document appears to be empty. Please add some content first.');
            return;
        }
        
        await performSummarization(content, 'document');
    } catch (error) {
        console.error('Document summarization error:', error);
        aiService.showError('Failed to summarize document. Please try again.');
    }
}

/**
 * Summarize selected text
 */
async function summarizeSelection() {
    try {
        const selectedText = await getSelectedText();
        if (!selectedText.trim()) {
            aiService.showError('Please select some text to summarize.');
            return;
        }
        
        await performSummarization(selectedText, 'selection');
    } catch (error) {
        console.error('Selection summarization error:', error);
        aiService.showError('Failed to summarize selection. Please try again.');
    }
}

/**
 * Perform summarization with specified options
 */
async function performSummarization(text, type) {
    const length = document.getElementById('summary-length').value;
    const format = document.getElementById('summary-format').value;
    
    try {
        const settings = odinUI.getSettings();
        const response = await aiService.summarizeText(text, {
            length: length,
            format: format,
            maxTokens: settings.maxTokens,
            temperature: settings.temperature
        });
        
        odinUI.showResponse('summarize-response', response, true);
        aiService.showSuccess(`${type.charAt(0).toUpperCase() + type.slice(1)} summarized successfully!`);
    } catch (error) {
        throw error;
    }
}

// === REWRITE FEATURES ===

/**
 * Rewrite selected text with specified tone
 */
async function rewriteText(tone) {
    try {
        const selectedText = await getSelectedText();
        if (!selectedText.trim()) {
            aiService.showError('Please select some text to rewrite.');
            return;
        }
        
        const customInstructions = document.getElementById('rewrite-instructions').value.trim();
        const settings = odinUI.getSettings();
        
        let enhancedTone = tone;
        if (customInstructions) {
            enhancedTone += `. Additional instructions: ${customInstructions}`;
        }
        
        const response = await aiService.rewriteText(selectedText, enhancedTone, {
            maxTokens: settings.maxTokens,
            temperature: settings.temperature
        });
        
        odinUI.showResponse('rewrite-response', response, true);
        aiService.showSuccess('Text rewritten successfully!');
    } catch (error) {
        console.error('Rewrite error:', error);
        aiService.showError('Failed to rewrite text. Please try again.');
    }
}

/**
 * Provide coaching and improvement suggestions
 */
async function coachText() {
    try {
        const selectedText = await getSelectedText();
        if (!selectedText.trim()) {
            aiService.showError('Please select some text to coach.');
            return;
        }
        
        const prompt = `Please analyze the following text and provide specific suggestions for improvement. Focus on clarity, grammar, style, and effectiveness. Provide the improved version along with explanations:\n\n${selectedText}`;
        
        const settings = odinUI.getSettings();
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3 // Lower temperature for coaching
        });
        
        odinUI.showResponse('rewrite-response', response, true);
        aiService.showSuccess('Coaching analysis complete!');
    } catch (error) {
        console.error('Coaching error:', error);
        aiService.showError('Failed to provide coaching. Please try again.');
    }
}

// === FORMAT FEATURES ===

/**
 * Apply document formatting
 */
async function formatDocument(formatType) {
    try {
        const selectedText = await getSelectedText();
        let textToFormat = selectedText;
        
        if (!textToFormat.trim()) {
            // If no selection, use entire document
            textToFormat = await getDocumentContent();
            if (!textToFormat.trim()) {
                aiService.showError('Document appears to be empty. Please add some content first.');
                return;
            }
        }
        
        let prompt = '';
        switch (formatType) {
            case 'headings':
                prompt = `Add appropriate headings and subheadings to structure this text:\n\n${textToFormat}`;
                break;
            case 'bullets':
                prompt = `Convert this text into a well-organized bullet point list:\n\n${textToFormat}`;
                break;
            case 'numbers':
                prompt = `Convert this text into a numbered list with logical sequence:\n\n${textToFormat}`;
                break;
            case 'table':
                prompt = `Organize this information into a clear table format:\n\n${textToFormat}`;
                break;
        }
        
        const settings = odinUI.getSettings();
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('format-response', response, true);
        aiService.showSuccess('Formatting suggestion generated!');
    } catch (error) {
        console.error('Formatting error:', error);
        aiService.showError('Failed to generate formatting. Please try again.');
    }
}

/**
 * Suggest custom formatting
 */
async function suggestFormatting() {
    const formatPrompt = document.getElementById('format-prompt').value.trim();
    
    if (!formatPrompt) {
        aiService.showError('Please describe the formatting you want.');
        return;
    }
    
    try {
        const selectedText = await getSelectedText();
        let textToFormat = selectedText;
        
        if (!textToFormat.trim()) {
            textToFormat = await getDocumentContent();
        }
        
        if (!textToFormat.trim()) {
            aiService.showError('No content available to format.');
            return;
        }
        
        const prompt = `Apply the following formatting instructions to this text: "${formatPrompt}"\n\nText to format:\n${textToFormat}`;
        
        const settings = odinUI.getSettings();
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('format-response', response, true);
        aiService.showSuccess('Custom formatting suggestion generated!');
    } catch (error) {
        console.error('Custom formatting error:', error);
        aiService.showError('Failed to generate custom formatting. Please try again.');
    }
}

// === ADVANCED FEATURES ===

/**
 * Analyze document content
 */
async function analyzeDocument() {
    try {
        const content = await getDocumentContent();
        if (!content.trim()) {
            aiService.showError('Document appears to be empty. Please add some content first.');
            return;
        }
        
        const prompt = `Analyze this document and provide insights including:
        1. Word count and reading level
        2. Main topics and themes
        3. Tone and style analysis
        4. Strengths and areas for improvement
        5. Suggestions for enhancement
        
        Document content:
        ${content}`;
        
        const settings = odinUI.getSettings();
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        // Show in draft section for now
        odinUI.showResponse('draft-response', response, false);
        showDraftSection();
        aiService.showSuccess('Document analysis complete!');
    } catch (error) {
        console.error('Document analysis error:', error);
        aiService.showError('Failed to analyze document. Please try again.');
    }
}

/**
 * Check grammar and suggest improvements
 */
async function checkGrammar() {
    try {
        const selectedText = await getSelectedText();
        let textToCheck = selectedText;
        
        if (!textToCheck.trim()) {
            textToCheck = await getDocumentContent();
        }
        
        if (!textToCheck.trim()) {
            aiService.showError('No content available to check.');
            return;
        }
        
        const prompt = `Check this text for grammar, spelling, and style issues. Provide corrections and explanations:\n\n${textToCheck}`;
        
        const settings = odinUI.getSettings();
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.2
        });
        
        odinUI.showResponse('rewrite-response', response, true);
        showRewriteSection();
        aiService.showSuccess('Grammar check complete!');
    } catch (error) {
        console.error('Grammar check error:', error);
        aiService.showError('Failed to check grammar. Please try again.');
    }
}

/**
 * Generate document outline
 */
async function generateOutline() {
    try {
        const content = await getDocumentContent();
        if (!content.trim()) {
            aiService.showError('Document appears to be empty. Please add some content first.');
            return;
        }
        
        const prompt = `Create a detailed outline for this document content. Include main sections, subsections, and key points:\n\n${content}`;
        
        const settings = odinUI.getSettings();
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('format-response', response, true);
        showFormatSection();
        aiService.showSuccess('Document outline generated!');
    } catch (error) {
        console.error('Outline generation error:', error);
        aiService.showError('Failed to generate outline. Please try again.');
    }
}

/**
 * Show settings panel
 */
function showSettings() {
    // Use the new settings UI instead of the old panel
    if (window.settingsUI) {
        window.settingsUI.open();
    } else {
        console.error('Settings UI not available');
        aiService.showError('Settings interface not available. Please refresh the page.');
    }
}

// === OFFICE.JS INTEGRATION ===

/**
 * Get the entire document content
 */
async function getDocumentContent() {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                const body = context.document.body;
                body.load('text');
                await context.sync();
                resolve(body.text);
            } catch (error) {
                reject(error);
            }
        });
    });
}

/**
 * Get selected text from document
 */
async function getSelectedText() {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                resolve(selection.text);
            } catch (error) {
                reject(error);
            }
        });
    });
}

/**
 * Insert text at current cursor position
 */
async function insertTextAtCursor(text) {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                const selection = context.document.getSelection();
                selection.insertText(text, Word.InsertLocation.replace);
                await context.sync();
                resolve(true);
            } catch (error) {
                reject(error);
            }
        });
    });
}

/**
 * Replace selected text
 */
async function replaceSelectedText(newText) {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                const selection = context.document.getSelection();
                selection.insertText(newText, Word.InsertLocation.replace);
                await context.sync();
                resolve(true);
            } catch (error) {
                reject(error);
            }
        });
    });
}

// === UTILITY FUNCTIONS ===

/**
 * Override the default insert text function to use Word-specific implementation
 */
odinUI.insertText = async function(responseDisplayId, event) {
    const display = document.getElementById(responseDisplayId);
    if (!display) return;
    
    const text = display.textContent;
    
    try {
        await insertTextAtCursor(text);
        aiService.showSuccess('Text inserted successfully!');
    } catch (error) {
        console.error('Insert text error:', error);
        aiService.showError('Failed to insert text. Please try again.');
    }
};

// Auto-focus on prompt when sections are shown
document.addEventListener('DOMContentLoaded', () => {
    // Setup section-specific focus behavior
    const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
            if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
                const target = mutation.target;
                if (target.style.display === 'block') {
                    // Section became visible, focus on appropriate input
                    setTimeout(() => {
                        if (target.id === 'draft-section') {
                            document.getElementById('draft-prompt')?.focus();
                        } else if (target.id === 'rewrite-section') {
                            document.getElementById('rewrite-instructions')?.focus();
                        } else if (target.id === 'format-section') {
                            document.getElementById('format-prompt')?.focus();
                        }
                    }, 100);
                }
            }
        });
    });
    
    const sections = ['draft-section', 'summarize-section', 'rewrite-section', 'format-section'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            observer.observe(section, { attributes: true });
        }
    });
});
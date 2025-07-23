/**
 * Odin Outlook Assistant - Main JavaScript
 * Implements all Outlook-specific AI features for email assistance
 */

// Global variables
let currentSection = null;
let emailContext = null;

// Initialize when Office is ready
Office.onReady((info) => {
    console.log('Office.js ready for Outlook');
    initializeOutlookAssistant();
});

/**
 * Initialize the Outlook assistant
 */
async function initializeOutlookAssistant() {
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

        console.log('Outlook Assistant initialized successfully');
    } catch (error) {
        console.error('Failed to initialize Outlook Assistant:', error);
        aiService.showError('Failed to initialize. Please check your configuration.');
    }
}

/**
 * Setup UI components
 */
function setupUI() {
    const sections = ['draft', 'summarize', 'coach', 'reply'];
    sections.forEach(section => {
        const actions = odinUI.createActionButtons(`${section}-response`, ['insert', 'copy']);
        const container = document.getElementById(`${section}-response-actions`);
        if (container) {
            container.replaceWith(actions);
        }
    });
}

/**
 * Show configuration help
 */
function showConfigurationHelp() {
    const helpContent = `
        <div class="odin-alert odin-alert-warning">
            <h3>🔧 Configuration Required</h3>
            <p>Welcome to Odin Outlook Assistant! To get started, you need to configure your AI provider API keys.</p>
            
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

/**
 * Show settings panel
 */
function showSettings() {
    // Use the new settings UI
    if (window.settingsUI) {
        window.settingsUI.open();
    } else {
        console.error('Settings UI not available');
        aiService.showError('Settings interface not available. Please refresh the page.');
    }
}

// Section management
function showDraftSection() {
    hideAllSections();
    document.getElementById('draft-section').style.display = 'block';
    currentSection = 'draft';
}

function showSummarizeSection() {
    hideAllSections();
    document.getElementById('summarize-section').style.display = 'block';
    currentSection = 'summarize';
}

function showCoachSection() {
    hideAllSections();
    document.getElementById('coach-section').style.display = 'block';
    currentSection = 'coach';
}

function showReplySection() {
    hideAllSections();
    document.getElementById('reply-section').style.display = 'block';
    currentSection = 'reply';
}

function hideAllSections() {
    const sections = ['draft-section', 'summarize-section', 'coach-section', 'reply-section'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            section.style.display = 'none';
        }
    });
}

// Main functions
async function draftEmail() {
    const purpose = document.getElementById('email-purpose').value.trim();
    const tone = document.getElementById('email-tone').value;
    
    if (!purpose) {
        aiService.showError('Please describe the email purpose.');
        return;
    }
    
    try {
        const settings = odinUI.getSettings();
        const prompt = `Draft a professional email with ${tone} tone for: ${purpose}`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: settings.temperature
        });
        
        odinUI.showResponse('draft-response', response, true);
        aiService.showSuccess('Email drafted successfully!');
    } catch (error) {
        console.error('Draft error:', error);
        aiService.showError('Failed to draft email. Please try again.');
    }
}

async function summarizeThread() {
    try {
        // Get current email content (simplified for now)
        const emailContent = "Email thread content would be extracted here";
        
        const settings = odinUI.getSettings();
        const prompt = `Summarize this email thread: ${emailContent}`;
        
        const response = await aiService.summarizeText(emailContent, {
            length: 'medium',
            format: 'bullets'
        });
        
        odinUI.showResponse('summarize-response', response, true);
        aiService.showSuccess('Thread summarized successfully!');
    } catch (error) {
        console.error('Summarize error:', error);
        aiService.showError('Failed to summarize thread. Please try again.');
    }
}

async function coachEmail() {
    try {
        const emailContent = "Current email content would be extracted here";
        
        const settings = odinUI.getSettings();
        const prompt = `Analyze this email and provide coaching suggestions: ${emailContent}`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('coach-response', response, false);
        aiService.showSuccess('Email coaching complete!');
    } catch (error) {
        console.error('Coach error:', error);
        aiService.showError('Failed to analyze email. Please try again.');
    }
}

async function generateReply() {
    const replyType = document.getElementById('reply-type').value;
    
    try {
        const emailContent = "Original email content would be extracted here";
        
        const settings = odinUI.getSettings();
        const prompt = `Generate a ${replyType} reply to: ${emailContent}`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: settings.temperature
        });
        
        odinUI.showResponse('reply-response', response, true);
        aiService.showSuccess('Reply generated successfully!');
    } catch (error) {
        console.error('Reply error:', error);
        aiService.showError('Failed to generate reply. Please try again.');
    }
}

// Override insert function for Outlook
odinUI.insertText = async function(responseDisplayId, event) {
    const display = document.getElementById(responseDisplayId);
    if (!display) return;
    
    const text = display.textContent;
    
    try {
        // Insert into email body or subject
        Office.context.mailbox.item.body.setAsync(
            text,
            { coercionType: Office.CoercionType.Text },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    aiService.showSuccess('Text inserted successfully!');
                } else {
                    aiService.showError('Failed to insert text.');
                }
            }
        );
    } catch (error) {
        console.error('Insert error:', error);
        aiService.showError('Failed to insert text. Please try again.');
    }
};
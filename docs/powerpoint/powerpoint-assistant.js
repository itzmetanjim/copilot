/**
 * Odin PowerPoint Assistant - Main JavaScript
 * Implements all PowerPoint-specific AI features for presentation creation and management
 */

// Global variables
let currentSection = null;
let selectedSlideType = 'title-content';
let presentationContext = null;

// Initialize when Office is ready
Office.onReady((info) => {
    console.log('Office.js ready for PowerPoint');
    initializePowerPointAssistant();
});

/**
 * Initialize the PowerPoint assistant
 */
async function initializePowerPointAssistant() {
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

        console.log('PowerPoint Assistant initialized successfully');
    } catch (error) {
        console.error('Failed to initialize PowerPoint Assistant:', error);
        aiService.showError('Failed to initialize. Please check your configuration.');
    }
}

/**
 * Setup UI components and event listeners
 */
function setupUI() {
    // Initialize Settings UI
    if (typeof SettingsUI !== 'undefined') {
        window.settingsUI = new SettingsUI();
        window.settingsUI.init();
        console.log('Settings UI initialized');
    } else {
        console.warn('SettingsUI class not available');
    }

    // Create action buttons for each response area
    const sections = ['create', 'slide', 'notes', 'design'];
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
    // Ctrl+Shift+P: Create presentation
    if (event.ctrlKey && event.shiftKey && event.key === 'P') {
        event.preventDefault();
        showCreateSection();
    }
    
    // Ctrl+Shift+N: Speaker notes
    if (event.ctrlKey && event.shiftKey && event.key === 'N') {
        event.preventDefault();
        showNotesSection();
    }
}

/**
 * Show configuration help
 */
function showConfigurationHelp() {
    const helpContent = `
        <div class="odin-alert odin-alert-warning">
            <h3>🔧 Configuration Required</h3>
            <p>Welcome to Odin PowerPoint Assistant! To get started, you need to configure your AI provider API keys.</p>
            
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

function showCreateSection() {
    hideAllSections();
    document.getElementById('create-section').style.display = 'block';
    currentSection = 'create';
    document.getElementById('presentation-topic').focus();
}

function showSlideSection() {
    hideAllSections();
    document.getElementById('slide-section').style.display = 'block';
    currentSection = 'slide';
}

function showNotesSection() {
    hideAllSections();
    document.getElementById('notes-section').style.display = 'block';
    currentSection = 'notes';
}

function showDesignSection() {
    hideAllSections();
    document.getElementById('design-section').style.display = 'block';
    currentSection = 'design';
}

function hideAllSections() {
    const sections = ['create-section', 'slide-section', 'notes-section', 'design-section', 'settings-panel'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            section.style.display = 'none';
        }
    });
    currentSection = null;
}

// === PRESENTATION CREATION ===

async function createPresentation() {
    const topic = document.getElementById('presentation-topic').value.trim();
    const type = document.getElementById('presentation-type').value;
    const audience = document.getElementById('target-audience').value.trim();
    const slideCount = document.getElementById('slide-count').value;
    const requirements = document.getElementById('additional-requirements').value.trim();
    
    if (!topic) {
        aiService.showError('Please enter a presentation topic.');
        return;
    }
    
    try {
        const settings = odinUI.getSettings();
        
        const prompt = `Create a comprehensive presentation outline for: "${topic}"
        
        Requirements:
        - Type: ${type}
        - Target Audience: ${audience || 'General audience'}
        - Number of slides: ${slideCount}
        ${requirements ? `- Additional requirements: ${requirements}` : ''}
        
        Please provide:
        1. Presentation title
        2. Detailed slide-by-slide outline with:
           - Slide number
           - Slide title
           - Key content points (3-5 bullet points per slide)
           - Suggested visuals or layout
        3. Opening hook/introduction strategy
        4. Key messages to emphasize
        5. Strong conclusion/call-to-action
        
        Format the outline clearly with slide numbers and titles.`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.4
        });
        
        // Display the outline in a structured format
        displayPresentationOutline(response);
        
        odinUI.showResponse('create-response', response, true);
        aiService.showSuccess('Presentation outline generated successfully!');
    } catch (error) {
        console.error('Presentation creation error:', error);
        aiService.showError('Failed to create presentation outline. Please try again.');
    }
}

function displayPresentationOutline(outline) {
    const outlineElement = document.getElementById('presentation-outline');
    
    // Parse the outline and format it nicely
    const lines = outline.split('\n');
    let formattedOutline = '';
    
    lines.forEach(line => {
        if (line.trim()) {
            if (line.match(/^\d+\.|Slide \d+/i)) {
                formattedOutline += `<div class="outline-item outline-level-1">${line}</div>`;
            } else if (line.match(/^[•\-\*]/)) {
                formattedOutline += `<div class="outline-item outline-level-2">${line}</div>`;
            } else if (line.match(/^\s+[•\-\*]/)) {
                formattedOutline += `<div class="outline-item outline-level-3">${line}</div>`;
            } else {
                formattedOutline += `<div class="outline-item outline-level-1">${line}</div>`;
            }
        }
    });
    
    outlineElement.innerHTML = formattedOutline;
    outlineElement.classList.remove('hidden');
}

// === SLIDE MANAGEMENT ===

function selectSlideType(type) {
    // Remove selection from all slide type buttons
    document.querySelectorAll('.slide-type-btn').forEach(btn => {
        btn.classList.remove('selected');
    });
    
    // Select the clicked type
    document.querySelector(`[data-type="${type}"]`).classList.add('selected');
    selectedSlideType = type;
}

async function addSlides() {
    const content = document.getElementById('slide-content').value.trim();
    const position = document.getElementById('insert-position').value;
    const slideCount = parseInt(document.getElementById('slides-to-add').value);
    
    if (!content) {
        aiService.showError('Please describe the slide content.');
        return;
    }
    
    try {
        const settings = odinUI.getSettings();
        
        const prompt = `Create ${slideCount} PowerPoint slide${slideCount > 1 ? 's' : ''} based on this content: "${content}"
        
        Slide type: ${selectedSlideType}
        
        For each slide, provide:
        1. Slide title
        2. Main content (bullet points, text, or structure)
        3. Layout suggestions
        4. Visual recommendations (images, charts, icons)
        
        Make sure the content is well-structured and presentation-ready.`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.4
        });
        
        // Show preview of the slides
        displaySlidePreview(response, slideCount);
        
        odinUI.showResponse('slide-response', response, true);
        aiService.showSuccess(`${slideCount} slide${slideCount > 1 ? 's' : ''} content generated!`);
    } catch (error) {
        console.error('Slide creation error:', error);
        aiService.showError('Failed to create slide content. Please try again.');
    }
}

function displaySlidePreview(content, slideCount) {
    const previewElement = document.getElementById('slide-preview');
    
    let previewHTML = `<h4>Preview of ${slideCount} New Slide${slideCount > 1 ? 's' : ''}</h4>`;
    
    // Create a simple preview structure
    for (let i = 1; i <= slideCount; i++) {
        previewHTML += `
            <div class="slide-content">
                <div class="slide-number">Slide ${i}</div>
                <div>Content will be generated based on your specifications</div>
            </div>
        `;
    }
    
    previewElement.innerHTML = previewHTML;
    previewElement.classList.remove('hidden');
}

// === SPEAKER NOTES ===

async function generateNotesFor(scope) {
    const style = document.getElementById('notes-style').value;
    const duration = document.getElementById('presentation-duration').value;
    const context = document.getElementById('notes-context').value.trim();
    
    try {
        let slideInfo = '';
        
        // Get slide information based on scope
        switch (scope) {
            case 'current':
                slideInfo = await getCurrentSlideInfo();
                break;
            case 'selected':
                slideInfo = await getSelectedSlidesInfo();
                break;
            case 'all':
                slideInfo = await getAllSlidesInfo();
                break;
        }
        
        if (!slideInfo) {
            aiService.showError('Could not retrieve slide information.');
            return;
        }
        
        const settings = odinUI.getSettings();
        
        const prompt = `Generate ${style} speaker notes for the following slide(s):
        
        ${slideInfo}
        
        Notes requirements:
        - Style: ${style}
        ${duration ? `- Presentation duration: ${duration} minutes` : ''}
        ${context ? `- Additional context: ${context}` : ''}
        
        Please provide:
        1. Speaking points for each slide
        2. Transition phrases between slides
        3. Key messages to emphasize
        4. Timing suggestions
        5. Tips for delivery
        
        Make the notes natural and easy to follow during presentation.`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        // Display in speaker notes format
        displaySpeakerNotes(response);
        
        odinUI.showResponse('notes-response', response, true);
        aiService.showSuccess('Speaker notes generated successfully!');
    } catch (error) {
        console.error('Speaker notes error:', error);
        aiService.showError('Failed to generate speaker notes. Please try again.');
    }
}

function displaySpeakerNotes(notes) {
    const notesElement = document.getElementById('speaker-notes-preview');
    notesElement.innerHTML = `
        <h4>📝 Speaker Notes Preview</h4>
        <div class="notes-content">${notes.replace(/\n/g, '<br>')}</div>
    `;
    notesElement.classList.remove('hidden');
}

// === DESIGN ASSISTANCE ===

async function improveDesign(designType) {
    try {
        const presentationInfo = await getPresentationInfo();
        const style = document.getElementById('design-style').value;
        
        let prompt = '';
        switch (designType) {
            case 'layout':
                prompt = `Analyze the layout of this presentation and suggest improvements for better visual hierarchy and readability: ${presentationInfo}`;
                break;
            case 'colors':
                prompt = `Suggest a cohesive color scheme for this presentation that aligns with ${style} design style: ${presentationInfo}`;
                break;
            case 'fonts':
                prompt = `Recommend typography improvements for this presentation, including font combinations and sizing: ${presentationInfo}`;
                break;
            case 'consistency':
                prompt = `Check for design consistency issues in this presentation and provide specific recommendations: ${presentationInfo}`;
                break;
        }
        
        const settings = odinUI.getSettings();
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('design-response', response, false);
        aiService.showSuccess('Design analysis complete!');
    } catch (error) {
        console.error('Design improvement error:', error);
        aiService.showError('Failed to analyze design. Please try again.');
    }
}

async function customDesignSuggestions() {
    const instructions = document.getElementById('design-instructions').value.trim();
    const style = document.getElementById('design-style').value;
    
    if (!instructions) {
        aiService.showError('Please provide design instructions.');
        return;
    }
    
    try {
        const presentationInfo = await getPresentationInfo();
        const settings = odinUI.getSettings();
        
        const prompt = `Based on these design instructions: "${instructions}"
        
        Current presentation: ${presentationInfo}
        Preferred style: ${style}
        
        Provide specific design recommendations including:
        1. Layout improvements
        2. Color palette suggestions
        3. Typography recommendations
        4. Visual element suggestions
        5. Step-by-step implementation guide`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.4
        });
        
        odinUI.showResponse('design-response', response, false);
        aiService.showSuccess('Custom design suggestions generated!');
    } catch (error) {
        console.error('Custom design error:', error);
        aiService.showError('Failed to generate design suggestions. Please try again.');
    }
}

// === ADVANCED FEATURES ===

async function analyzePresentation() {
    try {
        const presentationInfo = await getPresentationInfo();
        const settings = odinUI.getSettings();
        
        const prompt = `Analyze this PowerPoint presentation and provide comprehensive feedback:
        
        ${presentationInfo}
        
        Please evaluate:
        1. Content structure and flow
        2. Visual design effectiveness
        3. Audience engagement potential
        4. Areas for improvement
        5. Strengths to maintain
        6. Suggestions for enhancement`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('design-response', response, false);
        aiService.showSuccess('Presentation analysis complete!');
    } catch (error) {
        console.error('Presentation analysis error:', error);
        aiService.showError('Failed to analyze presentation. Please try again.');
    }
}

async function generateSummary() {
    try {
        const presentationInfo = await getPresentationInfo();
        const settings = odinUI.getSettings();
        
        const prompt = `Create a concise summary of this presentation:
        
        ${presentationInfo}
        
        Include:
        1. Main topic and objective
        2. Key points covered
        3. Important takeaways
        4. Conclusion/call-to-action
        5. One-sentence elevator pitch`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('design-response', response, true);
        aiService.showSuccess('Presentation summary generated!');
    } catch (error) {
        console.error('Summary generation error:', error);
        aiService.showError('Failed to generate summary. Please try again.');
    }
}

function showSettings() {
    // Use the new settings UI instead of the old panel
    if (window.settingsUI) {
        window.settingsUI.open();
    } else {
        console.error('Settings UI not available');
        aiService.showError('Settings interface not available. Please refresh the page.');
    }
}

// === POWERPOINT.JS INTEGRATION ===

async function getCurrentSlideInfo() {
    return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
            try {
                const slide = context.presentation.slides.getItemAt(0); // Current slide
                slide.load('title');
                await context.sync();
                resolve(`Current slide title: ${slide.title}`);
            } catch (error) {
                resolve('Current slide information not available');
            }
        });
    });
}

async function getPresentationInfo() {
    return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
            try {
                const presentation = context.presentation;
                const slides = presentation.slides;
                slides.load('items');
                await context.sync();
                
                let info = `Presentation with ${slides.items.length} slides`;
                resolve(info);
            } catch (error) {
                resolve('Presentation information not available');
            }
        });
    });
}

async function getSelectedSlidesInfo() {
    // Placeholder for selected slides info
    return 'Selected slides information';
}

async function getAllSlidesInfo() {
    return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
            try {
                const slides = context.presentation.slides;
                slides.load('items');
                await context.sync();
                
                let info = `All ${slides.items.length} slides in presentation`;
                resolve(info);
            } catch (error) {
                resolve('All slides information not available');
            }
        });
    });
}

// === UTILITY FUNCTIONS ===

// Override the default insert text function for PowerPoint
odinUI.insertText = async function(responseDisplayId, event) {
    const display = document.getElementById(responseDisplayId);
    if (!display) return;
    
    const text = display.textContent;
    
    try {
        await PowerPoint.run(async (context) => {
            // Insert text into current slide or create new slide
            const slides = context.presentation.slides;
            slides.load('items');
            await context.sync();
            
            if (slides.items.length === 0) {
                // Create new slide if none exist
                const slide = slides.add();
                const textBox = slide.shapes.addTextBox(text);
                textBox.left = 50;
                textBox.top = 100;
                textBox.width = 600;
                textBox.height = 400;
            } else {
                // Add to current slide
                const slide = slides.getItemAt(0);
                const textBox = slide.shapes.addTextBox(text);
                textBox.left = 50;
                textBox.top = 100;
                textBox.width = 600;
                textBox.height = 400;
            }
            
            await context.sync();
        });
        
        aiService.showSuccess('Content inserted successfully!');
    } catch (error) {
        console.error('Insert content error:', error);
        aiService.showError('Failed to insert content. Please try again.');
    }
};

// Auto-focus on appropriate inputs when sections are shown
document.addEventListener('DOMContentLoaded', () => {
    const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
            if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
                const target = mutation.target;
                if (target.style.display === 'block') {
                    setTimeout(() => {
                        if (target.id === 'create-section') {
                            document.getElementById('presentation-topic')?.focus();
                        } else if (target.id === 'slide-section') {
                            document.getElementById('slide-content')?.focus();
                        } else if (target.id === 'design-section') {
                            document.getElementById('design-instructions')?.focus();
                        }
                    }, 100);
                }
            }
        });
    });
    
    const sections = ['create-section', 'slide-section', 'notes-section', 'design-section'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            observer.observe(section, { attributes: true });
        }
    });
});
/**
 * Shared UI Components for Odin AI Assistant Suite
 * Reusable components for consistent user experience across all add-ins
 */

class OdinUI {
  constructor() {
    this.components = new Map();
    this.initialized = false;
  }

  /**
   * Initialize the UI system
   */
  initialize() {
    if (this.initialized) return;
    
    this.createGlobalElements();
    this.setupEventListeners();
    this.initialized = true;
    console.log('Odin UI initialized');
  }

  /**
   * Create global UI elements that all features need
   */
  createGlobalElements() {
    // Create notification containers if they don't exist
    if (!document.getElementById('ai-loading')) {
      const loading = document.createElement('div');
      loading.id = 'ai-loading';
      loading.className = 'odin-loading hidden';
      loading.innerHTML = '<div class="odin-spinner"></div><span>Processing...</span>';
      document.body.appendChild(loading);
    }

    if (!document.getElementById('ai-error')) {
      const error = document.createElement('div');
      error.id = 'ai-error';
      error.className = 'odin-alert odin-alert-error hidden';
      document.body.appendChild(error);
    }

    if (!document.getElementById('ai-success')) {
      const success = document.createElement('div');
      success.id = 'ai-success';
      success.className = 'odin-alert odin-alert-success hidden';
      document.body.appendChild(success);
    }
  }

  /**
   * Setup global event listeners
   */
  setupEventListeners() {
    // Close alerts when clicked
    document.addEventListener('click', (e) => {
      if (e.target.classList.contains('odin-alert')) {
        e.target.style.display = 'none';
      }
    });
  }

  /**
   * Create a header component
   */
  createHeader(title, subtitle = '') {
    const header = document.createElement('div');
    header.className = 'odin-header';
    
    const titleElement = document.createElement('h1');
    titleElement.textContent = title;
    header.appendChild(titleElement);
    
    if (subtitle) {
      const subtitleElement = document.createElement('p');
      subtitleElement.textContent = subtitle;
      subtitleElement.style.margin = '0';
      subtitleElement.style.opacity = '0.9';
      header.appendChild(subtitleElement);
    }
    
    return header;
  }

  /**
   * Create a provider selector
   */
  createProviderSelector(currentProvider = 'gemini') {
    const container = document.createElement('div');
    container.className = 'odin-provider-selector';
    
    const label = document.createElement('label');
    label.textContent = 'AI Provider:';
    label.htmlFor = 'provider-select';
    
    const select = document.createElement('select');
    select.id = 'provider-select';
    select.className = 'odin-select';
    select.style.width = 'auto';
    select.style.minWidth = '120px';
    
    const providers = [
      { value: 'gemini', label: 'Google Gemini' },
      { value: 'azure-openai', label: 'Azure OpenAI' },
      { value: 'cerebras', label: 'Cerebras' }
    ];
    
    providers.forEach(provider => {
      const option = document.createElement('option');
      option.value = provider.value;
      option.textContent = provider.label;
      option.selected = provider.value === currentProvider;
      select.appendChild(option);
    });
    
    select.addEventListener('change', (e) => {
      if (window.aiService) {
        window.aiService.switchProvider(e.target.value);
      }
    });
    
    container.appendChild(label);
    container.appendChild(select);
    
    return container;
  }

  /**
   * Create a form group with label and input
   */
  createFormGroup(labelText, inputType = 'text', placeholder = '', required = false) {
    const group = document.createElement('div');
    group.className = 'odin-form-group';
    
    const label = document.createElement('label');
    label.className = 'odin-label';
    label.textContent = labelText;
    
    let input;
    if (inputType === 'textarea') {
      input = document.createElement('textarea');
      input.className = 'odin-textarea';
      input.rows = 4;
    } else if (inputType === 'select') {
      input = document.createElement('select');
      input.className = 'odin-select';
    } else {
      input = document.createElement('input');
      input.type = inputType;
      input.className = 'odin-input';
    }
    
    input.placeholder = placeholder;
    input.required = required;
    
    group.appendChild(label);
    group.appendChild(input);
    
    return { group, label, input };
  }

  /**
   * Create a button with specified style and click handler
   */
  createButton(text, className = 'odin-btn-primary', clickHandler = null) {
    const button = document.createElement('button');
    button.className = `odin-btn ${className}`;
    button.textContent = text;
    
    if (clickHandler) {
      button.addEventListener('click', clickHandler);
    }
    
    return button;
  }

  /**
   * Create a card component
   */
  createCard(title, body, footer = null) {
    const card = document.createElement('div');
    card.className = 'odin-card';
    
    if (title) {
      const header = document.createElement('div');
      header.className = 'odin-card-header';
      header.textContent = title;
      card.appendChild(header);
    }
    
    const cardBody = document.createElement('div');
    cardBody.className = 'odin-card-body';
    
    if (typeof body === 'string') {
      cardBody.innerHTML = body;
    } else {
      cardBody.appendChild(body);
    }
    
    card.appendChild(cardBody);
    
    if (footer) {
      const cardFooter = document.createElement('div');
      cardFooter.className = 'odin-card-footer';
      
      if (typeof footer === 'string') {
        cardFooter.innerHTML = footer;
      } else {
        cardFooter.appendChild(footer);
      }
      
      card.appendChild(cardFooter);
    }
    
    return card;
  }

  /**
   * Create a response display area
   */
  createResponseDisplay(id = 'response-display') {
    const display = document.createElement('div');
    display.id = id;
    display.className = 'odin-response hidden';
    display.innerHTML = '<em>Response will appear here...</em>';
    
    return display;
  }

  /**
   * Create action buttons for response handling
   */
  createActionButtons(responseDisplayId, actions = ['insert', 'copy', 'regenerate']) {
    const container = document.createElement('div');
    container.className = 'odin-action-buttons hidden';
    container.id = `${responseDisplayId}-actions`;
    
    const buttonConfigs = {
      insert: { text: 'Insert', class: 'odin-btn-primary', handler: this.insertText },
      copy: { text: 'Copy', class: 'odin-btn-secondary', handler: this.copyText },
      regenerate: { text: 'Regenerate', class: 'odin-btn-outline', handler: this.regenerateResponse },
      approve: { text: 'Apply', class: 'odin-btn-success', handler: this.applyChanges },
      reject: { text: 'Reject', class: 'odin-btn-danger', handler: this.rejectChanges }
    };
    
    actions.forEach(action => {
      if (buttonConfigs[action]) {
        const button = this.createButton(
          buttonConfigs[action].text,
          buttonConfigs[action].class,
          (e) => buttonConfigs[action].handler(responseDisplayId, e)
        );
        container.appendChild(button);
      }
    });
    
    return container;
  }

  /**
   * Create a feature grid
   */
  createFeatureGrid(features) {
    const grid = document.createElement('div');
    grid.className = 'odin-feature-grid';
    
    features.forEach(feature => {
      const item = document.createElement('div');
      item.className = 'odin-feature-item';
      item.dataset.action = feature.action;
      
      const title = document.createElement('h3');
      title.textContent = feature.title;
      
      const description = document.createElement('p');
      description.textContent = feature.description;
      
      item.appendChild(title);
      item.appendChild(description);
      
      if (feature.clickHandler) {
        item.addEventListener('click', feature.clickHandler);
      }
      
      grid.appendChild(item);
    });
    
    return grid;
  }

  /**
   * Show a response in the display area
   */
  showResponse(responseDisplayId, content, showActions = true) {
    const display = document.getElementById(responseDisplayId);
    const actions = document.getElementById(`${responseDisplayId}-actions`);
    
    if (display) {
      display.textContent = content;
      display.classList.remove('hidden');
      display.scrollTop = 0;
    }
    
    if (actions && showActions) {
      actions.classList.remove('hidden');
    }
  }

  /**
   * Clear response display
   */
  clearResponse(responseDisplayId) {
    const display = document.getElementById(responseDisplayId);
    const actions = document.getElementById(`${responseDisplayId}-actions`);
    
    if (display) {
      display.innerHTML = '<em>Response will appear here...</em>';
      display.classList.add('hidden');
    }
    
    if (actions) {
      actions.classList.add('hidden');
    }
  }

  /**
   * Insert text into Office document
   */
  async insertText(responseDisplayId, event) {
    const display = document.getElementById(responseDisplayId);
    if (!display) return;
    
    const text = display.textContent;
    
    try {
      await Office.context.document.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Text },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            window.aiService?.showSuccess('Text inserted successfully!');
          } else {
            window.aiService?.showError('Failed to insert text: ' + result.error.message);
          }
        }
      );
    } catch (error) {
      console.error('Insert text error:', error);
      window.aiService?.showError('Failed to insert text. Please try again.');
    }
  }

  /**
   * Copy text to clipboard
   */
  async copyText(responseDisplayId, event) {
    const display = document.getElementById(responseDisplayId);
    if (!display) return;
    
    const text = display.textContent;
    
    try {
      await navigator.clipboard.writeText(text);
      window.aiService?.showSuccess('Text copied to clipboard!');
    } catch (error) {
      console.error('Copy text error:', error);
      
      // Fallback for older browsers
      try {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        textArea.remove();
        window.aiService?.showSuccess('Text copied to clipboard!');
      } catch (fallbackError) {
        window.aiService?.showError('Failed to copy text. Please manually select and copy.');
      }
    }
  }

  /**
   * Regenerate response
   */
  async regenerateResponse(responseDisplayId, event) {
    // This should be implemented by the specific feature
    console.log('Regenerate response for:', responseDisplayId);
    window.aiService?.showError('Regenerate functionality not implemented for this feature.');
  }

  /**
   * Apply changes (for editing features)
   */
  async applyChanges(responseDisplayId, event) {
    // This should be implemented by the specific feature
    console.log('Apply changes for:', responseDisplayId);
    await this.insertText(responseDisplayId, event);
  }

  /**
   * Reject changes (for editing features)
   */
  async rejectChanges(responseDisplayId, event) {
    this.clearResponse(responseDisplayId);
    window.aiService?.showSuccess('Changes rejected.');
  }

  /**
   * Create a loading overlay
   */
  createLoadingOverlay(message = 'Processing...') {
    const overlay = document.createElement('div');
    overlay.className = 'odin-loading-overlay';
    overlay.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(255, 255, 255, 0.9);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
    `;
    
    overlay.innerHTML = `
      <div class="odin-loading">
        <div class="odin-spinner"></div>
        <span>${message}</span>
      </div>
    `;
    
    return overlay;
  }

  /**
   * Show loading overlay
   */
  showLoadingOverlay(message = 'Processing...') {
    this.hideLoadingOverlay(); // Remove any existing overlay
    
    const overlay = this.createLoadingOverlay(message);
    overlay.id = 'odin-loading-overlay';
    document.body.appendChild(overlay);
  }

  /**
   * Hide loading overlay
   */
  hideLoadingOverlay() {
    const overlay = document.getElementById('odin-loading-overlay');
    if (overlay) {
      overlay.remove();
    }
  }

  /**
   * Create a settings panel
   */
  createSettingsPanel() {
    const panel = document.createElement('div');
    panel.className = 'odin-card';
    panel.innerHTML = `
      <div class="odin-card-header">
        Settings
      </div>
      <div class="odin-card-body">
        <div class="odin-form-group">
          <label class="odin-label">Temperature (Creativity)</label>
          <input type="range" id="temperature-slider" min="0" max="1" step="0.1" value="0.7" class="odin-input">
          <small>Lower = More focused, Higher = More creative</small>
        </div>
        <div class="odin-form-group">
          <label class="odin-label">Max Response Length</label>
          <select id="max-tokens-select" class="odin-select">
            <option value="1024">Short (1K tokens)</option>
            <option value="2048">Medium (2K tokens)</option>
            <option value="4096" selected>Long (4K tokens)</option>
            <option value="8192">Very Long (8K tokens)</option>
          </select>
        </div>
      </div>
    `;
    
    return panel;
  }

  /**
   * Get current settings from the UI
   */
  getSettings() {
    const temperatureSlider = document.getElementById('temperature-slider');
    const maxTokensSelect = document.getElementById('max-tokens-select');
    
    return {
      temperature: temperatureSlider ? parseFloat(temperatureSlider.value) : 0.7,
      maxTokens: maxTokensSelect ? parseInt(maxTokensSelect.value) : 4096
    };
  }

  /**
   * Initialize Office.js when ready
   */
  initializeOffice(callback) {
    if (typeof Office !== 'undefined') {
      Office.onReady((info) => {
        console.log('Office.js ready:', info);
        if (callback) callback(info);
      });
    } else {
      console.warn('Office.js not available - running in standalone mode');
      if (callback) callback(null);
    }
  }
}

// Create global instance
const odinUI = new OdinUI();

// Auto-initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => odinUI.initialize());
} else {
  odinUI.initialize();
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { OdinUI, odinUI };
} else {
  window.odinUI = odinUI;
  window.OdinUI = OdinUI;
}
/**
 * Settings UI Component - Provides interface for managing API keys and provider settings
 */

class SettingsUI {
  constructor() {
    this.isOpen = false;
    this.currentTab = 'providers';
    this.settingsManager = window.settingsManager;
  }

  /**
   * Initialize settings UI
   */
  init() {
    this.createSettingsModal();
    this.attachEventListeners();
    this.loadSettings();
  }

  /**
   * Create the settings modal HTML
   */
  createSettingsModal() {
    const modalHTML = `
      <div id="settings-modal" class="settings-modal" style="display: none;">
        <div class="settings-modal-overlay"></div>
        <div class="settings-modal-content">
          <div class="settings-header">
            <h2>🔧 AI Assistant Settings</h2>
            <button class="settings-close-btn" id="settings-close">&times;</button>
          </div>
          
          <div class="settings-tabs">
            <button class="settings-tab ${this.currentTab === 'providers' ? 'active' : ''}" data-tab="providers">
              🔑 API Keys
            </button>
            <button class="settings-tab ${this.currentTab === 'models' ? 'active' : ''}" data-tab="models">
              🤖 Models
            </button>
            <button class="settings-tab ${this.currentTab === 'advanced' ? 'active' : ''}" data-tab="advanced">
              ⚙️ Advanced
            </button>
          </div>

          <div class="settings-content">
            <div id="providers-tab" class="settings-tab-content ${this.currentTab === 'providers' ? 'active' : ''}">
              <div class="settings-description">
                <p>Configure your AI provider API keys. You can use one or more providers - leave unused providers empty.</p>
                <div class="settings-status" id="provider-status"></div>
              </div>
              <div id="providers-list"></div>
            </div>

            <div id="models-tab" class="settings-tab-content ${this.currentTab === 'models' ? 'active' : ''}">
              <div class="settings-description">
                <p>Select your preferred models and configure their parameters.</p>
              </div>
              <div id="models-config"></div>
            </div>

            <div id="advanced-tab" class="settings-tab-content ${this.currentTab === 'advanced' ? 'active' : ''}">
              <div class="settings-description">
                <p>Advanced features and data management options.</p>
              </div>
              <div id="advanced-options"></div>
            </div>
          </div>

          <div class="settings-footer">
            <button class="btn btn-secondary" id="settings-cancel">Cancel</button>
            <button class="btn btn-primary" id="settings-save">Save Settings</button>
          </div>
        </div>
      </div>

      <!-- Settings button for toolbar -->
      <button class="settings-toggle-btn" id="settings-toggle" title="Open Settings">
        ⚙️
      </button>
    `;

    // Add to page
    document.body.insertAdjacentHTML('beforeend', modalHTML);
    
    // Add styles
    this.addSettingsStyles();
  }

  /**
   * Add CSS styles for settings UI
   */
  addSettingsStyles() {
    const styles = `
      <style id="settings-styles">
        .settings-toggle-btn {
          position: fixed;
          top: 10px;
          right: 10px;
          z-index: 1000;
          background: var(--primary-color, #2B579A);
          color: white;
          border: none;
          border-radius: 50%;
          width: 40px;
          height: 40px;
          font-size: 16px;
          cursor: pointer;
          box-shadow: 0 2px 8px rgba(0,0,0,0.2);
          transition: all 0.2s ease;
        }

        .settings-toggle-btn:hover {
          background: var(--primary-color-dark, #1B4082);
          transform: scale(1.1);
        }

        .settings-modal {
          position: fixed;
          top: 0;
          left: 0;
          width: 100%;
          height: 100%;
          z-index: 10000;
        }

        .settings-modal-overlay {
          position: absolute;
          top: 0;
          left: 0;
          width: 100%;
          height: 100%;
          background: rgba(0, 0, 0, 0.5);
        }

        .settings-modal-content {
          position: absolute;
          top: 50%;
          left: 50%;
          transform: translate(-50%, -50%);
          background: white;
          border-radius: 12px;
          box-shadow: 0 4px 32px rgba(0, 0, 0, 0.3);
          width: 90%;
          max-width: 800px;
          max-height: 90vh;
          overflow: hidden;
          display: flex;
          flex-direction: column;
        }

        .settings-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 20px;
          border-bottom: 1px solid #e5e5e5;
          background: #f8f9fa;
        }

        .settings-header h2 {
          margin: 0;
          color: var(--primary-color, #2B579A);
          font-size: 20px;
        }

        .settings-close-btn {
          background: none;
          border: none;
          font-size: 24px;
          cursor: pointer;
          color: #666;
          padding: 0;
          width: 30px;
          height: 30px;
          display: flex;
          align-items: center;
          justify-content: center;
          border-radius: 4px;
        }

        .settings-close-btn:hover {
          background: #e5e5e5;
          color: #333;
        }

        .settings-tabs {
          display: flex;
          background: #f8f9fa;
          border-bottom: 1px solid #e5e5e5;
        }

        .settings-tab {
          flex: 1;
          padding: 12px 16px;
          background: none;
          border: none;
          cursor: pointer;
          font-size: 14px;
          color: #666;
          transition: all 0.2s ease;
          border-bottom: 3px solid transparent;
        }

        .settings-tab:hover {
          background: #e9ecef;
          color: #333;
        }

        .settings-tab.active {
          color: var(--primary-color, #2B579A);
          border-bottom-color: var(--primary-color, #2B579A);
          background: white;
        }

        .settings-content {
          flex: 1;
          overflow-y: auto;
          padding: 20px;
        }

        .settings-tab-content {
          display: none;
        }

        .settings-tab-content.active {
          display: block;
        }

        .settings-description {
          margin-bottom: 20px;
          padding: 12px;
          background: #f8f9fa;
          border-radius: 8px;
          border-left: 4px solid var(--primary-color, #2B579A);
        }

        .settings-description p {
          margin: 0;
          color: #666;
        }

        .settings-status {
          margin-top: 8px;
          padding: 8px;
          border-radius: 4px;
          font-size: 14px;
        }

        .settings-status.success {
          background: #d4edda;
          color: #155724;
        }

        .settings-status.warning {
          background: #fff3cd;
          color: #856404;
        }

        .settings-status.error {
          background: #f8d7da;
          color: #721c24;
        }

        .provider-card {
          margin-bottom: 20px;
          border: 1px solid #e5e5e5;
          border-radius: 8px;
          padding: 16px;
          background: white;
        }

        .provider-card.configured {
          border-color: #28a745;
          background: #f8fff9;
        }

        .provider-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 12px;
        }

        .provider-name {
          font-weight: 600;
          color: #333;
          margin: 0;
        }

        .provider-status-badge {
          padding: 4px 8px;
          border-radius: 12px;
          font-size: 12px;
          font-weight: 500;
        }

        .provider-status-badge.configured {
          background: #28a745;
          color: white;
        }

        .provider-status-badge.unconfigured {
          background: #6c757d;
          color: white;
        }

        .provider-description {
          color: #666;
          font-size: 14px;
          margin-bottom: 16px;
        }

        .provider-fields {
          display: grid;
          gap: 12px;
        }

        .field-group {
          display: flex;
          flex-direction: column;
        }

        .field-label {
          font-size: 14px;
          font-weight: 500;
          color: #333;
          margin-bottom: 4px;
        }

        .field-input {
          padding: 8px 12px;
          border: 1px solid #ccc;
          border-radius: 4px;
          font-size: 14px;
          transition: border-color 0.2s ease;
        }

        .field-input:focus {
          outline: none;
          border-color: var(--primary-color, #2B579A);
          box-shadow: 0 0 0 2px rgba(43, 87, 154, 0.2);
        }

        .field-input[type="password"] {
          font-family: monospace;
        }

        .toggle-password {
          position: relative;
        }

        .toggle-password-btn {
          position: absolute;
          right: 8px;
          top: 50%;
          transform: translateY(-50%);
          background: none;
          border: none;
          cursor: pointer;
          color: #666;
          font-size: 14px;
          padding: 4px;
        }

        .model-config-group {
          margin-bottom: 20px;
          padding: 16px;
          border: 1px solid #e5e5e5;
          border-radius: 8px;
        }

        .model-config-header {
          display: flex;
          justify-content: between;
          align-items: center;
          margin-bottom: 12px;
        }

        .model-select {
          padding: 8px 12px;
          border: 1px solid #ccc;
          border-radius: 4px;
          background: white;
          min-width: 200px;
        }

        .parameter-group {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 12px;
          margin-top: 12px;
        }

        .settings-footer {
          display: flex;
          justify-content: flex-end;
          gap: 12px;
          padding: 20px;
          border-top: 1px solid #e5e5e5;
          background: #f8f9fa;
        }

        .btn {
          padding: 10px 20px;
          border: none;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .btn-primary {
          background: var(--primary-color, #2B579A);
          color: white;
        }

        .btn-primary:hover {
          background: var(--primary-color-dark, #1B4082);
        }

        .btn-secondary {
          background: #6c757d;
          color: white;
        }

        .btn-secondary:hover {
          background: #545b62;
        }

        .btn-danger {
          background: #dc3545;
          color: white;
        }

        .btn-danger:hover {
          background: #c82333;
        }

        .advanced-option {
          margin-bottom: 16px;
          padding: 12px;
          border: 1px solid #e5e5e5;
          border-radius: 6px;
        }

        .advanced-option h4 {
          margin: 0 0 8px 0;
          color: #333;
        }

        .advanced-option p {
          margin: 0 0 12px 0;
          color: #666;
          font-size: 14px;
        }

        .loading-spinner {
          display: inline-block;
          width: 16px;
          height: 16px;
          border: 2px solid #f3f3f3;
          border-top: 2px solid var(--primary-color, #2B579A);
          border-radius: 50%;
          animation: spin 1s linear infinite;
        }

        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }

        @media (max-width: 600px) {
          .settings-modal-content {
            width: 95%;
            height: 95vh;
          }
          
          .settings-tabs {
            flex-direction: column;
          }
          
          .parameter-group {
            grid-template-columns: 1fr;
          }
        }
      </style>
    `;

    document.head.insertAdjacentHTML('beforeend', styles);
  }

  /**
   * Attach event listeners
   */
  attachEventListeners() {
    // Toggle settings modal
    document.getElementById('settings-toggle').addEventListener('click', () => {
      this.open();
    });

    // Close modal
    document.getElementById('settings-close').addEventListener('click', () => {
      this.close();
    });

    document.getElementById('settings-cancel').addEventListener('click', () => {
      this.close();
    });

    // Close on overlay click
    document.querySelector('.settings-modal-overlay').addEventListener('click', () => {
      this.close();
    });

    // Tab switching
    document.querySelectorAll('.settings-tab').forEach(tab => {
      tab.addEventListener('click', (e) => {
        this.switchTab(e.target.dataset.tab);
      });
    });

    // Save settings
    document.getElementById('settings-save').addEventListener('click', () => {
      this.saveSettings();
    });

    // ESC key to close
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && this.isOpen) {
        this.close();
      }
    });
  }

  /**
   * Open settings modal
   */
  open() {
    this.isOpen = true;
    document.getElementById('settings-modal').style.display = 'block';
    document.body.style.overflow = 'hidden';
    this.loadSettings();
  }

  /**
   * Close settings modal
   */
  close() {
    this.isOpen = false;
    document.getElementById('settings-modal').style.display = 'none';
    document.body.style.overflow = '';
  }

  /**
   * Switch between tabs
   */
  switchTab(tabName) {
    this.currentTab = tabName;

    // Update tab buttons
    document.querySelectorAll('.settings-tab').forEach(tab => {
      tab.classList.toggle('active', tab.dataset.tab === tabName);
    });

    // Update tab content
    document.querySelectorAll('.settings-tab-content').forEach(content => {
      content.classList.toggle('active', content.id === `${tabName}-tab`);
    });

    // Load content for the active tab
    switch (tabName) {
      case 'providers':
        this.loadProvidersTab();
        break;
      case 'models':
        this.loadModelsTab();
        break;
      case 'advanced':
        this.loadAdvancedTab();
        break;
    }
  }

  /**
   * Load current settings
   */
  loadSettings() {
    this.loadProvidersTab();
    this.updateProviderStatus();
  }

  /**
   * Load providers tab content
   */
  loadProvidersTab() {
    const settings = this.settingsManager.getSettings();
    const providersContainer = document.getElementById('providers-list');
    
    providersContainer.innerHTML = '';

    Object.entries(this.settingsManager.providers).forEach(([providerId, providerInfo]) => {
      const currentConfig = settings.apiKeys[providerId];
      const isConfigured = this.settingsManager.isProviderConfigured(providerId, currentConfig);

      const providerCard = document.createElement('div');
      providerCard.className = `provider-card ${isConfigured ? 'configured' : ''}`;
      
      const fieldsHTML = providerInfo.fields.map(field => {
        const currentValue = typeof currentConfig === 'string' 
          ? (field.key === 'apiKey' ? currentConfig : '')
          : (currentConfig[field.key] || '');

        const isPassword = field.type === 'password';
        
        return `
          <div class="field-group">
            <label class="field-label">${field.label}</label>
            <div class="${isPassword ? 'toggle-password' : ''}">
              <input 
                type="${field.type}" 
                class="field-input" 
                placeholder="${field.placeholder}"
                data-provider="${providerId}"
                data-field="${field.key}"
                value="${currentValue}"
                ${field.required ? 'required' : ''}
              />
              ${isPassword ? '<button type="button" class="toggle-password-btn">👁️</button>' : ''}
            </div>
          </div>
        `;
      }).join('');

      providerCard.innerHTML = `
        <div class="provider-header">
          <h3 class="provider-name">${providerInfo.name}</h3>
          <span class="provider-status-badge ${isConfigured ? 'configured' : 'unconfigured'}">
            ${isConfigured ? '✓ Configured' : 'Not Configured'}
          </span>
        </div>
        <p class="provider-description">${providerInfo.description}</p>
        <div class="provider-fields">
          ${fieldsHTML}
        </div>
      `;

      providersContainer.appendChild(providerCard);
    });

    // Add password toggle functionality
    document.querySelectorAll('.toggle-password-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const input = e.target.previousElementSibling;
        const type = input.type === 'password' ? 'text' : 'password';
        input.type = type;
        e.target.textContent = type === 'password' ? '👁️' : '🙈';
      });
    });
  }

  /**
   * Load models tab content
   */
  async loadModelsTab() {
    const modelsContainer = document.getElementById('models-config');
    const settings = this.settingsManager.getSettings();
    
    modelsContainer.innerHTML = '<div class="loading-spinner"></div> Loading available models...';

    const configuredProviders = this.settingsManager.getConfiguredProviders();
    
    if (configuredProviders.length === 0) {
      modelsContainer.innerHTML = `
        <div class="settings-description">
          <p>⚠️ No providers configured yet. Please configure at least one AI provider in the API Keys tab first.</p>
        </div>
      `;
      return;
    }

    let modelsHTML = '';

    for (const provider of configuredProviders) {
      const models = await this.settingsManager.getAvailableModels(provider.id);
      const currentModel = settings.models[provider.id];

      modelsHTML += `
        <div class="model-config-group">
          <div class="model-config-header">
            <h4>${provider.name} Models</h4>
          </div>
          <div class="field-group">
            <label class="field-label">Model</label>
            <select class="model-select" data-provider="${provider.id}" data-field="textModel">
              ${models.map(model => 
                `<option value="${model.id}" ${model.id === currentModel.textModel ? 'selected' : ''}>
                  ${model.name} (${model.maxTokens} tokens)
                </option>`
              ).join('')}
            </select>
          </div>
          <div class="parameter-group">
            <div class="field-group">
              <label class="field-label">Max Tokens</label>
              <input type="number" class="field-input" min="1" max="32000" 
                     data-provider="${provider.id}" data-field="maxTokens"
                     value="${currentModel.maxTokens}">
            </div>
            <div class="field-group">
              <label class="field-label">Temperature</label>
              <input type="number" class="field-input" min="0" max="1" step="0.1"
                     data-provider="${provider.id}" data-field="temperature"
                     value="${currentModel.temperature}">
            </div>
          </div>
        </div>
      `;
    }

    modelsContainer.innerHTML = modelsHTML;
  }

  /**
   * Load advanced tab content
   */
  loadAdvancedTab() {
    const advancedContainer = document.getElementById('advanced-options');
    const settings = this.settingsManager.getSettings();

    advancedContainer.innerHTML = `
      <div class="advanced-option">
        <h4>Default Provider</h4>
        <p>Choose which AI provider to use by default when multiple are configured.</p>
        <select class="field-input" id="default-provider">
          ${this.settingsManager.getConfiguredProviders().map(provider => 
            `<option value="${provider.id}" ${provider.id === settings.defaultProvider ? 'selected' : ''}>
              ${provider.name}
            </option>`
          ).join('')}
        </select>
      </div>

      <div class="advanced-option">
        <h4>Feature Settings</h4>
        <p>Enable or disable optional features.</p>
        <label style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
          <input type="checkbox" ${settings.features.enableAutoComplete ? 'checked' : ''} 
                 data-feature="enableAutoComplete">
          Enable auto-complete suggestions
        </label>
        <label style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
          <input type="checkbox" ${settings.features.enableContextAnalysis ? 'checked' : ''} 
                 data-feature="enableContextAnalysis">
          Enable context analysis
        </label>
        <label style="display: flex; align-items: center; gap: 8px;">
          <input type="checkbox" ${settings.features.cacheResponses ? 'checked' : ''} 
                 data-feature="cacheResponses">
          Cache AI responses locally
        </label>
      </div>

      <div class="advanced-option">
        <h4>Data Management</h4>
        <p>Export, import, or clear your settings and data.</p>
        <div style="display: flex; gap: 8px; flex-wrap: wrap;">
          <button class="btn btn-secondary" id="export-settings">Export Settings</button>
          <button class="btn btn-secondary" id="import-settings">Import Settings</button>
          <button class="btn btn-danger" id="clear-settings">Clear All Data</button>
        </div>
        <input type="file" id="import-file" accept=".json" style="display: none;">
      </div>
    `;

    // Add event listeners for advanced options
    document.getElementById('export-settings').addEventListener('click', () => {
      this.exportSettings();
    });

    document.getElementById('import-settings').addEventListener('click', () => {
      document.getElementById('import-file').click();
    });

    document.getElementById('import-file').addEventListener('change', (e) => {
      this.importSettings(e.target.files[0]);
    });

    document.getElementById('clear-settings').addEventListener('click', () => {
      this.clearAllSettings();
    });
  }

  /**
   * Update provider status display
   */
  updateProviderStatus() {
    const statusElement = document.getElementById('provider-status');
    const configuredProviders = this.settingsManager.getConfiguredProviders();
    
    if (configuredProviders.length === 0) {
      statusElement.className = 'settings-status warning';
      statusElement.textContent = '⚠️ No AI providers configured. Please add at least one API key to get started.';
    } else {
      statusElement.className = 'settings-status success';
      statusElement.textContent = `✅ ${configuredProviders.length} provider(s) configured: ${configuredProviders.map(p => p.name).join(', ')}`;
    }
  }

  /**
   * Save all settings
   */
  async saveSettings() {
    const settings = this.settingsManager.getSettings();
    
    // Save provider configurations
    document.querySelectorAll('.provider-card').forEach(card => {
      const inputs = card.querySelectorAll('.field-input');
      inputs.forEach(input => {
        const providerId = input.dataset.provider;
        const fieldKey = input.dataset.field;
        const value = input.value;

        if (typeof settings.apiKeys[providerId] === 'string') {
          // Simple string config (gemini, cerebras)
          if (fieldKey === 'apiKey') {
            settings.apiKeys[providerId] = value;
          }
        } else {
          // Object config (azureOpenAI)
          settings.apiKeys[providerId][fieldKey] = value;
        }
      });
    });

    // Save model configurations
    document.querySelectorAll('[data-provider][data-field]').forEach(input => {
      const providerId = input.dataset.provider;
      const fieldKey = input.dataset.field;
      let value = input.value;

      if (fieldKey === 'maxTokens') {
        value = parseInt(value, 10);
      } else if (fieldKey === 'temperature') {
        value = parseFloat(value);
      }

      if (settings.models[providerId]) {
        settings.models[providerId][fieldKey] = value;
      }
    });

    // Save default provider
    const defaultProviderSelect = document.getElementById('default-provider');
    if (defaultProviderSelect) {
      settings.defaultProvider = defaultProviderSelect.value;
    }

    // Save feature settings
    document.querySelectorAll('[data-feature]').forEach(checkbox => {
      const feature = checkbox.dataset.feature;
      settings.features[feature] = checkbox.checked;
    });

    // Save to localStorage
    const success = this.settingsManager.saveSettings(settings);
    
    if (success) {
      this.showNotification('Settings saved successfully!', 'success');
      this.updateProviderStatus();
      
      // Reinitialize AI service with new settings
      if (window.aiService) {
        await window.aiService.initialize(settings);
      }
    } else {
      this.showNotification('Failed to save settings. Please try again.', 'error');
    }
  }

  /**
   * Export settings to file
   */
  exportSettings() {
    const settingsData = this.settingsManager.exportSettings();
    const blob = new Blob([settingsData], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = `odin-ai-settings-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    this.showNotification('Settings exported successfully!', 'success');
  }

  /**
   * Import settings from file
   */
  async importSettings(file) {
    if (!file) return;
    
    try {
      const text = await file.text();
      const success = this.settingsManager.importSettings(text);
      
      if (success) {
        this.showNotification('Settings imported successfully! API keys were not imported for security.', 'success');
        this.loadSettings();
      } else {
        this.showNotification('Failed to import settings. Please check the file format.', 'error');
      }
    } catch (error) {
      console.error('Import error:', error);
      this.showNotification('Error reading settings file.', 'error');
    }
  }

  /**
   * Clear all settings
   */
  clearAllSettings() {
    if (confirm('Are you sure you want to clear all settings? This action cannot be undone.')) {
      const success = this.settingsManager.clearSettings();
      
      if (success) {
        this.showNotification('All settings cleared successfully!', 'success');
        this.loadSettings();
      } else {
        this.showNotification('Failed to clear settings.', 'error');
      }
    }
  }

  /**
   * Show notification
   */
  showNotification(message, type = 'info') {
    // Create notification element if it doesn't exist
    let notification = document.getElementById('settings-notification');
    if (!notification) {
      notification = document.createElement('div');
      notification.id = 'settings-notification';
      notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 16px;
        border-radius: 6px;
        color: white;
        font-weight: 500;
        z-index: 10001;
        max-width: 300px;
        transition: all 0.3s ease;
        transform: translateX(100%);
      `;
      document.body.appendChild(notification);
    }

    // Set message and type
    notification.textContent = message;
    notification.className = `notification-${type}`;
    
    // Set background color based on type
    const colors = {
      success: '#28a745',
      error: '#dc3545',
      warning: '#ffc107',
      info: '#17a2b8'
    };
    notification.style.background = colors[type] || colors.info;

    // Show notification
    notification.style.transform = 'translateX(0)';

    // Hide after 4 seconds
    setTimeout(() => {
      notification.style.transform = 'translateX(100%)';
    }, 4000);
  }
}

// Initialize settings UI when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    const settingsUI = new SettingsUI();
    settingsUI.init();
    window.settingsUI = settingsUI;
  });
} else {
  const settingsUI = new SettingsUI();
  settingsUI.init();
  window.settingsUI = settingsUI;
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { SettingsUI };
} else {
  window.SettingsUI = SettingsUI;
}
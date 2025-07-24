/**
 * Settings Manager - Handles local storage of API keys and provider configurations
 * Replaces the need for config.js file
 */

class SettingsManager {
  constructor() {
    this.storageKey = 'odin-ai-settings';
    this.modelCache = new Map();
    this.providers = {
      gemini: {
        name: 'Google Gemini',
        description: 'Google\'s most capable AI model',
        fields: [{
          key: 'apiKey',
          label: 'API Key',
          type: 'password',
          placeholder: 'Enter your Gemini API key',
          required: true
        }]
      },
      azureOpenAI: {
        name: 'Azure OpenAI',
        description: 'Microsoft\'s Azure-hosted OpenAI models',
        fields: [
          {
            key: 'apiKey',
            label: 'API Key',
            type: 'password',
            placeholder: 'Enter your Azure OpenAI API key',
            required: true
          },
          {
            key: 'endpoint',
            label: 'Endpoint URL',
            type: 'url',
            placeholder: 'https://your-resource.openai.azure.com/',
            required: true
          },
          {
            key: 'deploymentName',
            label: 'Deployment Name',
            type: 'text',
            placeholder: 'your-deployment-name',
            required: true
          }
        ]
      },
      cerebras: {
        name: 'Cerebras',
        description: 'Ultra-fast AI inference platform',
        fields: [{
          key: 'apiKey',
          label: 'API Key',
          type: 'password',
          placeholder: 'Enter your Cerebras API key',
          required: true
        }]
      }
    };
  }

  /**
   * Get all settings from localStorage
   */
  getSettings() {
    try {
      const settings = localStorage.getItem(this.storageKey);
      return settings ? JSON.parse(settings) : this.getDefaultSettings();
    } catch (error) {
      console.error('Error loading settings:', error);
      return this.getDefaultSettings();
    }
  }

  /**
   * Save settings to localStorage
   */
  saveSettings(settings) {
    try {
      localStorage.setItem(this.storageKey, JSON.stringify(settings));
      return true;
    } catch (error) {
      console.error('Error saving settings:', error);
      return false;
    }
  }

  /**
   * Get default settings structure
   */
  getDefaultSettings() {
    return {
      defaultProvider: 'gemini',
      apiKeys: {
        gemini: '',
        azureOpenAI: {
          apiKey: '',
          endpoint: '',
          deploymentName: ''
        },
        cerebras: ''
      },
      models: {
        gemini: {
          textModel: 'gemini-1.5-flash',
          maxTokens: 8192,
          temperature: 0.7
        },
        azureOpenAI: {
          textModel: 'gpt-4',
          maxTokens: 4096,
          temperature: 0.7
        },
        cerebras: {
          textModel: 'llama3.1-70b',
          maxTokens: 8192,
          temperature: 0.7
        }
      },
      features: {
        enableImageGeneration: false,
        enableAutoComplete: true,
        enableContextAnalysis: true,
        cacheResponses: false
      }
    };
  }

  /**
   * Get configured providers (those with valid API keys)
   */
  getConfiguredProviders() {
    const settings = this.getSettings();
    const configured = [];

    for (const [providerId, config] of Object.entries(settings.apiKeys)) {
      if (this.isProviderConfigured(providerId, config)) {
        configured.push({
          id: providerId,
          name: this.providers[providerId]?.name || providerId,
          ...this.providers[providerId]
        });
      }
    }

    return configured;
  }

  /**
   * Check if a provider is properly configured
   */
  isProviderConfigured(providerId, config) {
    if (!config) return false;

    switch (providerId) {
      case 'gemini':
      case 'cerebras':
        return typeof config === 'string' && config.trim() !== '';
      
      case 'azureOpenAI':
        return config.apiKey && config.apiKey.trim() !== '' &&
               config.endpoint && config.endpoint.trim() !== '' &&
               config.deploymentName && config.deploymentName.trim() !== '';
      
      default:
        return false;
    }
  }

  /**
   * Get available models for a provider
   */
  async getAvailableModels(providerId) {
    // Check cache first
    if (this.modelCache.has(providerId)) {
      return this.modelCache.get(providerId);
    }

    const settings = this.getSettings();
    const config = settings.apiKeys[providerId];

    if (!this.isProviderConfigured(providerId, config)) {
      return [];
    }

    try {
      let models = [];
      
      switch (providerId) {
        case 'gemini':
          models = await this.fetchGeminiModels(config);
          break;
        case 'azureOpenAI':
          models = await this.fetchAzureOpenAIModels(config);
          break;
        case 'cerebras':
          models = await this.fetchCerebrasModels(config);
          break;
      }

      // Cache for 1 hour
      this.modelCache.set(providerId, models);
      setTimeout(() => this.modelCache.delete(providerId), 60 * 60 * 1000);

      return models;
    } catch (error) {
      console.error(`Error fetching models for ${providerId}:`, error);
      return this.getFallbackModels(providerId);
    }
  }

  /**
   * Fetch Gemini models
   */
  async fetchGeminiModels(apiKey) {
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
    
    if (!response.ok) {
      throw new Error(`Gemini API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.models
      .filter(model => model.supportedGenerationMethods?.includes('generateContent'))
      .map(model => ({
        id: model.name.replace('models/', ''),
        name: model.displayName || model.name,
        maxTokens: model.outputTokenLimit || 8192,
        description: model.description
      }));
  }

  /**
   * Fetch Azure OpenAI models
   */
  async fetchAzureOpenAIModels(config) {
    const response = await fetch(`${config.endpoint}openai/deployments?api-version=2024-02-15-preview`, {
      headers: {
        'api-key': config.apiKey
      }
    });

    if (!response.ok) {
      throw new Error(`Azure OpenAI API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.data.map(deployment => ({
      id: deployment.id,
      name: deployment.model,
      maxTokens: this.getModelMaxTokens(deployment.model),
      description: `Deployment: ${deployment.id}`
    }));
  }

  /**
   * Fetch Cerebras models
   */
  async fetchCerebrasModels(apiKey) {
    const response = await fetch('https://api.cerebras.ai/v1/models', {
      headers: {
        'Authorization': `Bearer ${apiKey}`
      }
    });

    if (!response.ok) {
      throw new Error(`Cerebras API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.data.map(model => ({
      id: model.id,
      name: model.id,
      maxTokens: model.max_tokens || 8192,
      description: model.description || 'Cerebras AI model'
    }));
  }

  /**
   * Get fallback models if API call fails
   */
  getFallbackModels(providerId) {
    const fallbacks = {
      gemini: [
        { id: 'gemini-1.5-flash', name: 'Gemini 1.5 Flash', maxTokens: 8192, description: 'Fast and efficient model' },
        { id: 'gemini-1.5-pro', name: 'Gemini 1.5 Pro', maxTokens: 8192, description: 'Most capable model' }
      ],
      azureOpenAI: [
        { id: 'gpt-4', name: 'GPT-4', maxTokens: 4096, description: 'Most capable GPT model' },
        { id: 'gpt-3.5-turbo', name: 'GPT-3.5 Turbo', maxTokens: 4096, description: 'Fast and efficient' }
      ],
      cerebras: [
        { id: 'llama3.1-70b', name: 'Llama 3.1 70B', maxTokens: 8192, description: 'High-performance model' },
        { id: 'llama3.1-8b', name: 'Llama 3.1 8B', maxTokens: 8192, description: 'Fast inference model' }
      ]
    };

    return fallbacks[providerId] || [];
  }

  /**
   * Get approximate max tokens for known models
   */
  getModelMaxTokens(modelName) {
    const tokenLimits = {
      'gpt-4': 4096,
      'gpt-4-turbo': 4096,
      'gpt-3.5-turbo': 4096,
      'gemini-1.5-flash': 8192,
      'gemini-1.5-pro': 8192,
      'llama3.1-70b': 8192,
      'llama3.1-8b': 8192
    };

    return tokenLimits[modelName] || 4096;
  }

  /**
   * Update provider configuration
   */
  updateProvider(providerId, config) {
    const settings = this.getSettings();
    settings.apiKeys[providerId] = config;
    
    // Clear model cache for this provider
    this.modelCache.delete(providerId);
    
    return this.saveSettings(settings);
  }

  /**
   * Set default provider
   */
  setDefaultProvider(providerId) {
    const settings = this.getSettings();
    settings.defaultProvider = providerId;
    return this.saveSettings(settings);
  }

  /**
   * Clear all settings
   */
  clearSettings() {
    try {
      localStorage.removeItem(this.storageKey);
      this.modelCache.clear();
      return true;
    } catch (error) {
      console.error('Error clearing settings:', error);
      return false;
    }
  }

  /**
   * Export settings for backup
   */
  exportSettings() {
    const settings = this.getSettings();
    // Remove sensitive data from export
    const exportData = {
      ...settings,
      apiKeys: Object.keys(settings.apiKeys).reduce((acc, key) => {
        acc[key] = '[REDACTED]';
        return acc;
      }, {})
    };
    
    return JSON.stringify(exportData, null, 2);
  }

  /**
   * Import settings from backup (excluding API keys for security)
   */
  importSettings(settingsJson) {
    try {
      const importedSettings = JSON.parse(settingsJson);
      const currentSettings = this.getSettings();
      
      // Only import non-sensitive settings
      const safeSettings = {
        ...currentSettings,
        defaultProvider: importedSettings.defaultProvider || currentSettings.defaultProvider,
        models: importedSettings.models || currentSettings.models,
        features: importedSettings.features || currentSettings.features
      };
      
      return this.saveSettings(safeSettings);
    } catch (error) {
      console.error('Error importing settings:', error);
      return false;
    }
  }
}

// Create global instance
const settingsManager = new SettingsManager();

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { SettingsManager, settingsManager };
} else {
  window.settingsManager = settingsManager;
  window.SettingsManager = SettingsManager;
}
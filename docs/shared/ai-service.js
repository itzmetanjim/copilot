/**
 * AI Service Module - Handles interactions with multiple AI providers
 * Supports Gemini, Azure OpenAI, and Cerebras APIs
 */

class AIService {
  constructor() {
    this.config = null;
    this.currentProvider = null;
    this.initialized = false;
    this.rateLimitManager = new Map();
  }

  /**
   * Initialize the AI service with configuration
   */
  async initialize(config) {
    try {
      // Use provided config, settings manager, or fallback to window.CONFIG
      if (config) {
        this.config = config;
      } else if (window.settingsManager) {
        this.config = window.settingsManager.getSettings();
      } else if (window.CONFIG) {
        this.config = window.CONFIG;
      } else {
        throw new Error('No configuration available. Please configure your AI providers in Settings.');
      }
      
      // Validate that at least one provider is configured
      const configuredProviders = this.getConfiguredProviders();
      if (configuredProviders.length === 0) {
        throw new Error('No AI providers configured. Please add at least one API key in Settings.');
      }
      
      this.currentProvider = this.config.defaultProvider;
      
      // If default provider is not configured, use the first configured one
      if (!this.isProviderConfigured(this.currentProvider)) {
        this.currentProvider = configuredProviders[0];
        console.log(`Default provider not configured, switching to: ${this.currentProvider}`);
      }
      
      this.initialized = true;
      
      console.log(`AI Service initialized with provider: ${this.currentProvider}`);
      return true;
    } catch (error) {
      console.error('Failed to initialize AI Service:', error);
      this.showError('Failed to initialize AI service. Please configure your providers in Settings.');
      return false;
    }
  }

  /**
   * Generate text based on a prompt
   */
  async generateText(prompt, options = {}) {
    if (!this.initialized) {
      throw new Error('AI Service not initialized');
    }

    const provider = options.provider || this.currentProvider;
    const maxTokens = options.maxTokens || this.config.models[provider].maxTokens;
    const temperature = options.temperature || this.config.models[provider].temperature;

    try {
      this.showLoading('Generating content...');
      
      let response;
      switch (provider) {
        case 'gemini':
          response = await this.callGemini(prompt, { maxTokens, temperature });
          break;
        case 'azure-openai':
          response = await this.callAzureOpenAI(prompt, { maxTokens, temperature });
          break;
        case 'cerebras':
          response = await this.callCerebras(prompt, { maxTokens, temperature });
          break;
        default:
          throw new Error(`Unsupported provider: ${provider}`);
      }
      
      this.hideLoading();
      return response;
    } catch (error) {
      this.hideLoading();
      this.handleError(error, 'text generation');
      throw error;
    }
  }

  /**
   * Summarize text content
   */
  async summarizeText(text, options = {}) {
    const length = options.length || 'medium'; // short, medium, long
    const format = options.format || 'paragraph'; // paragraph, bullets, outline
    
    const prompt = `Please summarize the following text in a ${length} ${format} format:\n\n${text}`;
    return await this.generateText(prompt, options);
  }

  /**
   * Rewrite text with specific tone or style
   */
  async rewriteText(text, tone = 'professional', options = {}) {
    const prompt = `Please rewrite the following text in a ${tone} tone while maintaining the original meaning:\n\n${text}`;
    return await this.generateText(prompt, options);
  }

  /**
   * Analyze data and answer questions
   */
  async analyzeData(data, query, options = {}) {
    const dataStr = typeof data === 'object' ? JSON.stringify(data, null, 2) : data;
    const prompt = `Based on the following data, please answer this question: ${query}\n\nData:\n${dataStr}`;
    return await this.generateText(prompt, options);
  }

  /**
   * Generate images (if supported by provider)
   */
  async generateImage(prompt, options = {}) {
    if (!this.config.features.enableImageGeneration) {
      throw new Error('Image generation is not enabled for the current provider');
    }
    
    // Implementation would depend on the specific provider's image generation API
    // For now, return a placeholder
    return {
      success: false,
      message: 'Image generation not yet implemented for this provider'
    };
  }

  /**
   * Call Gemini API
   */
  async callGemini(prompt, options) {
    const apiKey = this.config.apiKeys.gemini;
    if (!this.isProviderConfigured('gemini')) {
      throw new Error('Gemini API key not configured. Please add your API key in Settings.');
    }

    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${this.config.models.gemini.textModel}:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        contents: [{
          parts: [{
            text: prompt
          }]
        }],
        generationConfig: {
          maxOutputTokens: options.maxTokens,
          temperature: options.temperature
        }
      })
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Gemini API error: ${error.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data.candidates[0]?.content?.parts[0]?.text || 'No response generated';
  }

  /**
   * Call Azure OpenAI API
   */
  async callAzureOpenAI(prompt, options) {
    const config = this.config.apiKeys.azureOpenAI;
    if (!this.isProviderConfigured('azureOpenAI')) {
      throw new Error('Azure OpenAI not configured. Please add your API key, endpoint, and deployment name in Settings.');
    }

    const response = await fetch(`${config.endpoint}openai/deployments/${config.deploymentName}/chat/completions?api-version=2024-02-15-preview`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': config.apiKey
      },
      body: JSON.stringify({
        messages: [
          { role: 'user', content: prompt }
        ],
        max_tokens: options.maxTokens,
        temperature: options.temperature
      })
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Azure OpenAI API error: ${error.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0]?.message?.content || 'No response generated';
  }

  /**
   * Call Cerebras API
   */
  async callCerebras(prompt, options) {
    const apiKey = this.config.apiKeys.cerebras;
    if (!this.isProviderConfigured('cerebras')) {
      throw new Error('Cerebras API key not configured. Please add your API key in Settings.');
    }

    const response = await fetch('https://api.cerebras.ai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: this.config.models.cerebras.textModel,
        messages: [
          { role: 'user', content: prompt }
        ],
        max_tokens: options.maxTokens,
        temperature: options.temperature
      })
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(`Cerebras API error: ${error.error?.message || response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0]?.message?.content || 'No response generated';
  }

  /**
   * Get configured providers
   */
  getConfiguredProviders() {
    if (!this.config) return [];
    
    const configured = [];
    for (const [providerId, config] of Object.entries(this.config.apiKeys)) {
      if (this.isProviderConfigured(providerId, config)) {
        configured.push(providerId);
      }
    }
    return configured;
  }

  /**
   * Switch AI provider
   */
  switchProvider(provider) {
    if (!this.isProviderConfigured(provider)) {
      throw new Error(`Provider ${provider} not configured`);
    }
    this.currentProvider = provider;
    console.log(`Switched to AI provider: ${provider}`);
  }

  /**
   * Refresh configuration from settings manager
   */
  refreshConfig() {
    if (window.settingsManager) {
      this.config = window.settingsManager.getSettings();
      
      // Update current provider if it's no longer configured
      if (!this.isProviderConfigured(this.currentProvider)) {
        const configuredProviders = this.getConfiguredProviders();
        if (configuredProviders.length > 0) {
          this.currentProvider = configuredProviders[0];
          console.log(`Switched to available provider: ${this.currentProvider}`);
        } else {
          this.initialized = false;
          console.warn('No providers configured');
        }
      }
    }
  }

  /**
   * Handle errors gracefully
   */
  handleError(error, context) {
    console.error(`AI Service error in ${context}:`, error);
    
    let userMessage = 'An error occurred while processing your request.';
    
    if (error.message.includes('not configured') || error.message.includes('API key')) {
      userMessage = 'AI provider not configured. Please configure your API keys in Settings (⚙️ button).';
    } else if (error.message.includes('rate limit') || error.message.includes('quota')) {
      userMessage = 'Rate limit exceeded. Please try again in a moment.';
    } else if (error.message.includes('network') || error.message.includes('fetch')) {
      userMessage = 'Network error. Please check your internet connection.';
    }
    
    this.showError(userMessage);
  }

  /**
   * Show loading indicator
   */
  showLoading(message = 'Processing...') {
    const loadingElement = document.getElementById('ai-loading');
    if (loadingElement) {
      loadingElement.textContent = message;
      loadingElement.style.display = 'block';
    }
  }

  /**
   * Hide loading indicator
   */
  hideLoading() {
    const loadingElement = document.getElementById('ai-loading');
    if (loadingElement) {
      loadingElement.style.display = 'none';
    }
  }

  /**
   * Show error message
   */
  showError(message) {
    const errorElement = document.getElementById('ai-error');
    if (errorElement) {
      errorElement.textContent = message;
      errorElement.style.display = 'block';
      
      // Auto-hide after 5 seconds
      setTimeout(() => {
        errorElement.style.display = 'none';
      }, 5000);
    } else {
      // Fallback to alert if no error element
      alert(`Error: ${message}`);
    }
  }

  /**
   * Show success message
   */
  showSuccess(message) {
    const successElement = document.getElementById('ai-success');
    if (successElement) {
      successElement.textContent = message;
      successElement.style.display = 'block';
      
      // Auto-hide after 3 seconds
      setTimeout(() => {
        successElement.style.display = 'none';
      }, 3000);
    }
  }
}

// Create global instance
const aiService = new AIService();

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { AIService, aiService };
} else {
  window.aiService = aiService;
  window.AIService = AIService;
}
// Example configuration file - Copy this to config.js and fill in your API keys
// DO NOT commit config.js to version control

const config = {
  // Choose your preferred AI provider: 'gemini', 'azure-openai', or 'cerebras'
  defaultProvider: 'gemini',
  
  // API Keys - Get these from your respective AI providers
  apiKeys: {
    gemini: 'your-gemini-api-key-here',
    azureOpenAI: {
      apiKey: 'your-azure-openai-api-key-here',
      endpoint: 'https://your-resource-name.openai.azure.com/',
      deploymentName: 'your-deployment-name'
    },
    cerebras: 'your-cerebras-api-key-here'
  },
  
  // Model configurations
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
  
  // Feature flags
  features: {
    enableImageGeneration: false, // Set to true if your provider supports image generation
    enableAutoComplete: true,
    enableContextAnalysis: true,
    cacheResponses: false
  }
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = config;
} else {
  window.CONFIG = config;
}
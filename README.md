# 🧙‍♂️ Odin AI Assistant Suite

> **"Odin" - An Open-Source AI Assistant Suite for Microsoft 365**

A comprehensive, production-ready alternative to Microsoft 365 Copilot that integrates with Gemini, Azure OpenAI, and Cerebras APIs to deliver powerful AI capabilities across Word, Excel, PowerPoint, and Outlook.

## 🌟 Features Overview

Odin provides **full feature parity** with Microsoft 365 Copilot, including all the advanced features and intelligent context awareness you expect from a world-class AI assistant.

### 📝 Word Assistant
- **✨ Draft with Odin**: Generate complete documents, paragraphs, outlines, and lists
- **📄 Smart Summarization**: Analyze entire documents or selected text with multiple format options
- **✏️ Rewrite & Coach**: Improve text clarity, tone, and style with AI coaching
- **🎨 Advanced Formatting**: Smart document structure and style suggestions
- **🔍 Document Analysis**: Comprehensive content analysis and improvement recommendations
- **✅ Grammar & Style Check**: Professional writing assistance and error correction

### 📊 Excel Assistant  
- **🧮 Formula Generation**: Create complex Excel formulas from natural language descriptions
- **📈 Data Insights**: Advanced statistical analysis and trend identification
- **📊 Chart Creation**: Generate professional charts and visualizations from data
- **🎨 Smart Formatting**: Intelligent data formatting, conditional formatting, and table styling
- **🔄 PivotTable Assistant**: AI-guided PivotTable creation and analysis
- **🎯 What-If Analysis**: Scenario planning and goal seek assistance
- **✅ Data Validation**: Intelligent data validation rules and quality checks

### 🎯 PowerPoint Assistant
- **🎨 Presentation Creation**: Generate complete presentations from topics with structured outlines
- **➕ Smart Slide Generation**: Create individual slides with intelligent layouts and content
- **🗣️ Speaker Notes**: Detailed, natural-sounding speaker notes for any presentation
- **🎨 Design Assistant**: Professional design suggestions, layout improvements, and consistency checks
- **📋 Content Structuring**: Intelligent outline generation and content organization
- **🌐 Advanced Features**: Translation, accessibility checks, and content enhancement

### 📧 Outlook Assistant
- **✉️ Email Drafting**: Generate professional emails from simple prompts
- **📬 Thread Summarization**: Intelligent email conversation analysis and summaries
- **💡 Email Coaching**: Tone improvement, clarity enhancement, and communication optimization
- **⚡ Smart Replies**: Context-aware response suggestions
- **📅 Meeting Integration**: Meeting summary generation and follow-up assistance

## 🏗️ Architecture

### Monorepo Structure
```
odin-ai-assistant-suite/
├── shared/                 # Shared components and utilities
│   ├── ai-service.js      # Multi-provider AI service layer
│   ├── ui-components.js   # Reusable UI components
│   └── styles.css         # Consistent theming and styles
├── word-addin/            # Word add-in manifest
├── excel-addin/           # Excel add-in manifest  
├── powerpoint-addin/      # PowerPoint add-in manifest
├── outlook-addin/         # Outlook add-in manifest
├── docs/                  # GitHub Pages webroot
│   ├── word/             # Word add-in web app
│   ├── excel/            # Excel add-in web app
│   ├── powerpoint/       # PowerPoint add-in web app
│   ├── outlook/          # Outlook add-in web app
│   └── shared/           # Shared web resources
├── config.example.js      # Configuration template
└── README.md             # This file
```

### AI Provider Support
- **Google Gemini**: Excellent free tier with high-quality responses
- **Cerebras**: Lightning-fast inference with generous free limits
- **Azure OpenAI**: Enterprise-grade reliability (free $100 credit for students)

### Web Architecture
All add-ins are designed as client-side web applications that can be hosted directly from GitHub Pages, making deployment and distribution seamless.

## 🚀 Quick Start

### Prerequisites
- Microsoft Office 365 or Office 2019+
- API keys for at least one AI provider:
  - [Google Gemini API](https://ai.google.dev/)
  - [Cerebras API](https://cerebras.ai/)
  - [Azure OpenAI](https://azure.microsoft.com/en-us/products/ai-services/openai-service)

### Installation Steps

1. **Deploy to GitHub Pages**
   - Fork this repository or use it directly
   - Go to repository Settings → Pages
   - Set source to "Deploy from a branch" and select `main` branch with `/docs` folder
   - Your add-ins will be available at `https://yourusername.github.io/copilot/`

2. **Sideload the Add-ins**
   
   **Option A: Direct Sideloading**
   - Download the manifest files from this repository:
     - [Word Manifest](word-addin/word-manifest.xml)
     - [Excel Manifest](excel-addin/excel-manifest.xml) 
     - [PowerPoint Manifest](powerpoint-addin/powerpoint-manifest.xml)
     - [Outlook Manifest](outlook-addin/outlook-manifest.xml)
   - Follow Microsoft's [sideloading instructions](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)

   **Option B: Use from GitHub Pages**
   - The manifests in this repo already point to `https://itzmetanjim.github.io/copilot/`
   - You can use them directly without forking if you just want to try the add-ins

3. **Configure Your AI Providers**
   
   ⚙️ **No more config files needed!** Everything is now managed through the settings UI.
   
   - Open any of the add-ins (Word, Excel, PowerPoint, or Outlook)
   - Click the **⚙️ Settings** button in the top-right corner
   - Go to the **🔑 API Keys** tab
   - Add your API key for at least one provider:

   **Available Providers:**
   - **Google Gemini** (Recommended) - Free tier available
   - **Azure OpenAI** - Requires Azure subscription  
   - **Cerebras** - Fast inference platform

   - Save your settings and start using AI features!

4. **Test Your Setup**
   - Visit `/docs/test-settings.html` in your deployment to verify everything works
   - Try generating some content with your configured AI provider

### ✨ Key Advantages of New Setup

- **🚫 No Config Files**: No need to edit or create configuration files
- **🔐 Secure Storage**: API keys stored locally in your browser's secure storage
- **🔄 Easy Switching**: Change providers or update keys anytime through the UI
- **📱 Works Anywhere**: Use the same settings across all Office applications
- **🛡️ Privacy First**: Your API keys never leave your browser
   - Select "main" branch and "/docs" folder
   - Save the settings

## 🔧 Configuration Options

### Settings UI Overview

All configuration is now managed through an intuitive settings interface accessible via the ⚙️ button in any add-in:

- **🔑 API Keys Tab**: Configure your AI provider credentials
- **🤖 Models Tab**: Select models and adjust parameters  
- **⚙️ Advanced Tab**: Feature toggles and data management

### AI Provider Settings

The following settings are configurable through the UI:

```javascript
// Default provider selection
defaultProvider: 'gemini' // 'gemini', 'azureOpenAI', or 'cerebras'

// Model configurations (dynamically fetched from APIs)
models: {
  gemini: {
    textModel: 'gemini-1.5-flash',  // Auto-populated from API
    maxTokens: 8192,                // Adjustable in UI
    temperature: 0.7                // Adjustable in UI
  },
  azureOpenAI: {
    textModel: 'gpt-4',            // Based on your deployment
    maxTokens: 4096,               // Adjustable in UI
    temperature: 0.7               // Adjustable in UI
  },
  cerebras: {
    textModel: 'llama3.1-70b',     // Auto-populated from API
    maxTokens: 8192,               // Adjustable in UI
    temperature: 0.7               // Adjustable in UI
  }
}
```

### Legacy Configuration (Optional)

For advanced users or developers, `config.example.js` is still available as a reference. However, the UI-based settings system is recommended for all users as it provides:

- ✅ Real-time model fetching from provider APIs
- ✅ Secure local storage of credentials  
- ✅ Easy provider switching
- ✅ No file editing required

### Feature Flags
```javascript
features: {
  enableImageGeneration: false,
  enableAutoComplete: true,
  enableContextAnalysis: true,
  cacheResponses: false
}
```

## 🎯 Key Features Detail

### Advanced AI Capabilities
- **Context-Aware Responses**: Understands document structure and content context
- **Multi-Modal Intelligence**: Processes text, data, and document structure
- **Intelligent Autocomplete**: Predictive text and formula suggestions
- **Error Handling**: Robust error management with user-friendly messages
- **Provider Switching**: Seamless switching between AI providers

### Office Integration
- **Deep Office.js Integration**: Full access to Office document APIs
- **Real-time Collaboration**: Works with Office 365 collaboration features
- **Cross-Platform**: Compatible with Office on Windows, Mac, and web
- **Security**: Client-side processing with secure API communication

### User Experience
- **Consistent UI**: Unified design language across all applications
- **Keyboard Shortcuts**: Productivity-focused hotkey support
- **Responsive Design**: Adapts to different screen sizes and orientations
- **Accessibility**: WCAG-compliant design for all users

## 📚 API Key Setup Guides

### Google Gemini API
1. Visit [Google AI Studio](https://ai.google.dev/)
2. Create a new project or select existing
3. Generate an API key
4. Copy the key to your `config.js`

### Cerebras API
1. Sign up at [Cerebras Cloud](https://cerebras.ai/)
2. Navigate to the API section
3. Generate your API key
4. Copy the key to your `config.js`

### Azure OpenAI
1. Create an [Azure account](https://azure.microsoft.com/) (students get $100 free)
2. Create an OpenAI resource
3. Deploy a model (e.g., GPT-4)
4. Get your endpoint URL, API key, and deployment name
5. Add all details to your `config.js`

## 🔒 Security & Privacy

- **Client-Side Processing**: All AI requests are made directly from your browser
- **No Data Storage**: No user data is stored on external servers
- **API Key Security**: Keys are stored locally and never transmitted except to official AI providers
- **HTTPS Required**: All communications use secure HTTPS connections

## 🛠️ Development

### Local Development
```bash
# Start development server (optional)
npm run dev

# Build for production
npm run build
```

### Testing
```bash
# Run tests
npm test
```

### Contributing
1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make your changes
4. Test thoroughly with all Office applications
5. Submit a pull request

## 📋 System Requirements

### Minimum Requirements
- **Office**: Microsoft 365, Office 2019, or Office 2021
- **Browser**: Chrome 70+, Firefox 65+, Safari 12+, Edge 79+
- **Internet**: Stable connection for AI provider APIs

### Recommended
- **Office**: Microsoft 365 (latest)
- **RAM**: 8GB+ for optimal performance
- **Storage**: 100MB for add-in files

## 🆘 Troubleshooting

### Common Issues

**Add-in not loading?**
- Check internet connection
- Verify manifest file paths
- Ensure GitHub Pages is enabled
- Check browser console for errors

**AI not responding?**
- Verify API keys in `config.js`
- Check API provider status
- Confirm API quota limits
- Try switching providers

**Features not working?**
- Update to latest Office version
- Clear browser cache
- Reload the add-in
- Check Office.js compatibility

### Getting Help
- 📖 [Documentation](https://github.com/itzmetanjim/copilot/wiki)
- 🐛 [Report Issues](https://github.com/itzmetanjim/copilot/issues)
- 💬 [Discussions](https://github.com/itzmetanjim/copilot/discussions)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Microsoft Office JavaScript APIs
- Google Gemini AI
- Cerebras AI Platform
- Azure OpenAI Service
- The open-source community

## 🚀 What's Next

- **Enhanced Image Support**: AI-powered image generation and analysis
- **Advanced Analytics**: Detailed usage and performance insights
- **Template Library**: Pre-built templates for common use cases
- **Plugin System**: Extensible architecture for custom features
- **Enterprise Features**: Advanced security and management capabilities

---

**Built with ❤️ by the Odin AI Team**

*Making AI accessible to everyone, one Office document at a time.*
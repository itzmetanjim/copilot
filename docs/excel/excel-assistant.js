/**
 * Odin Excel Assistant - Main JavaScript
 * Implements all Excel-specific AI features for data analysis, formula generation, and chart creation
 */

// Global variables
let currentSection = null;
let selectedRange = null;
let workbookData = null;
let selectedChartType = 'column';

// Initialize when Office is ready
Office.onReady((info) => {
    console.log('Office.js ready for Excel');
    initializeExcelAssistant();
});

/**
 * Initialize the Excel assistant
 */
async function initializeExcelAssistant() {
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

        // Set default chart type
        selectChartType('column');

        console.log('Excel Assistant initialized successfully');
    } catch (error) {
        console.error('Failed to initialize Excel Assistant:', error);
        aiService.showError('Failed to initialize. Please check your configuration.');
    }
}

/**
 * Setup UI components and event listeners
 */
function setupUI() {
    // Create action buttons for each response area
    const sections = ['formula', 'analysis', 'chart', 'format'];
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
    // Ctrl+Shift+F: Formula generator
    if (event.ctrlKey && event.shiftKey && event.key === 'F') {
        event.preventDefault();
        showFormulaSection();
    }
    
    // Ctrl+Shift+A: Data analysis
    if (event.ctrlKey && event.shiftKey && event.key === 'A') {
        event.preventDefault();
        showAnalysisSection();
    }
    
    // Ctrl+Shift+C: Create chart
    if (event.ctrlKey && event.shiftKey && event.key === 'C') {
        event.preventDefault();
        showChartSection();
    }
}

/**
 * Show configuration help
 */
function showConfigurationHelp() {
    const helpContent = `
        <div class="odin-alert odin-alert-warning">
            <h3>Configuration Required</h3>
            <p>To use Odin Excel Assistant, you need to configure your AI provider API keys:</p>
            <ol>
                <li>Copy <code>config.example.js</code> to <code>config.js</code></li>
                <li>Add your API keys for Gemini, Azure OpenAI, or Cerebras</li>
                <li>Choose your default provider</li>
                <li>Reload the add-in</li>
            </ol>
            <p>Visit the <a href="https://github.com/itzmetanjim/copilot" target="_blank">GitHub repository</a> for detailed setup instructions.</p>
        </div>
    `;
    
    document.querySelector('.odin-container').innerHTML = helpContent;
}

// === SECTION MANAGEMENT ===

/**
 * Show the formula generation section
 */
function showFormulaSection() {
    hideAllSections();
    document.getElementById('formula-section').style.display = 'block';
    currentSection = 'formula';
    document.getElementById('formula-prompt').focus();
}

/**
 * Show the data analysis section
 */
function showAnalysisSection() {
    hideAllSections();
    document.getElementById('analysis-section').style.display = 'block';
    currentSection = 'analysis';
}

/**
 * Show the chart creation section
 */
function showChartSection() {
    hideAllSections();
    document.getElementById('chart-section').style.display = 'block';
    currentSection = 'chart';
}

/**
 * Show the formatting section
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
    const sections = ['formula-section', 'analysis-section', 'chart-section', 'format-section', 'settings-panel'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            section.style.display = 'none';
        }
    });
    currentSection = null;
}

// === FORMULA GENERATION ===

/**
 * Generate Excel formula from natural language description
 */
async function generateFormula() {
    const prompt = document.getElementById('formula-prompt').value.trim();
    const formulaType = document.getElementById('formula-type').value;
    const targetRange = document.getElementById('target-range').value.trim();
    
    if (!prompt) {
        aiService.showError('Please describe what you want to calculate.');
        return;
    }
    
    try {
        const settings = odinUI.getSettings();
        
        // Get current worksheet structure for context
        const worksheetInfo = await getWorksheetContext();
        
        const enhancedPrompt = `Generate an Excel formula for the following request: "${prompt}"
        
        ${formulaType !== 'any' ? `Focus on ${formulaType} functions.` : ''}
        ${targetRange ? `Target cell/range: ${targetRange}` : ''}
        
        Worksheet context:
        ${worksheetInfo}
        
        Provide:
        1. The complete formula (starting with =)
        2. Brief explanation of how it works
        3. Any assumptions made about the data structure
        
        Formula:`;
        
        const response = await aiService.generateText(enhancedPrompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3 // Lower temperature for more precise formulas
        });
        
        // Extract and display the formula
        const formulaMatch = response.match(/=[\w\s\(\),:.!"'$\-+*/]+/);
        if (formulaMatch) {
            const formula = formulaMatch[0];
            document.getElementById('formula-result').innerHTML = `<strong>Formula:</strong> ${formula}`;
            document.getElementById('formula-result').classList.remove('hidden');
        }
        
        odinUI.showResponse('formula-response', response, true);
        aiService.showSuccess('Formula generated successfully!');
    } catch (error) {
        console.error('Formula generation error:', error);
        aiService.showError('Failed to generate formula. Please try again.');
    }
}

/**
 * Get current cell for range input
 */
async function selectCurrentCell() {
    try {
        const address = await getCurrentCellAddress();
        document.getElementById('target-range').value = address;
    } catch (error) {
        console.error('Error getting current cell:', error);
        aiService.showError('Failed to get current cell address.');
    }
}

// === DATA ANALYSIS ===

/**
 * Select data range for analysis
 */
async function selectDataRange() {
    try {
        const range = await getSelectedRange();
        if (range) {
            document.getElementById('analysis-range').value = range.address;
            selectedRange = range;
            
            // Show preview of data
            const preview = await getDataPreview(range.address);
            const previewElement = document.getElementById('data-preview');
            previewElement.innerHTML = preview;
            previewElement.classList.remove('hidden');
        } else {
            aiService.showError('Please select a range in the worksheet first.');
        }
    } catch (error) {
        console.error('Error selecting data range:', error);
        aiService.showError('Failed to select data range.');
    }
}

/**
 * Auto-detect data range
 */
async function detectDataRange() {
    try {
        const range = await autoDetectDataRange();
        if (range) {
            document.getElementById('analysis-range').value = range;
            
            // Show preview of data
            const preview = await getDataPreview(range);
            const previewElement = document.getElementById('data-preview');
            previewElement.innerHTML = preview;
            previewElement.classList.remove('hidden');
        } else {
            aiService.showError('Could not detect data range. Please select manually.');
        }
    } catch (error) {
        console.error('Error auto-detecting range:', error);
        aiService.showError('Failed to auto-detect data range.');
    }
}

/**
 * Perform quick analysis
 */
async function performQuickAnalysis(type) {
    const rangeAddress = document.getElementById('analysis-range').value.trim();
    
    if (!rangeAddress) {
        aiService.showError('Please select a data range first.');
        return;
    }
    
    try {
        const data = await getRangeData(rangeAddress);
        if (!data || data.length === 0) {
            aiService.showError('No data found in the specified range.');
            return;
        }
        
        let prompt = '';
        switch (type) {
            case 'summary':
                prompt = `Provide a comprehensive statistical summary of this data including mean, median, mode, standard deviation, min, max, and count:`;
                break;
            case 'trends':
                prompt = `Analyze this data for trends, patterns, and significant changes. Identify any notable increases, decreases, or cyclic patterns:`;
                break;
            case 'outliers':
                prompt = `Identify outliers and anomalies in this data. Explain which values are unusual and why:`;
                break;
            case 'correlations':
                prompt = `Analyze correlations between different columns in this data. Identify strong positive or negative relationships:`;
                break;
        }
        
        await performDataAnalysis(data, prompt, rangeAddress);
    } catch (error) {
        console.error('Quick analysis error:', error);
        aiService.showError('Failed to perform analysis.');
    }
}

/**
 * Analyze data with custom question
 */
async function analyzeData() {
    const rangeAddress = document.getElementById('analysis-range').value.trim();
    const question = document.getElementById('analysis-question').value.trim();
    
    if (!rangeAddress) {
        aiService.showError('Please select a data range first.');
        return;
    }
    
    if (!question) {
        aiService.showError('Please enter an analysis question.');
        return;
    }
    
    try {
        const data = await getRangeData(rangeAddress);
        if (!data || data.length === 0) {
            aiService.showError('No data found in the specified range.');
            return;
        }
        
        await performDataAnalysis(data, question, rangeAddress);
    } catch (error) {
        console.error('Data analysis error:', error);
        aiService.showError('Failed to analyze data.');
    }
}

/**
 * Perform data analysis with AI
 */
async function performDataAnalysis(data, prompt, rangeAddress) {
    try {
        const settings = odinUI.getSettings();
        
        // Convert data to a readable format
        const dataStr = data.map(row => row.join('\t')).join('\n');
        
        const enhancedPrompt = `${prompt}

Data from range ${rangeAddress}:
${dataStr}

Please provide:
1. Direct answer to the question/analysis
2. Key insights and findings
3. Relevant statistics or calculations
4. Recommendations based on the data`;
        
        const response = await aiService.analyzeData(dataStr, prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        // Create insights summary
        generateInsightsSummary(data, rangeAddress);
        
        odinUI.showResponse('analysis-response', response, true);
        aiService.showSuccess('Data analysis complete!');
    } catch (error) {
        throw error;
    }
}

/**
 * Generate insights summary card
 */
function generateInsightsSummary(data, rangeAddress) {
    if (!data || data.length === 0) return;
    
    // Basic statistics
    const numericData = data.flat().filter(val => !isNaN(val) && val !== '').map(Number);
    if (numericData.length === 0) return;
    
    const sum = numericData.reduce((a, b) => a + b, 0);
    const avg = sum / numericData.length;
    const min = Math.min(...numericData);
    const max = Math.max(...numericData);
    const count = numericData.length;
    
    const insightsElement = document.getElementById('analysis-insights');
    insightsElement.innerHTML = `
        <h4>Quick Insights - ${rangeAddress}</h4>
        <div class="insight-metric">
            <span>Data Points:</span>
            <span class="metric-value">${count}</span>
        </div>
        <div class="insight-metric">
            <span>Average:</span>
            <span class="metric-value">${avg.toFixed(2)}</span>
        </div>
        <div class="insight-metric">
            <span>Range:</span>
            <span class="metric-value">${min} - ${max}</span>
        </div>
        <div class="insight-metric">
            <span>Total Sum:</span>
            <span class="metric-value">${sum.toFixed(2)}</span>
        </div>
    `;
    insightsElement.classList.remove('hidden');
}

// === CHART CREATION ===

/**
 * Select chart type
 */
function selectChartType(type) {
    // Remove selection from all chart type buttons
    document.querySelectorAll('.chart-type-btn').forEach(btn => {
        btn.classList.remove('selected');
    });
    
    // Select the clicked type
    document.querySelector(`[data-type="${type}"]`).classList.add('selected');
    selectedChartType = type;
}

/**
 * Select range for chart
 */
async function selectChartRange() {
    try {
        const range = await getSelectedRange();
        if (range) {
            document.getElementById('chart-range').value = range.address;
        } else {
            aiService.showError('Please select a range in the worksheet first.');
        }
    } catch (error) {
        console.error('Error selecting chart range:', error);
        aiService.showError('Failed to select chart range.');
    }
}

/**
 * Create chart from selected data
 */
async function createChart() {
    const rangeAddress = document.getElementById('chart-range').value.trim();
    const description = document.getElementById('chart-description').value.trim();
    
    if (!rangeAddress) {
        aiService.showError('Please select a data range for the chart.');
        return;
    }
    
    try {
        const data = await getRangeData(rangeAddress);
        if (!data || data.length === 0) {
            aiService.showError('No data found in the specified range.');
            return;
        }
        
        // Create the chart using Excel API
        const chartCreated = await insertChart(rangeAddress, selectedChartType);
        
        if (chartCreated) {
            let response = `Chart created successfully!\n\nChart Type: ${selectedChartType.charAt(0).toUpperCase() + selectedChartType.slice(1)}\nData Range: ${rangeAddress}`;
            
            if (description) {
                // Generate AI suggestions for chart improvement
                const settings = odinUI.getSettings();
                const prompt = `Based on this chart description: "${description}" and data range ${rangeAddress}, provide suggestions for:
                1. Chart title
                2. Axis labels
                3. Color scheme
                4. Formatting improvements
                5. Additional insights the chart could show`;
                
                const suggestions = await aiService.generateText(prompt, {
                    maxTokens: settings.maxTokens,
                    temperature: 0.4
                });
                
                response += '\n\nAI Suggestions:\n' + suggestions;
            }
            
            odinUI.showResponse('chart-response', response, false);
            aiService.showSuccess('Chart created successfully!');
        } else {
            aiService.showError('Failed to create chart. Please try again.');
        }
    } catch (error) {
        console.error('Chart creation error:', error);
        aiService.showError('Failed to create chart.');
    }
}

// === FORMATTING ===

/**
 * Select range for formatting
 */
async function selectFormatRange() {
    try {
        const range = await getSelectedRange();
        if (range) {
            document.getElementById('format-range').value = range.address;
        } else {
            aiService.showError('Please select a range in the worksheet first.');
        }
    } catch (error) {
        console.error('Error selecting format range:', error);
        aiService.showError('Failed to select format range.');
    }
}

/**
 * Apply quick formatting
 */
async function applyQuickFormat(formatType) {
    const rangeAddress = document.getElementById('format-range').value.trim();
    
    if (!rangeAddress) {
        aiService.showError('Please select a range to format first.');
        return;
    }
    
    try {
        let success = false;
        let message = '';
        
        switch (formatType) {
            case 'table':
                success = await formatAsTable(rangeAddress);
                message = 'Data formatted as table';
                break;
            case 'currency':
                success = await applyCurrencyFormat(rangeAddress);
                message = 'Currency formatting applied';
                break;
            case 'percentage':
                success = await applyPercentageFormat(rangeAddress);
                message = 'Percentage formatting applied';
                break;
            case 'date':
                success = await applyDateFormat(rangeAddress);
                message = 'Date formatting applied';
                break;
            case 'conditional':
                success = await applyConditionalFormatting(rangeAddress);
                message = 'Conditional formatting applied';
                break;
            case 'headers':
                success = await styleHeaders(rangeAddress);
                message = 'Header styling applied';
                break;
            case 'borders':
                success = await addBorders(rangeAddress);
                message = 'Borders added';
                break;
            case 'colors':
                success = await applyColorCoding(rangeAddress);
                message = 'Color coding applied';
                break;
        }
        
        if (success) {
            odinUI.showResponse('format-response', message, false);
            aiService.showSuccess(message);
        } else {
            aiService.showError('Failed to apply formatting.');
        }
    } catch (error) {
        console.error('Quick format error:', error);
        aiService.showError('Failed to apply formatting.');
    }
}

/**
 * Apply custom formatting based on AI suggestions
 */
async function customFormat() {
    const rangeAddress = document.getElementById('format-range').value.trim();
    const instructions = document.getElementById('format-instructions').value.trim();
    
    if (!rangeAddress) {
        aiService.showError('Please select a range to format first.');
        return;
    }
    
    if (!instructions) {
        aiService.showError('Please provide formatting instructions.');
        return;
    }
    
    try {
        const settings = odinUI.getSettings();
        
        // Get data preview for context
        const data = await getRangeData(rangeAddress);
        const dataPreview = data.slice(0, 5).map(row => row.join('\t')).join('\n');
        
        const prompt = `Based on these formatting instructions: "${instructions}"
        
        Data range: ${rangeAddress}
        Sample data:
        ${dataPreview}
        
        Provide step-by-step Excel formatting instructions that can be applied manually or programmatically. Include:
        1. Specific formatting steps
        2. Color codes (if applicable)
        3. Number format codes
        4. Border and alignment settings
        5. Conditional formatting rules (if needed)`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('format-response', response, true);
        aiService.showSuccess('Formatting instructions generated!');
    } catch (error) {
        console.error('Custom format error:', error);
        aiService.showError('Failed to generate formatting instructions.');
    }
}

// === ADVANCED FEATURES ===

/**
 * Create PivotTable
 */
async function createPivotTable() {
    try {
        const range = await getSelectedRange();
        if (!range) {
            aiService.showError('Please select a data range first.');
            return;
        }
        
        const settings = odinUI.getSettings();
        const data = await getRangeData(range.address);
        const dataPreview = data.slice(0, 10).map(row => row.join('\t')).join('\n');
        
        const prompt = `Based on this data structure, suggest the best PivotTable configuration:

Data from ${range.address}:
${dataPreview}

Provide:
1. Recommended row fields
2. Recommended column fields  
3. Recommended value fields
4. Suggested aggregation methods
5. Potential insights this PivotTable could reveal
6. Step-by-step creation instructions`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        // Show in analysis section
        odinUI.showResponse('analysis-response', response, false);
        showAnalysisSection();
        aiService.showSuccess('PivotTable suggestions generated!');
    } catch (error) {
        console.error('PivotTable error:', error);
        aiService.showError('Failed to generate PivotTable suggestions.');
    }
}

/**
 * Data validation setup
 */
async function dataValidation() {
    try {
        const range = await getSelectedRange();
        if (!range) {
            aiService.showError('Please select a range first.');
            return;
        }
        
        const settings = odinUI.getSettings();
        const prompt = `For the selected range ${range.address}, suggest appropriate data validation rules. Consider:

1. Data type validation (numbers, dates, text)
2. Value range restrictions
3. List validation options
4. Custom validation formulas
5. Error messages for invalid entries

Provide specific Excel data validation settings and formulas.`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('format-response', response, false);
        showFormatSection();
        aiService.showSuccess('Data validation suggestions generated!');
    } catch (error) {
        console.error('Data validation error:', error);
        aiService.showError('Failed to generate data validation suggestions.');
    }
}

/**
 * What-if analysis
 */
async function whatIfAnalysis() {
    try {
        const range = await getSelectedRange();
        if (!range) {
            aiService.showError('Please select a data range first.');
            return;
        }
        
        const settings = odinUI.getSettings();
        const data = await getRangeData(range.address);
        const dataPreview = data.slice(0, 10).map(row => row.join('\t')).join('\n');
        
        const prompt = `Based on this data, suggest what-if analysis scenarios:

Data from ${range.address}:
${dataPreview}

Provide:
1. Goal Seek scenarios (what input value achieves a target output)
2. Data Table scenarios (how changing 1-2 variables affects results)
3. Scenario Manager setups (best case, worst case, most likely)
4. Specific cell references and formulas to use
5. Step-by-step setup instructions`;
        
        const response = await aiService.generateText(prompt, {
            maxTokens: settings.maxTokens,
            temperature: 0.3
        });
        
        odinUI.showResponse('analysis-response', response, false);
        showAnalysisSection();
        aiService.showSuccess('What-if analysis suggestions generated!');
    } catch (error) {
        console.error('What-if analysis error:', error);
        aiService.showError('Failed to generate what-if analysis suggestions.');
    }
}

/**
 * Show settings panel
 */
function showSettings() {
    hideAllSections();
    const panel = document.getElementById('settings-panel');
    if (panel) {
        panel.innerHTML = '';
        const settingsCard = odinUI.createSettingsPanel();
        panel.appendChild(settingsCard);
        panel.style.display = 'block';
    }
}

// === EXCEL.JS INTEGRATION ===

/**
 * Get current cell address
 */
async function getCurrentCellAddress() {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const range = context.workbook.getSelectedRange();
                range.load('address');
                await context.sync();
                resolve(range.address.split('!')[1]); // Remove worksheet name
            } catch (error) {
                reject(error);
            }
        });
    });
}

/**
 * Get selected range
 */
async function getSelectedRange() {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const range = context.workbook.getSelectedRange();
                range.load(['address', 'rowCount', 'columnCount']);
                await context.sync();
                
                if (range.rowCount > 0 && range.columnCount > 0) {
                    resolve({
                        address: range.address.split('!')[1],
                        rowCount: range.rowCount,
                        columnCount: range.columnCount
                    });
                } else {
                    resolve(null);
                }
            } catch (error) {
                reject(error);
            }
        });
    });
}

/**
 * Get worksheet context for AI
 */
async function getWorksheetContext() {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const usedRange = worksheet.getUsedRange();
                
                worksheet.load('name');
                usedRange.load(['address', 'rowCount', 'columnCount']);
                
                await context.sync();
                
                const contextInfo = `
Worksheet: ${worksheet.name}
Used Range: ${usedRange.address}
Dimensions: ${usedRange.rowCount} rows × ${usedRange.columnCount} columns`;
                
                resolve(contextInfo);
            } catch (error) {
                resolve('No data found in worksheet');
            }
        });
    });
}

/**
 * Get data from specified range
 */
async function getRangeData(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                range.load('values');
                await context.sync();
                resolve(range.values);
            } catch (error) {
                reject(error);
            }
        });
    });
}

/**
 * Get data preview for display
 */
async function getDataPreview(rangeAddress) {
    try {
        const data = await getRangeData(rangeAddress);
        if (!data || data.length === 0) return 'No data found';
        
        const preview = data.slice(0, 10).map((row, index) => {
            const rowData = row.map(cell => cell || '').join('\t');
            return `${index + 1}: ${rowData}`;
        }).join('\n');
        
        return `Data Preview (${rangeAddress}):\n${preview}${data.length > 10 ? '\n... and ' + (data.length - 10) + ' more rows' : ''}`;
    } catch (error) {
        return 'Error loading data preview';
    }
}

/**
 * Auto-detect data range
 */
async function autoDetectDataRange() {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const usedRange = worksheet.getUsedRange();
                usedRange.load('address');
                await context.sync();
                resolve(usedRange.address.split('!')[1]);
            } catch (error) {
                resolve(null);
            }
        });
    });
}

/**
 * Insert chart into worksheet
 */
async function insertChart(rangeAddress, chartType) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                
                const chart = worksheet.charts.add(chartType, range, 'Auto');
                chart.title.text = `${chartType.charAt(0).toUpperCase() + chartType.slice(1)} Chart`;
                
                await context.sync();
                resolve(true);
            } catch (error) {
                console.error('Chart creation error:', error);
                resolve(false);
            }
        });
    });
}

// === FORMATTING FUNCTIONS ===

/**
 * Format range as table
 */
async function formatAsTable(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                
                const table = worksheet.tables.add(range, true);
                table.style = 'TableStyleMedium2';
                
                await context.sync();
                resolve(true);
            } catch (error) {
                console.error('Table formatting error:', error);
                resolve(false);
            }
        });
    });
}

/**
 * Apply currency formatting
 */
async function applyCurrencyFormat(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                range.numberFormat = [['$#,##0.00']];
                await context.sync();
                resolve(true);
            } catch (error) {
                resolve(false);
            }
        });
    });
}

/**
 * Apply percentage formatting
 */
async function applyPercentageFormat(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                range.numberFormat = [['0.00%']];
                await context.sync();
                resolve(true);
            } catch (error) {
                resolve(false);
            }
        });
    });
}

/**
 * Apply date formatting
 */
async function applyDateFormat(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                range.numberFormat = [['mm/dd/yyyy']];
                await context.sync();
                resolve(true);
            } catch (error) {
                resolve(false);
            }
        });
    });
}

/**
 * Apply conditional formatting
 */
async function applyConditionalFormatting(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                
                const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                conditionalFormat.colorScale.criteria = {
                    minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "red" },
                    midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percentile, color: "yellow" },
                    maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "green" }
                };
                
                await context.sync();
                resolve(true);
            } catch (error) {
                resolve(false);
            }
        });
    });
}

/**
 * Style headers
 */
async function styleHeaders(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                const headerRow = range.getRow(0);
                
                headerRow.format.font.bold = true;
                headerRow.format.fill.color = '#4472C4';
                headerRow.format.font.color = 'white';
                
                await context.sync();
                resolve(true);
            } catch (error) {
                resolve(false);
            }
        });
    });
}

/**
 * Add borders
 */
async function addBorders(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                
                range.format.borders.getItem('EdgeTop').style = 'Continuous';
                range.format.borders.getItem('EdgeBottom').style = 'Continuous';
                range.format.borders.getItem('EdgeLeft').style = 'Continuous';
                range.format.borders.getItem('EdgeRight').style = 'Continuous';
                range.format.borders.getItem('InsideVertical').style = 'Continuous';
                range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
                
                await context.sync();
                resolve(true);
            } catch (error) {
                resolve(false);
            }
        });
    });
}

/**
 * Apply color coding
 */
async function applyColorCoding(rangeAddress) {
    return new Promise((resolve, reject) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                const range = worksheet.getRange(rangeAddress);
                
                // Apply alternating row colors
                for (let i = 0; i < range.rowCount; i++) {
                    const row = range.getRow(i);
                    if (i % 2 === 0) {
                        row.format.fill.color = '#F2F2F2';
                    }
                }
                
                await context.sync();
                resolve(true);
            } catch (error) {
                resolve(false);
            }
        });
    });
}

// === UTILITY FUNCTIONS ===

/**
 * Override the default insert text function to use Excel-specific implementation
 */
odinUI.insertText = async function(responseDisplayId, event) {
    const display = document.getElementById(responseDisplayId);
    if (!display) return;
    
    const text = display.textContent;
    
    // Check if it's a formula (starts with =)
    if (text.trim().startsWith('=')) {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.formulas = [[text.trim()]];
                await context.sync();
            });
            aiService.showSuccess('Formula inserted successfully!');
        } catch (error) {
            console.error('Insert formula error:', error);
            aiService.showError('Failed to insert formula. Please try again.');
        }
    } else {
        // Insert as regular text
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.values = [[text]];
                await context.sync();
            });
            aiService.showSuccess('Text inserted successfully!');
        } catch (error) {
            console.error('Insert text error:', error);
            aiService.showError('Failed to insert text. Please try again.');
        }
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
                        if (target.id === 'formula-section') {
                            document.getElementById('formula-prompt')?.focus();
                        } else if (target.id === 'analysis-section') {
                            document.getElementById('analysis-question')?.focus();
                        } else if (target.id === 'chart-section') {
                            document.getElementById('chart-range')?.focus();
                        } else if (target.id === 'format-section') {
                            document.getElementById('format-instructions')?.focus();
                        }
                    }, 100);
                }
            }
        });
    });
    
    const sections = ['formula-section', 'analysis-section', 'chart-section', 'format-section'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            observer.observe(section, { attributes: true });
        }
    });
});
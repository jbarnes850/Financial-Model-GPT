// Global Constants
const OPENAI_API_ENDPOINT = "https://api.openai.com/v1/chat/completions";

/**
 * Retrieves the OpenAI API key from script properties.
 * Ensure you've set the API key correctly in the script's project properties.
 */
function getOpenAiApiKey() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var apiKey = scriptProperties.getProperty('OPENAI_API_KEY');
    if (!apiKey) {
        Logger.log('The API key has not been set in the script properties.');
        throw new Error('The API key has not been set in the script properties.');
    }
    return apiKey;
}

/**
 * Function that orchestrates the whole analysis process when menu item is clicked.
 * This function is triggered from the UI and orchestrates the analysis process.
 */
function startFinancialAnalysis() {
    var financialModelData = getFinancialModelData();
    if (financialModelData && isValidFinancialData(financialModelData)) {
        showSidebar();  // Show the sidebar first to provide a visual cue
        try {
            var prompt = constructPrompt(financialModelData);
            getGptApiResponse(prompt, handleApiResponse, handleApiError);
        } catch (error) {
            // Handle construction errors and inform the user
            SpreadsheetApp.getUi().alert("Failed to construct the analysis prompt: " + error.message);
            Logger.log("Error during prompt construction: " + error.toString());
        }
    } else {
        SpreadsheetApp.getUi().alert("Failed to retrieve financial data. Please check the 'Forecast Dashboard' sheet.");
    }
}

function getFinancialModelData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Forecast Dashboard");

    if (!sheet) {
        Logger.log("Error: Forecast Dashboard sheet not found.");
        return null;  // Indicating that the sheet wasn't found
    }

    try {
        // Fetching all required data in a single call
        var dataRange = sheet.getRange("C4:R17").getValues();

        // Map data to respective fields
        var financialData = {
            revenue: dataRange[0], // Corresponds to C4:R4
            revenueGrowthRate: dataRange[1], // C5:R5
            grossMargin: dataRange[2], // C6:R6
            CoGS: dataRange[3], // C7:R7
            totalQuarterlyBurn: dataRange[4], // C8:R8
            cashOnHand: dataRange[5], // C9:R9
            cashFromNewFinancing: dataRange[6], // C10:R10
            burnMultiple: dataRange[7], // C11:R11
            remainingMonthsOfRunway: dataRange[8], // C12:R12
            averageContractValue: dataRange[10], // C14:R14 (Skipping one row)
            netRevenueRetention: dataRange[11], // C15:R15
            userRetention: dataRange[12], // C16:R16
            cacPaybackInMonths: dataRange[13] // C17:R17
        };

        Logger.log("Financial Data Extracted: " + JSON.stringify(financialData));
        return financialData;

    } catch (e) {
        Logger.log("Unexpected error: " + e.toString()); // Detailed logging for internal use
        // User-friendly message
        SpreadsheetApp.getUi().alert("An unexpected error occurred. Please try again or contact support.");
        return null;  // Return null indicating that an error occurred
    }
}

/**
 * Validates the financial data structure to ensure all required fields are present and correct.
 * @param {Object} financialModelData - The financial data to validate.
 * @return {Boolean} - True if valid, false otherwise.
 */
function isValidFinancialData(financialModelData) {
    const requiredFields = [
        'revenue', 'revenueGrowthRate', 'grossMargin', 'CoGS',
        'totalQuarterlyBurn', 'cashOnHand', 'cashFromNewFinancing',
        'burnMultiple', 'remainingMonthsOfRunway', 'averageContractValue',
        'netRevenueRetention', 'userRetention', 'cacPaybackInMonths'
    ];

    if (!financialModelData || typeof financialModelData !== 'object') {
        Logger.log('Financial data is not valid or not found:', financialModelData);
        return false;
    }

    return requiredFields.every(function(field) {
        return financialModelData.hasOwnProperty(field) &&
               Array.isArray(financialModelData[field]) &&
               financialModelData[field].length > 0;
    });
}

/**
 * Opens a sidebar in the document containing the HTML file for user interaction.
 */
function showSidebar() {
    try {
        var htmlOutput = HtmlService.createHtmlOutputFromFile('Page')
            .setTitle('Financial Model Analysis')
            .setWidth(300);
        SpreadsheetApp.getUi().showSidebar(htmlOutput);
    } catch (e) {
        Logger.log("Failed to open the sidebar: " + e.toString());
    }
}

/**
 * Constructs the prompt for the OpenAI API from the given financial data.
 * @param {Object} financialModelData - The financial data object.
 * @return {String} - The constructed prompt.
 */
function constructPrompt(financialModelData) {
    // Before constructing the prompt, ensure the financial data is valid
    if (!isValidFinancialData(financialModelData)) {
        throw new Error("Financial data validation failed. Ensure all required fields are present and are arrays of strings.");
    }

    // Construct the prompt using template literals for readability
    var prompt = `Financial Data (monthly):\n` +
    `Revenue: ${financialModelData.revenue.join(", ")}\n` +
    `Revenue Growth Rate: ${financialModelData.revenueGrowthRate.join(", ")}\n` +
    `Gross Margin: ${financialModelData.grossMargin.join(", ")}\n` +
    `CoGS: ${financialModelData.CoGS.join(", ")}\n` +
    `Total Quarterly Burn: ${financialModelData.totalQuarterlyBurn.join(", ")}\n` +
    `Cash on Hand: ${financialModelData.cashOnHand.join(", ")}\n` +
    `Cash from New Financing: ${financialModelData.cashFromNewFinancing.join(", ")}\n` +
    `Burn Multiple: ${financialModelData.burnMultiple.join(", ")}\n` +
    `Remaining Months of Runway: ${financialModelData.remainingMonthsOfRunway.join(", ")}\n` +
    `Average Contract Value: ${financialModelData.averageContractValue.join(", ")}\n` +
    `You are an expert coach to startup founders. Assess the financial health and growth potential of an early-stage, pre-seed to seed SaaS startup based on a 3-year financial forecast model. Benchmark the startup's financial performance against typical industry standards for early-stage startups. The financial model includes a range of metrics such as monthly revenue, growth rate, gross margin, and more.\n\n` +
    `Net Revenue Retention: ${financialModelData.netRevenueRetention.join(", ")}\n` +
    `User Retention: ${financialModelData.userRetention.join(", ")}\n` +
    `CAC Payback in Months: ${financialModelData.cacPaybackInMonths.join(", ")}\n\n` +
    `Given this comprehensive data, please:\n` +
    `1. Analyze the startup's financial health, considering the revenue, margins, burn rate, and retention metrics.\n` +
    `2. Evaluate the startup's growth trajectory and investment efficiency, taking into account the revenue growth rate, burn multiple, and runway.\n` +
    `3. Provide concise strategic recommendations for optimizing the startup's financial model and improving key metrics to better align with successful industry standards.\n\n` +
    `Be kind but direct and encouraging. Deliver your analysis with actionable insights and detailed suggestions based on the provided data and industry benchmarks.`;
    // Log the constructed prompt for debugging
    Logger.log("Constructed Prompt: " + prompt);
    return prompt; // Return the constructed prompt
}

/**
 * Makes an API call to OpenAI's chat completions endpoint with the constructed prompt.
 * @param {String} prompt - The constructed prompt.
 * @param {Function} successCallback - The callback to escalate a success response.
 * @param {Function} errorCallback - The callback to escalate an error response.
 */
function getGptApiResponse(prompt, successCallback, errorCallback) {
    Logger.log('Received prompt for API call:', prompt);
    if (typeof prompt !== 'string' || !prompt.trim()) {
        const errorMessage = 'Prompt is not a valid string.';
        Logger.log(errorMessage);
        errorCallback(errorMessage);
        return;
    }

    var data = {
        "model": "gpt-4-1106-preview",
        "messages": [{"role": "user", "content": prompt}]
    };

    var options = {
        "method": "post",
        "contentType": "application/json",
        "headers": {"Authorization": "Bearer " + getOpenAiApiKey()},
        "payload": JSON.stringify(data),
        "muteHttpExceptions": true
    };

    try {
        var response = UrlFetchApp.fetch(OPENAI_API_ENDPOINT, options);
        var content = JSON.parse(response.getContentText());

        // Check for errors and handle the response appropriately
        if (content.error) {
            // Handle error
            errorCallback("API error: " + content.error.message);
        } else if (content.choices && content.choices.length > 0) {
            // Send the analysis result back to the client side
            successCallback(content.choices[0].message.content);
        } else {
            // Handle unexpected response format
            errorCallback("Received an unexpected format from API.");
        }
    } catch (error) {
        // Handle general errors
        errorCallback("Error calling OpenAI API: " + error.toString());
    }
}

// Callback function for successful API response
function handleApiResponse(response) {
    Logger.log("Analysis Result: " + response);
    // Call server-side function to update sidebar
    updateSidebarContent(response);
}

function updateSidebarContent(response) {
    // Use google.script.run to send data to client-side
    google.script.run.withSuccessHandler(updateSidebarClient).sendDataToSidebar(response);
}

// Callback function for handling API errors
function handleApiError(error) {
    Logger.log(error); // Detailed logging for internal use
    // User-friendly message
    SpreadsheetApp.getUi().alert("We encountered an issue while processing your request. Please try again later or contact support if the problem persists.");
}

/**
 * Function called when the spreadsheet is opened. Adds custom menu items.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Financial Analysis')
        .addItem('Start Analysis', 'startFinancialAnalysis')
        .addToUi();
}

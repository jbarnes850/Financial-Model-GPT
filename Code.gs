// Global Constants
const OPENAI_API_ENDPOINT = "https://api.openai.com/v1/chat/completions"; // Using the chat completions endpoint
const OPENAI_API_KEY = "INSERT OPENAI KEY";

/**
 * Main function to initiate analysis process.
 */
function myFunction() {
    try {
        var financialModelData = getFinancialModelData();
        Logger.log("Financial Model Data in myFunction: " + JSON.stringify(financialModelData));

        // Check if financial data is valid and present
        if (!financialModelData || Object.keys(financialModelData).length === 0) {
            throw new Error("No financial data found or it's invalid.");
        }

        // Call the function to use financial data with the OpenAI API
        var analysisResult = callOpenAiApi(financialModelData);
        Logger.log("Analysis Result: " + analysisResult);

        showSidebar(); // Display the sidebar if everything is successful

    } catch (e) {
        // Log the error
        Logger.log("Error in myFunction: " + e.toString());

        // Optionally, inform the user with an alert
        SpreadsheetApp.getUi().alert("Error occurred in analysis: " + e.toString());
    }
}
/**
 * Extracts comprehensive financial data from the "Forecast Dashboard" sheet.
 * Returns the data for all provided ranges.
 */

function getFinancialModelData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Forecast Dashboard");
    
    // Check if the sheet exists
    if (!sheet) {
        Logger.log("Error: Forecast Dashboard sheet not found.");
        return null;  // Returning null to indicate the sheet wasn't found
    }
    
    try {
        // Attempt to extract data from defined ranges and log it
        var revenue = sheet.getRange("C4:R4").getValues()[0];
        var revenueGrowthRate = sheet.getRange("C5:R5").getValues()[0];
        var grossMargin = sheet.getRange("C6:R6").getValues()[0];
        var CoGS = sheet.getRange("C7:R7").getValues()[0];
        var totalQuarterlyBurn = sheet.getRange("C8:R8").getValues()[0];
        var cashOnHand = sheet.getRange("C9:R9").getValues()[0];
        var cashFromNewFinancing = sheet.getRange("C10:R10").getValues()[0];
        var burnMultiple = sheet.getRange("C11:R11").getValues()[0];
        var remainingMonthsOfRunway = sheet.getRange("C12:R12").getValues()[0];
        var averageContractValue = sheet.getRange("C14:R14").getValues()[0];
        var netRevenueRetention = sheet.getRange("C15:R15").getValues()[0];
        var userRetention = sheet.getRange("C16:R16").getValues()[0];
        var cacPaybackInMonths = sheet.getRange("C17:R17").getValues()[0];

        // Constructing the data object
        var financialData = {
            revenue: revenue,
            revenueGrowthRate: revenueGrowthRate,
            grossMargin: grossMargin,
            CoGS: CoGS,
            totalQuarterlyBurn: totalQuarterlyBurn,
            cashOnHand: cashOnHand,
            cashFromNewFinancing: cashFromNewFinancing,
            burnMultiple: burnMultiple,
            remainingMonthsOfRunway: remainingMonthsOfRunway,
            averageContractValue: averageContractValue,
            netRevenueRetention: netRevenueRetention,
            userRetention: userRetention,
            cacPaybackInMonths: cacPaybackInMonths
        };

        // Logging the final data structure before returning it
        Logger.log("Financial Data Extracted: " + JSON.stringify(financialData));
        return financialData;
    } catch (e) {
        Logger.log("Failed to retrieve or parse financial data: " + e.toString());
        return null;
    }
}

/**
 * Calls the OpenAI API with extracted financial data.
 * Constructs a prompt for analysis and handles the API response.
 */
function callOpenAiApi(financialModelData) {
    // Ensure that financialModelData is not null or undefined and is an object
    if (typeof financialModelData !== 'object' || financialModelData === null) {
        Logger.log('Financial data is not valid or not found:', financialModelData);
        return 'Error: Financial data is not valid or not found.';
    }

    // Check for the existence and validity of all required fields in financialModelData
    const requiredFields = [
        'revenue', 'revenueGrowthRate', 'grossMargin', 'CoGS',
        'totalQuarterlyBurn', 'cashOnHand', 'cashFromNewFinancing',
        'burnMultiple', 'remainingMonthsOfRunway', 'averageContractValue',
        'netRevenueRetention', 'userRetention', 'cacPaybackInMonths'
    ];

    for (let field of requiredFields) {
        if (!financialModelData.hasOwnProperty(field) || !Array.isArray(financialModelData[field]) || financialModelData[field].length === 0) {
            Logger.log(`Missing or invalid field in financial data: ${field}`);
            return `Error: Missing or invalid field in financial data: ${field}`;
        }
    }

    try {
        // Construct prompt from financialModelData
    var prompt = "You are an expert coach to startup founders. Assess the financial health and growth potential of an early-stage, pre-seed to seed SaaS startup based on a 3-year financial forecast model. Benchmark the startup's financial performance against typical industry standards for early-stage startups. The financial model includes a range of metrics such as monthly revenue, growth rate, gross margin, and more.\n\n" +
                 "Financial Data (monthly):\n" +
                 "Revenue: " + financialModelData.revenue.join(", ") + "\n" +
                 "Revenue Growth Rate: " + financialModelData.revenueGrowthRate.join(", ") + "\n" +
                 "Gross Margin: " + financialModelData.grossMargin.join(", ") + "\n" +
                 "CoGS: " + financialModelData.CoGS.join(", ") + "\n" +
                 "Total Quarterly Burn: " + financialModelData.totalQuarterlyBurn.join(", ") + "\n" +
                 "Cash on Hand: " + financialModelData.cashOnHand.join(", ") + "\n" +
                 "Cash from New Financing: " + financialModelData.cashFromNewFinancing.join(", ") + "\n" +
                 "Burn Multiple: " + financialModelData.burnMultiple.join(", ") + "\n" +
                 "Remaining Months of Runway: " + financialModelData.remainingMonthsOfRunway.join(", ") + "\n" +
                 "Average Contract Value: " + financialModelData.averageContractValue.join(", ") + "\n" +
                 "Net Revenue Retention: " + financialModelData.netRevenueRetention.join(", ") + "\n" +
                 "User Retention: " + financialModelData.userRetention.join(", ") + "\n" +
                 "CAC Payback in Months: " + financialModelData.cacPaybackInMonths.join(", ") + "\n\n" +
                 "Given this comprehensive data, please:\n" +
                 "1. Analyze the startup's financial health, considering the revenue, margins, burn rate, and retention metrics.\n" +
                 "2. Evaluate the startup's growth trajectory and investment efficiency, taking into account the revenue growth rate, burn multiple, and runway.\n" +
                 "3. Provide concise strategic recommendations for optimizing the startup's financial model and improving key metrics to better align with successful industry standards.\n\n" +
                 "Be kind but direct and encouraging. Deliver your analysis with actionable insights and detailed suggestions based on the provided data and industry benchmarks.";

// Log the constructed prompt for debugging
        Logger.log("Constructed Prompt: " + prompt);

// Check if prompt is valid
        if (typeof prompt !== 'string' || prompt.trim() === '') {
            Logger.log('Prompt is not valid:', prompt);
            return 'Error: Prompt is empty or not a string.';
        }

        // Proceed with making the API call
        var analysisResult = getGptApiResponse(prompt);
        Logger.log("Analysis Result: " + analysisResult);
        return analysisResult;

    } catch (e) {
        Logger.log("Error constructing prompt or calling API: " + e.toString());
        return "Error in calling API with constructed prompt.";
    }
}

/**
 * Adds a custom menu to the Google Sheets UI.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi(); // Get the user interface
    ui.createMenu('Financial Analysis') // Create a new menu
        .addItem('Start Analysis', 'myFunction') // Add an item to the menu
        .addToUi(); // Render the menu in the UI
}

/**
 * Opens a sidebar in the document containing the HTML file for user interaction.
 */
function showSidebar() {
    try {
        var htmlOutput = HtmlService.createHtmlOutputFromFile('Page')
            .setTitle('Financial Model Analysis')
            .setWidth(300);

        // Log the HTML content for debugging
        Logger.log(htmlOutput.getContent());
        
        SpreadsheetApp.getUi().showSidebar(htmlOutput);
    } catch (e) {
        Logger.log("Failed to open the sidebar: " + e.toString());
        // Inform the user if in a production environment
        SpreadsheetApp.getUi().alert("Failed to open the sidebar. Please check the logs for more details.");
    }
}

function getGptApiResponse(prompt) {
  Logger.log("Received prompt: " + prompt); // Log the received prompt

  // Validate the prompt
  if (typeof prompt !== 'string' || !prompt.trim()) {
    var errorMessage = 'Prompt is not valid or empty: ' + JSON.stringify(prompt);
    Logger.log(errorMessage);
    return errorMessage;
  }

  // API details
  var apiUrl = "https://api.openai.com/v1/chat/completions";
  var apiKey = "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";

  // Payload for GPT API
  var data = {
    "model": "gpt-4-1106-preview",
    "messages": [
      {"role": "system", "content": "Analyze financial data"},
      {"role": "user", "content": prompt}
    ]
  };

  // Options including headers with API Key
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(data),
    "muteHttpExceptions": true // To get detailed errors
  };

  try {
    var response = UrlFetchApp.fetch(apiUrl, options); // Making the API request
    Logger.log("API Response: " + response.getContentText()); // Log full API response
    var content = JSON.parse(response.getContentText());

    // Return the part of the response we're interested in, typically the content
    if (content.choices && content.choices.length > 0) {
      return content.choices[0].message.content;
    } else {
      var unexpectedFormatError = "Received an unexpected format from API: " + JSON.stringify(content);
      Logger.log(unexpectedFormatError);
      return unexpectedFormatError;
    }
  } catch (error) {
    var errorLog = "Error calling OpenAI API: " + error.toString();
    Logger.log(errorLog);
    return errorLog; // Returning error for debugging
  }
}

/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  DocumentApp.getUi()
      .createMenu('Docopilot')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Opens a sidebar in the document.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Docopilot');
  DocumentApp.getUi().showSidebar(html);
}

const HIGHLIGHT_COLOR = '#FFF8C4'; // Light yellow/orange
const FOCUSED_HIGHLIGHT_COLOR = '#FFD54F'; // A slightly more orange/yellow, similar to Docs focus

/**
 * Gets the text content of the body of the current document.
 *
 * @return {string} The text content of the document body.
 */
function getDocumentContent() {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    return body.getText();
  } catch (e) {
    console.error("Error getting document content: " + e.stack);
    throw new Error("Could not retrieve document content: " + e.message);
  }
}

/**
 * Calls the Gemini API to get comments on the provided text.
 *
 * @param {string} finalPrompt The complete prompt including document text and instructions.
 * @param {string} apiKey The user's Gemini API key.
 * @return {Array<Object>} An array of comment objects { quote: string, comment: string }.
 */
function getGeminiComments(finalPrompt, apiKey) {
  if (!apiKey) {
    throw new Error("API Key is required to contact Gemini.");
  }

  const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;

  const requestBody = {
    contents: [{
      parts: [{
        text: finalPrompt
      }]
    }],
    generationConfig: {
      responseMimeType: "application/json",
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(API_ENDPOINT, options);
  } catch (e) {
    console.error("Network error calling Gemini API: " + e.stack);
    throw new Error("Network error communicating with Gemini: " + e.message);
  }

  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    try {
      const jsonResponse = JSON.parse(responseBody);
      const candidates = jsonResponse.candidates;
      if (candidates && candidates.length > 0 && candidates[0].content && candidates[0].content.parts && candidates[0].content.parts.length > 0) {
        const geminiJsonString = candidates[0].content.parts[0].text;
        // Parse the JSON string returned by Gemini
        const resultObject = JSON.parse(geminiJsonString);

        // Validate the new structure
        if (typeof resultObject !== 'object' || resultObject === null || !Array.isArray(resultObject.comments)) {
             console.error("Gemini response was not in the expected format {thinking: string, comments: array}:", geminiJsonString);
             throw new Error("Gemini returned data in an unexpected format. Expected {thinking: ..., comments: [...]}.");
        }

        // Extract the comments array
        const comments = resultObject.comments;
        console.log("Extracted comments from Gemini response:", JSON.stringify(comments, null, 2));
        
        // Optional: Log the thinking part
        if (resultObject.thinking) {
             console.log("Gemini Thinking:", resultObject.thinking);
        }

        // Validate the comments array structure (as before)
        if (comments.every(c => typeof c === 'object' && 'quote' in c && 'comment' in c)) {
            return comments; // Return only the comments array
        } else {
            console.error("Gemini 'comments' array contents were not in the expected format [{quote, comment}, ...]:", JSON.stringify(comments, null, 2));
            throw new Error("Gemini returned comments array with unexpected item structure.");
        }
      } else {
        // Handle cases where the response structure is unexpected or indicates no content/error (e.g., safety filters)
        console.error("Unexpected Gemini response structure or no content:", JSON.stringify(jsonResponse, null, 2));
        const finishReason = candidates?.[0]?.finishReason;
        const safetyRatings = candidates?.[0]?.safetyRatings;
        let errorMessage = "Gemini returned an unexpected response structure.";
        if (finishReason && finishReason !== "STOP") {
            errorMessage = `Gemini generation finished unexpectedly: ${finishReason}.`;
            if(safetyRatings) {
                errorMessage += ` Safety ratings: ${JSON.stringify(safetyRatings)}`;
            }
        }
        throw new Error(errorMessage);
      }
    } catch (e) {
      console.error("Error parsing Gemini response: " + e.stack + "\nResponse Body:\n" + responseBody);
      // Include original error message and a LARGER snippet of the body in the thrown error
      const snippet = responseBody ? responseBody.substring(0, 500) + '...' : 'N/A'; // Increased snippet size
      throw new Error(`Could not parse Gemini response. Parse Error: ${e.message || e.toString()}. Body Snippet: ${snippet}`);
    }
  } else {
    console.error(`Gemini API Error: ${responseCode}\nResponse Body: ${responseBody}`);
    let errorDetail = responseBody;
    try {
      // Try parsing error details if they are JSON
      const errorJson = JSON.parse(responseBody);
      if (errorJson.error && errorJson.error.message) {
          errorDetail = errorJson.error.message;
      }
    } catch (parseError) { /* Ignore if not JSON */ }
    throw new Error(`Gemini API request failed with status ${responseCode}: ${errorDetail}`);
  }
}

/**
 * Highlights the specified text segments (quotes) in the document.
 * Optionally highlights one specific quote with a different focus color.
 *
 * @param {Array<string>} quotesToHighlight An array of quote strings to find and highlight.
 * @param {string} [quoteInFocus] Optional. The specific quote string to highlight with the focus color.
 */
function highlightCommentsInDoc(quotesToHighlight, quoteInFocus = null) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    // 1. Clear existing highlights
    console.log("Clearing existing highlights...");
    try {
      let searchResult = body.findText('.*');
      while (searchResult) {
        const textElement = searchResult.getElement().asText();
        if (textElement) {
          const startOffset = searchResult.getStartOffset();
          const endOffset = searchResult.getEndOffsetInclusive(); 
          textElement.setBackgroundColor(startOffset, endOffset, null);
        }
        searchResult = body.findText('.*', searchResult);
      }
      console.log("Finished clearing highlights.");
    } catch (e) {
        console.error("Error during highlight clearing: " + e.stack);
    }

    if (!quotesToHighlight || quotesToHighlight.length === 0) {
      console.log("No quotes provided for highlighting.");
      return;
    }

    // 2. Highlight all quotes with the standard color
    console.log(`Attempting to highlight ${quotesToHighlight.length} quotes with standard color.`);
    quotesToHighlight.forEach((quote, index) => {
      if (!quote || typeof quote !== 'string' || quote.trim() === '') {
          console.warn(`Skipping quote index ${index} due to empty or invalid value.`);
          return;
      }
      if (quote === quoteInFocus) return;
      applyHighlight(body, quote, HIGHLIGHT_COLOR);
    });

    // 3. Highlight the specific focused quote with the focus color
    if (quoteInFocus) {
        console.log(`Attempting to highlight focused quote: "${quoteInFocus}"`);
        applyHighlight(body, quoteInFocus, FOCUSED_HIGHLIGHT_COLOR);
    }

    console.log("Highlighting process completed.");
  } catch (e) {
    console.error("Error highlighting quotes in document: " + e.stack);
    throw new Error("Failed to highlight quotes in document: " + e.message);
  }
}

/**
 * Helper function to apply background color to all occurrences of a text quote.
 * @param {Body} body The document body element.
 * @param {string} quote The text to find.
 * @param {string} color The background color to apply.
 */
function applyHighlight(body, quote, color) {
    if (!quote || !color) return;
    let searchResult = body.findText(quote);
    let count = 0;
    while (searchResult !== null) {
        const element = searchResult.getElement();
        const start = searchResult.getStartOffset();
        const end = searchResult.getEndOffsetInclusive();

        if (element.asText()) {
            element.asText().setBackgroundColor(start, end, color);
            count++;
        }
        searchResult = body.findText(quote, searchResult);
    }
    if (count === 0) {
        console.warn(`Text not found for applying highlight (${color}): "${quote}"`);
    }
}

/**
 * Scrolls the document view to the first occurrence of the specified quote.
 * @param {string} quote The text quote to find and scroll to.
 */
function scrollToQuote(quote) {
  if (!quote || typeof quote !== 'string' || quote.trim() === '') {
    console.log("scrollToQuote: Invalid quote provided.");
    return; 
  }
  console.log(`scrollToQuote: Attempting to find and scroll to "${quote}"`);
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const searchResult = body.findText(quote); // Find first occurrence

    if (searchResult) {
        const element = searchResult.getElement();
        const start = searchResult.getStartOffset();
        const end = searchResult.getEndOffsetInclusive();
        
        // Build the range for the first occurrence
        const rangeBuilder = doc.newRange();
        rangeBuilder.addElement(element, start, end);
        const rangeToSelect = rangeBuilder.build();
        
        console.log("scrollToQuote: Found quote, setting selection.");
        doc.setSelection(rangeToSelect);
    } else {
        console.warn(`scrollToQuote: Quote not found in document: "${quote}"`);
    }
  } catch (e) {
      console.error(`scrollToQuote: Error finding/setting selection for "${quote}": ${e.stack}`);
      // Don't throw, just log the error
  }
}

/**
 * Gets the current state of the user's cursor, primarily its location,
 * the background color, and the text of the surrounding element.
 *
 * @return {object|null} An object like {bgColor: string|null, offset: number, elementText: string|null } or null if no cursor.
 */
function getCursorState() {
  console.log("getCursorState: Running...");
  try {
    const doc = DocumentApp.getActiveDocument();
    const cursor = doc.getCursor();

    if (!cursor) {
      console.log("getCursorState: No cursor found.");
      return null;
    }

    const offset = cursor.getOffset();
    const element = cursor.getElement();
    console.log(`getCursorState: Cursor offset=${offset}, Element type=${element ? element.getType() : 'null'}`);

    let bgColor = null;
    let elementText = null; // Variable to hold the text

    if (element && element.getType() === DocumentApp.ElementType.TEXT) {
        const textElement = element.asText();
        elementText = textElement.getText(); // Get the text content
        console.log(`getCursorState: Cursor in TEXT element. Length=${elementText.length}`);

        // Check background color *before* the cursor first
        if (offset > 0) {
            bgColor = textElement.getBackgroundColor(offset - 1);
            console.log(`getCursorState: Background at offset ${offset - 1} (before cursor): ${bgColor}`);
        }

        // If not highlighted before, check *at* the cursor
        if (bgColor !== HIGHLIGHT_COLOR && offset < elementText.length) { // Use elementText.length
            const bgColorAt = textElement.getBackgroundColor(offset);
            console.log(`getCursorState: Background at offset ${offset} (at cursor): ${bgColorAt}`);
            if (bgColorAt === HIGHLIGHT_COLOR) {
                bgColor = HIGHLIGHT_COLOR;
            }
        }
    } else if (element) {
         console.log(`getCursorState: Cursor in non-TEXT element type: ${element.getType()}`);
         // Try getting surrounding text even for non-text elements (might be near text)
         const surroundingTextElement = cursor.getSurroundingText();
         if (surroundingTextElement) {
             elementText = surroundingTextElement.getText();
             if(elementText.length > 0) {
                 console.log(`getCursorState: Found surrounding text. Length=${elementText.length}`);
                 // Check background at the start of the surrounding text if cursor is at offset 0
                 if (offset === 0) { 
                    const surroundingBg = surroundingTextElement.getBackgroundColor(0);
                    console.log(`getCursorState: Surrounding text(0) background: ${surroundingBg}`);
                    if (surroundingBg === HIGHLIGHT_COLOR) {
                        bgColor = HIGHLIGHT_COLOR;
                    }
                 } 
             } else {
                 elementText = null; // No actual text found
             }
         } 
    }

    console.log(`getCursorState: Returning state: bgColor=${bgColor}, offset=${offset}, text='${elementText ? elementText.substring(0, 50) + '...' : 'null'}'`);
    return {
        bgColor: bgColor,
        offset: offset,
        elementText: elementText // Include the element text
    };

  } catch (e) {
    console.error("Error getting cursor state: " + e.stack);
    return null;
  }
}

/**
 * Saves the user's Gemini API key to User Properties.
 *
 * @param {string} apiKey The API key to save.
 */
function saveApiKey(apiKey) {
  try {
    if (apiKey && typeof apiKey === 'string') {
      PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', apiKey);
      console.log("API Key saved successfully.");
    } else {
      // Optionally clear the key if an empty/invalid value is passed
      PropertiesService.getUserProperties().deleteProperty('GEMINI_API_KEY');
      console.log("Cleared saved API Key.");
    }
  } catch (e) {
    console.error("Error saving API Key: " + e.stack);
    // Decide if we should throw or just log, maybe log is better here
    // throw new Error("Could not save API key.");
  }
}

/**
 * Retrieves the user's saved Gemini API key from User Properties.
 *
 * @return {string|null} The saved API key, or null if not found.
 */
function getApiKey() {
  try {
    const savedKey = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
    console.log("Retrieved API Key: " + (savedKey ? 'Found' : 'Not Found'));
    return savedKey;
  } catch (e) {
    console.error("Error retrieving API Key: " + e.stack);
    return null; // Return null on error
  }
}

// -- Prompt Management --

const USER_PROMPT_KEY = 'USER_ANALYSIS_PROMPT';
let defaultPromptCache = null; // Cache for the default prompt

/**
 * Reads the default prompt text from the 'default_prompt.txt' file.
 * Caches the result for efficiency.
 * @return {string} The default prompt text.
 */
function getDefaultPrompt() {
  if (defaultPromptCache === null) {
    try {
      // Use HtmlService to read a text file (common workaround in Apps Script)
      const htmlOutput = HtmlService.createHtmlOutputFromFile('default_prompt.txt');
      defaultPromptCache = htmlOutput.getContent();
      console.log("Successfully read and cached default prompt.");
    } catch (e) {
      console.error("Error reading default_prompt.txt: " + e.stack);
      // Fallback prompt if file reading fails
      defaultPromptCache = "Please review the following document text and provide constructive comments. For each comment, identify the exact phrase or sentence from the text that the comment refers to. Present your output STRICTLY as a JSON array of objects, where each object has a \"quote\" key (containing the exact text phrase) and a \"comment\" key (containing your feedback).\n\nDocument Text:\n---\n{docText}\n---";
    }
  }
  return defaultPromptCache;
}

/**
 * Saves the user's custom analysis prompt to User Properties.
 * @param {string} promptText The prompt text to save.
 */
function saveUserPrompt(promptText) {
  try {
    if (promptText && typeof promptText === 'string') {
      // Avoid saving if it's exactly the default prompt to save space/quota
      if (promptText === getDefaultPrompt()) {
          PropertiesService.getUserProperties().deleteProperty(USER_PROMPT_KEY);
          console.log("User prompt matches default, cleared saved property.");
      } else {
          PropertiesService.getUserProperties().setProperty(USER_PROMPT_KEY, promptText);
          console.log("User prompt saved successfully.");
      }
    } else {
      // Clear if invalid prompt is passed
      PropertiesService.getUserProperties().deleteProperty(USER_PROMPT_KEY);
      console.log("Cleared saved user prompt.");
    }
  } catch (e) {
    console.error("Error saving user prompt: " + e.stack);
  }
}

/**
 * Retrieves the user's saved custom analysis prompt from User Properties.
 * If no custom prompt is saved, returns the default prompt.
 * @return {string} The saved or default prompt text.
 */
function getUserOrDefaultPrompt() {
  let userPrompt = null;
  try {
    userPrompt = PropertiesService.getUserProperties().getProperty(USER_PROMPT_KEY);
  } catch (e) {
    console.error("Error retrieving user prompt: " + e.stack);
  }
  
  if (userPrompt) {
    console.log("Retrieved saved user prompt.");
    return userPrompt;
  } else {
    console.log("No saved user prompt found, returning default.");
    return getDefaultPrompt();
  }
} 
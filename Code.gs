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
      .createMenu('My Add-on')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Opens a sidebar in the document.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('My Custom Sidebar');
  DocumentApp.getUi().showSidebar(html);
}

const HIGHLIGHT_COLOR = '#FFF8C4'; // Light yellow/orange

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
    muteHttpExceptions: true // Important to handle errors manually
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
      // Navigate the typical Gemini JSON structure
      const candidates = jsonResponse.candidates;
      if (candidates && candidates.length > 0 && candidates[0].content && candidates[0].content.parts && candidates[0].content.parts.length > 0) {
        const commentJsonString = candidates[0].content.parts[0].text;
        // The model should return a JSON string, parse it again
        const comments = JSON.parse(commentJsonString);
        // Basic validation
        if (Array.isArray(comments) && comments.every(c => typeof c === 'object' && 'quote' in c && 'comment' in c)) {
            console.log("Received comments from Gemini:", JSON.stringify(comments, null, 2));
            return comments;
        } else {
            console.error("Gemini response was not in the expected format (array of {quote, comment}):", commentJsonString);
            throw new Error("Gemini returned data in an unexpected format.");
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
      throw new Error("Could not parse the response from Gemini. " + e.message);
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
 *
 * @param {Array<string>} quotesToHighlight An array of quote strings to find and highlight.
 */
function highlightCommentsInDoc(quotesToHighlight) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    // Clear existing highlights first
    body.editAsText().setBackgroundColor(null);

    if (!quotesToHighlight || quotesToHighlight.length === 0) {
      console.log("No quotes provided for highlighting.");
      return;
    }

    console.log(`Attempting to highlight ${quotesToHighlight.length} quotes.`);
    const bodyText = body.getText();

    quotesToHighlight.forEach((quote, index) => {
      if (!quote || typeof quote !== 'string' || quote.trim() === '') {
          console.warn(`Skipping quote index ${index} due to empty or invalid value.`);
          return;
      }

      let searchResult = body.findText(quote);
      let count = 0;
      while (searchResult !== null) {
        const element = searchResult.getElement();
        const start = searchResult.getStartOffset();
        const end = searchResult.getEndOffsetInclusive();

        if (element.asText()) {
          element.asText().setBackgroundColor(start, end, HIGHLIGHT_COLOR);
          count++;
        }
        searchResult = body.findText(quote, searchResult);
      }
      if (count === 0) {
          console.warn(`Quote not found in document: "${quote}"`);
      }
    });
    console.log("Highlighting process completed.");
    DocumentApp.flush();
  } catch (e) {
    console.error("Error highlighting quotes in document: " + e.stack);
    throw new Error("Failed to highlight quotes in document: " + e.message);
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
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
 * Highlights the text segments (quotes) from the comments in the document.
 *
 * @param {Array<Object>} comments An array of comment objects { quote: string, comment: string }.
 */
function highlightCommentsInDoc(comments) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    // Clear existing highlights first (optional, but good for clean state)
    body.editAsText().setBackgroundColor(null); // Set background to transparent/default

    if (!comments || comments.length === 0) {
      console.log("No comments provided for highlighting.");
      return; 
    }

    console.log(`Attempting to highlight ${comments.length} comment quotes.`);
    const bodyText = body.getText(); // Get text once for efficiency

    comments.forEach((comment, index) => {
      const quote = comment.quote;
      if (!quote || typeof quote !== 'string' || quote.trim() === '') {
          console.warn(`Skipping comment index ${index} due to empty or invalid quote.`);
          return;
      }

      let searchResult = body.findText(quote);
      let count = 0;
      while (searchResult !== null) {
        const element = searchResult.getElement();
        const start = searchResult.getStartOffset();
        const end = searchResult.getEndOffsetInclusive();

        // Check if the found element is part of the body's text content
        if (element.asText()) {
          element.asText().setBackgroundColor(start, end, HIGHLIGHT_COLOR);
          count++;
        }
        // Find the next occurrence
        searchResult = body.findText(quote, searchResult);
      }
      if (count === 0) {
          console.warn(`Quote not found in document: "${quote}"`);
      }
    });
    console.log("Highlighting process completed.");
    DocumentApp.flush(); // Ensure changes are applied
  } catch (e) {
    console.error("Error highlighting comments in document: " + e.stack);
    // Rethrow the error so the client-side can handle it (e.g., in onHighlightError)
    throw new Error("Failed to highlight comments in document: " + e.message);
  }
}

/**
 * Finds the text ("quote") associated with the highlighted section
 * where the user's cursor is currently located.
 *
 * @return {string|null} The quote text if the cursor is in a highlighted area,
 *                      otherwise null.
 */
function getFocusedQuote() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const cursor = doc.getCursor();

    if (!cursor) {
      // No cursor, maybe focus is outside the doc editor
      return null;
    }

    const element = cursor.getElement();
    const offset = cursor.getOffset();

    // Check if the element containing the cursor is text
    if (!element || element.getType() !== DocumentApp.ElementType.TEXT) {
        // Cursor might be at the start/end of a paragraph, or on an image etc.
        // Try checking the element *before* the cursor if offset is 0
        if (offset === 0 && cursor.getSurroundingText().getText().length > 0) {
            // Likely at the beginning of a text element, check the background there
            const surroundingText = cursor.getSurroundingText();
            if (surroundingText.getBackgroundColor(0) === HIGHLIGHT_COLOR) {
                return findQuoteFromPosition(surroundingText, 0);
            }
        }
        // Otherwise, not in a text element or not at a checkable position
        return null;
    }

    const textElement = element.asText();
    
    // Check the background color *at* the cursor position.
    // Note: If the cursor is at the boundary between highlighted/non-highlighted,
    // this checks the character *preceding* the cursor offset.
    let checkOffset = offset > 0 ? offset - 1 : 0; 
    const bgColor = textElement.getBackgroundColor(checkOffset);

    if (bgColor !== HIGHLIGHT_COLOR) {
      // If the char before isn't highlighted, check the char *at* the cursor 
      // (if the cursor is not at the very end of the text element)
      if (offset < textElement.getText().length && textElement.getBackgroundColor(offset) === HIGHLIGHT_COLOR) {
           // Cursor is likely at the *start* of a highlighted section
           return findQuoteFromPosition(textElement, offset);
      } 
      // Not in a highlighted section
      return null;
    }
    
    // Cursor is within or at the end of a highlighted section
    return findQuoteFromPosition(textElement, checkOffset);

  } catch (e) {
    console.error("Error getting focused quote: " + e.stack);
    // Don't disrupt the user, just return null if there's an error
    return null;
  }
}

/**
 * Helper function to find the full highlighted quote based on a position
 * known to be within a highlighted text element.
 *
 * @param {Text} textElement The Google Apps Script Text element.
 * @param {number} knownHighlightOffset An offset within the textElement known to be highlighted.
 * @return {string|null} The full text of the highlighted segment, or null if error.
 */
function findQuoteFromPosition(textElement, knownHighlightOffset) {
    try {
        const text = textElement.getText();
        let start = knownHighlightOffset;
        let end = knownHighlightOffset;

        // Find the start of the highlight by searching backwards
        while (start > 0 && textElement.getBackgroundColor(start - 1) === HIGHLIGHT_COLOR) {
            start--;
        }
        // Ensure the start itself is highlighted (handles edge case where cursor is after last char)
        if(textElement.getBackgroundColor(start) !== HIGHLIGHT_COLOR) {
             console.warn("Could not confirm start highlight at offset", start);
             return null; // Should not happen if called correctly
        }

        // Find the end of the highlight by searching forwards
        while (end < text.length - 1 && textElement.getBackgroundColor(end + 1) === HIGHLIGHT_COLOR) {
            end++;
        }

        const quote = text.substring(start, end + 1);
        console.log(`Identified focused quote: "${quote}" from range ${start}-${end}`);
        return quote.trim(); // Trim whitespace just in case
    } catch(e) {
        console.error("Error in findQuoteFromPosition: " + e.stack);
        return null;
    }
} 
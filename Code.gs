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
 * @param {string} docContent The text content of the document.
 * @param {string} apiKey The user's Gemini API key.
 * @return {Array<Object>} An array of comment objects { quote: string, comment: string }.
 */
function getGeminiComments(docContent, apiKey) {
  if (!apiKey) {
    throw new Error("API Key is required to contact Gemini.");
  }
  if (!docContent || docContent.trim().length === 0) {
    console.log("Document content is empty, skipping Gemini call.");
    return []; // Return empty array if no content
  }

  const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;

  const prompt = `Please review the following document text and provide constructive comments. For each comment, identify the exact phrase or sentence from the text that the comment refers to. Present your output STRICTLY as a JSON array of objects, where each object has a "quote" key (containing the exact text phrase) and a "comment" key (containing your feedback). Do not include any text outside of the JSON array.

Example format:
[
  {
    "quote": "This is a sentence to comment on.",
    "comment": "This sentence could be clearer."
  },
  {
    "quote": "Another phrase needing feedback.",
    "comment": "Consider rephrasing this part."
  }
]

Document Text:
---
${docContent}
---`;

  const requestBody = {
    contents: [{
      parts: [{
        text: prompt
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
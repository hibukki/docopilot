# Docopilot

Proactively get comments on your Google Doc with Gemini.

## Development Setup (macOS)

1.  **Install Clasp:**
    ```bash
    npm install -g @google/clasp
    ```

2.  **Login to Google:**
    ```bash
    clasp login
    ```
    (Follow the browser authentication flow)

3.  **Push Code:**
    ```bash
    clasp push
    ```
    This uploads the current code to the associated Apps Script project.

## Usage

1.  Open the Google Doc associated with this Apps Script project (the one created when you first ran `clasp create` or cloned an existing project connected to this script).
2.  Look for the **Docopilot** menu item in the Google Docs menu bar.
3.  Select **Docopilot > Show sidebar**.
4.  Enter your Gemini API Key in the sidebar settings.
5.  Click **Analyze for Comments** or wait ~1.5 seconds after you stop editing the document for comments to be generated automatically. 
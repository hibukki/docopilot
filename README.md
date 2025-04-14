# Docopilot

Proactively get comments on your Google Doc with Gemini.

## Development Setup (macOS)

These steps assume you have [Node.js](https://nodejs.org/) and npm installed.

1.  **Install Clasp:**
    ```bash
    npm install -g @google/clasp
    ```

2.  **Login to Google:**
    ```bash
    clasp login
    ```
    (Follow the browser authentication flow)

3.  **Create/Clone Project:**
    *   **If starting new:** Create a new Doc and linked script project:
        ```bash
        # Replace "My Docopilot Doc" with your desired document title
        clasp create --title "My Docopilot Doc" --type DOCS
        ```
        This will output links to the created Google Doc and Apps Script project.
    *   **If using existing code:** Clone the repository and ensure a `.clasp.json` file pointing to your Apps Script project ID exists.

4.  **Push Code:**
    ```bash
    clasp push
    ```
    This uploads the code to the Apps Script project.

## Usage

1.  Open the Google Doc associated with this script project (use the link from `clasp create` or open the Doc connected to your cloned project).
2.  Look for the **Docopilot** menu item in the Google Docs menu bar.
3.  Select **Docopilot > Show sidebar**.
4.  Configure settings (API Key, Prompt) as needed within the sidebar.
5.  Comments will be generated automatically ~1.5 seconds after you stop editing, or you can click **Analyze for Comments**. 
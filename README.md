# Docopilot

Proactively get comments on your Google Doc with Gemini.

## Development Setup (macOS)

These steps assume you have [Node.js](https://nodejs.org/) and npm installed.

### Install Clasp

```bash
npm install -g @google/clasp
```

It's a tool that let's you write code in your IDE and push it to e.g be a Docs plugin.

### Login to Google

```bash
clasp login
```

This gives clasp permission to create a google doc, add a script (plugin) to it, and so on.

### Create Project

This will create one google doc that has an empty script attached.

```bash
# Replace "My Docopilot Doc" with your desired document title
clasp create --title "My Docopilot Doc" --type DOCS
```

#### Find the link to the Google Doc just created

Get in to it, just to make sure it works.

### Push Code

This will make the script that is attached to the google doc we just created contain the local code.

```bash
clasp push
```

### Refresh the google doc

When it loads, it will run the code that was just pushed.

### Open the new menu item "Docopilot"

In your Google Doc.

### Enter the settings

(e.g gemini api key)

### Put content in the doc

For Gemini to comment on

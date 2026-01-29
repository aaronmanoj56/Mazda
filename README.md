# Mazda Validator

A web application for validating and comparing Google Sheets data.

## Setup Instructions

### Option 1: Service Account (Recommended - No Token Needed!)

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Create a Google Service Account:**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing
   - Enable "Google Sheets API"
   - Go to "APIs & Services" → "Credentials"
   - Click "Create Credentials" → "Service Account"
   - Create a service account and download the JSON key file
   - Save it as `service-account.json` in this directory

3. **Share your Google Sheets:**
   - Open your Google Sheet
   - Click "Share" button
   - Add the service account email (found in the JSON file, looks like `xxx@xxx.iam.gserviceaccount.com`)
   - Give it "Editor" access
   - Click "Send"

4. **Start the server:**
   ```bash
   npm start
   ```

5. **Open the application:**
   - Open `http://localhost:3000` in your browser
   - Or if running from file, make sure the server is running and open `Index.html`

### Option 2: OAuth2 (Alternative)

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Set up OAuth2 credentials:**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create OAuth 2.0 Client ID
   - Set environment variables:
     ```bash
     export GOOGLE_CLIENT_ID="your-client-id"
     export GOOGLE_CLIENT_SECRET="your-client-secret"
     ```

3. **Start the server:**
   ```bash
   npm start
   ```

4. **Authorize (first time only):**
   - Visit `http://localhost:3000/auth`
   - Sign in and authorize
   - Close the window after authorization

5. **Use the application:**
   - Open `http://localhost:3000` in your browser

## Features

- **Compare Images**: Compare image URLs between two sheets
- **Find Blank Space**: Find and highlight blank cells in Google Sheets
- **No Manual Tokens**: Service account handles authentication automatically

## Troubleshooting

- **"Cannot connect to server"**: Make sure the Node.js server is running (`npm start`)
- **"Authentication not configured"**: Set up either Service Account or OAuth2 credentials
- **"Permission denied"**: Make sure you've shared the Google Sheet with the service account email

## Notes

- Service Account method is recommended as it requires no user interaction
- The server runs on port 3000 by default
- Service account JSON file should be kept secure and not committed to version control





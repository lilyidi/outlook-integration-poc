# Outlook Integration Proof of Concept (POC)

This project demonstrates how to integrate with Microsoft Outlook using the Microsoft Graph API. It provides functionalities for managing emails, setting up webhooks for notifications, and handling OAuth authentication.

## Table of Contents

- [Features](#features)
- [Getting Started](#getting-started)
- [Setup Instructions](#setup-instructions)
- [Usage](#usage)
- [API Endpoints](#api-endpoints)
- [Webhook Notifications](#webhook-notifications)
- [Contributing](#contributing)
- [License](#license)

## Features

- User authentication using OAuth 2.0
- Fetching and displaying emails
- Sending emails
- Setting up webhook subscriptions for email notifications
- Listing and managing subscriptions

## Getting Started

### Prerequisites

- Node.js (v14 or higher)
- npm (Node Package Manager)
- A Microsoft Azure account to register your application and obtain credentials

### Setup Instructions

1. **Clone the repository:**

   ```bash
   git clone https://github.com/vinodhkps/outlook-integration-poc.git
   cd outlook-integration-poc
   ```

2. **Install dependencies:**

   ```bash
   npm install
   ```

3. **Create a `.env` file** in the root directory and add your Microsoft app credentials:

   ```plaintext
   CLIENT_ID=your_client_id
   CLIENT_SECRET=your_client_secret
   REDIRECT_HOST=http://localhost:3000 (if testing in local, map it to ngrok or localtunnel)
   ```
   REDIRECT_HOST needs to map with what's configured in your app within azure.

4. **Run the application:**

   ```bash
   npm start
   ```

5. **Access the application** at `http://localhost:3000/login`.

## Usage

- **Login**: Navigate to `/login` to authenticate with your Microsoft account.
- **View Emails**: After logging in, you can view your emails at `/emails`.
- **Send Email**: Use the `/send-email` endpoint to send emails.
- **Manage Subscriptions**: Access `/list-subs` to view your webhook subscriptions.

## API Endpoints

- `GET /login`: Initiates the OAuth login process.
- `GET /auth/callback`: Handles the OAuth callback and retrieves access tokens.
- `GET /emails`: Fetches and displays the user's emails.
- `POST /send-email`: Sends an email using the authenticated user's account.
- `POST /list-subs`: Lists all webhook subscriptions for the user.
- `POST /extend-subs`: Renews existing subscriptions.

## Webhook Notifications

This application supports webhook notifications for email events. To set up a webhook:

1. Ensure your application is running and accessible from the internet (e.g., using ngrok).
2. Subscribe to email notifications by calling the appropriate endpoint to create a subscription.

Notifications will be sent to your specified `notificationUrl` when there are changes to your emails; you should be able to see in the console.


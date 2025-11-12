# Syncro Ticket Creator

A Tampermonkey userscript that enables creating Syncro MSP tickets from natural language descriptions using AI-powered parsing.

## Features

- **AI-Powered Parsing**: Uses OpenRouter API to intelligently extract ticket details from plain English descriptions
- **Syncro Integration**: Seamlessly creates tickets in your Syncro MSP instance
- **User-Friendly Interface**: Sidebar interface with dropdowns for organizations, users, and computers
- **Smart Matching**: Automatically finds and suggests customers and contacts based on AI-extracted information
- **Computer Assignment**: Optionally assign tickets to specific user computers/assets
- **Email Notifications**: Control whether email notifications are sent on ticket creation
- **Configurable AI Models**: Choose from various AI models for parsing (Claude, GPT, etc.)

## Prerequisites

Before installing and using this script, ensure you have:

1. **Tampermonkey Extension**: Install Tampermonkey for your browser:
   - [Chrome](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
   - [Firefox](https://addons.mozilla.org/en-US/firefox/addon/tampermonkey/)
   - [Safari](https://apps.apple.com/us/app/tampermonkey/id1482490089)
   - [Edge](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)

2. **OpenRouter Account**: Sign up at [https://openrouter.ai](https://openrouter.ai) and obtain an API key

3. **Syncro MSP Access**: You need a Syncro MSP account with API access enabled and an API key

4. **Browser Access**: The script works on Syncro MSP web interface (syncromsp.com)

## Installation

1. Install Tampermonkey extension for your browser (see Prerequisites)

2. Click the following link to install the script directly:
   - [Install Syncro Ticket Creator](syncro-ticket-creator.user.js)

   Or manually:
   - Copy the contents of `syncro-ticket-creator.user.js`
   - Open Tampermonkey dashboard
   - Create a new script
   - Paste the code and save

3. The script will automatically activate on Syncro MSP pages

## Configuration

After installation, configure the script with your API keys:

1. Navigate to any Syncro MSP page (e.g., https://your-subdomain.syncromsp.com)

2. Look for the "🎫 Create Ticket" tab on the left side of the screen

3. Click the gear icon (⚙️) in the ticket creator sidebar

4. Enter your API keys:
   - **OpenRouter API Key**: Your key from openrouter.ai
   - **Syncro API Key**: Your Syncro MSP API key
   - **Syncro Subdomain**: Your Syncro subdomain (e.g., "yourcompany")
   - **Default AI Model**: Choose your preferred AI model for parsing

5. Click "Save" to store your configuration

## Usage

1. Open your Syncro MSP dashboard

2. Click the "🎫 Create Ticket" tab on the left

3. Enter a natural language description of the ticket, for example:
   ```
   John from ABC Corp called about his computer not connecting to the network
   ```

4. Select your preferred AI model

5. Click "Send" to parse the description

6. Review the extracted information:
   - Organization
   - User/Contact
   - Computer (if applicable)
   - Subject
   - Issue description
   - Problem type

7. Make any necessary adjustments

8. Check "Send email notification" if desired

9. Click "Submit Ticket" to create the ticket in Syncro

## How It Works

The script uses AI to parse your natural language input and extract structured ticket information. It then:

1. Searches your Syncro customers and contacts to match the extracted names
2. Populates dropdowns with relevant options
3. Allows you to review and edit the parsed information
4. Creates the ticket via Syncro's API with all the specified details

## Troubleshooting

- **Script not appearing**: Ensure you're on a Syncro MSP page and Tampermonkey is enabled
- **API errors**: Verify your API keys are correct and have proper permissions
- **Parsing issues**: Try rephrasing your description or selecting a different AI model
- **Customer/User not found**: Check spelling and ensure the names exist in your Syncro instance

## Privacy & Security

- API keys are stored locally in your browser using Tampermonkey's secure storage
- No data is transmitted to third parties except OpenRouter and your Syncro instance
- The script only accesses Syncro data through official APIs

## Contributing

This script is open source. Feel free to submit issues or pull requests on GitHub.

## License

This project is released under the MIT License.

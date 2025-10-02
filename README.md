# TN-OfficeDebug

A tool to display various hidden settings / debug information in an Office.js window for Microsoft Word.

## Overview

This Office Add-in provides IT professionals and developers with comprehensive diagnostic and debug information about the Office environment, including:

- **Office Environment**: Host type, platform, version, language settings, and theme information
- **Document Information**: Document properties, metadata, author information, timestamps, and statistics
- **Add-in Information**: Current add-in details and settings (Note: Office Add-ins cannot enumerate other installed add-ins due to security restrictions)
- **System Information**: Browser details, memory usage (Chrome only), screen resolution, and performance metrics
- **Office Context**: Complete Office.js context object with all available API information

## Features

✨ **Real-time Information Display**: All information is gathered and displayed in real-time when the add-in is opened

🔄 **Refresh Capability**: Click the refresh button to reload all diagnostic information

🎨 **Modern UI**: Clean, professional interface with color-coded sections for easy reading

📊 **Comprehensive Data**: Displays all available diagnostic information accessible through Office.js APIs

## Project Structure

- **`/src`** - Source code files (HTML, CSS, JS)
- **`/config`** - Configuration files (manifest.xml, web.config)  
- **`/docs`** - Documentation and guides
- **`/assets`** - Static assets (icons, images)

See `PROJECT_STRUCTURE.md` for detailed information.

## Installation

### Option 1: Sideload in Office Online (Web)

1. Open Word Online (https://office.com/launch/word)
2. Create a new document
3. Go to **Insert** > **Add-ins** > **Upload My Add-in**
4. Click **Browse** and select the `config/manifest.xml` file
5. The add-in will appear in the ribbon

### Option 2: Sideload in Office Desktop

#### Windows:
1. Save all files to a local web server or use Office's debugging tools
2. Update the manifest.xml to point to your local server URL
3. Follow Microsoft's documentation for sideloading: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing

#### Mac:
1. Follow the same steps as Windows, but use the Mac-specific sideload process
2. See: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac

### Option 3: Using Office Add-in Debugging Tools (Development)

```bash
# Install dependencies
npm install

# Start the debugging server (requires local development server)
npm start

# Validate the manifest
npm run validate
```

## Usage

1. Open Microsoft Word (Online or Desktop)
2. Open the Office Debug Tool from the **Home** tab or Add-ins menu
3. The task pane will open showing all available debug information
4. Click the **🔄 Refresh** button to reload information at any time

## Debug Information Sections

### Office Environment
- Host application name (Word)
- Platform (Windows, Mac, iOS, Web, etc.)
- Display and content languages
- Office version and diagnostics
- Touch capability status
- Current Office theme colors

### Document Information
- Document URL and mode
- Title, subject, author, keywords
- Creation date, last save time, last print date
- Revision number and application name
- Character and paragraph counts

### Add-ins Information
- Current add-in settings availability
- Roaming settings count
- License information
- Manifest ID and version

### System Information
- Browser user agent and platform
- Online/offline status
- Screen dimensions and color depth
- Window dimensions
- JavaScript heap memory usage (Chrome only)
- Page load timing
- Current time and timezone

### Office Context Details
- Complete Office.js context object in JSON format
- All available API information
- Office.js library status

## File Structure

```
TN-OfficeDebug/
├── manifest.xml       # Office Add-in manifest
├── taskpane.html      # Main UI HTML
├── taskpane.css       # Styling
├── taskpane.js        # Main logic and data gathering
├── commands.html      # Ribbon commands (minimal)
├── package.json       # NPM package configuration
├── LICENSE            # MIT License
└── README.md          # This file
```

## Development

The add-in is built with:
- **Office.js**: Microsoft's JavaScript API for Office Add-ins
- **Vanilla JavaScript**: No frameworks required
- **Modern CSS**: Responsive design with gradient themes

To modify the add-in:
1. Edit the HTML/CSS/JS files as needed
2. Update the manifest.xml if changing functionality
3. Test using Office Online or Desktop sideloading

## Limitations

- **Add-in Enumeration**: Office Add-ins run in a sandboxed environment and cannot enumerate other installed add-ins for security reasons
- **Memory Information**: JavaScript heap memory info is only available in Chrome/Edge browsers
- **Platform Differences**: Some information may vary between Office Online and Desktop versions

## Browser Compatibility

- Microsoft Edge (recommended for Office Online)
- Google Chrome
- Safari (Mac)
- Internet Explorer 11 (limited support for Office Desktop)

## Support

For issues, feature requests, or contributions, please visit:
https://github.com/Toakan-Network/TN-OfficeDebug

## License

MIT License - See LICENSE file for details

Copyright (c) 2025 Toakan Network

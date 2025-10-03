# TN-OfficeDebug

A tool to display various hidden settings / debug information in an Office.js window for Microsoft Word.

## Overview

This Office Add-in provides customer support engineers and IT professionals with comprehensive diagnostic information for Microsoft Word troubleshooting:

- **Document Information**: Properties, custom properties, metadata, statistics, and security settings
- **Office Environment**: Version, platform, API support, license, and network status
- **System Information**: Browser details, security context, memory usage, and performance metrics
- **Add-in Information**: Project details, repository links, and troubleshooting resources
- **Debug Console**: Real-time logging for advanced diagnostics (restricted access)

## Features

✨ **Tabbed Interface**: Organized diagnostic information across Document Info, Office Environment, System Info, and Add-in Info tabs

� **Custom Properties Management**: Full CRUD support for Word document custom properties with green highlighting

🔒 **Domain-Based Security**: Debug tab restricted to @bighand.services domain users with multiple detection methods

🎨 **Card-Based Layout**: Professional UI with consistent styling to prevent text overflow and improve readability

🔄 **Real-Time Diagnostics**: Live information gathering optimized for customer support workflows

🛠️ **Comprehensive Error Handling**: Graceful degradation when Office.js APIs are unavailable

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
3. Go to **Home** > **Add-ins** > **More Settings** (or **Advanced**)
4. On the Office Add-ins dialog, select **Upload My Add-in**
5. Browse to the `config/manifest.xml` file and select **Upload**
6. The add-in will appear in the ribbon as "Show Debug Info"

### Option 2: Sideload in Office Desktop (Windows)

1. Place the `config/manifest.xml` file in a shared network folder
2. Open Word Desktop application
3. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings**
4. Select **Trusted Add-in Catalogs**
5. In the **Catalog Url** box, enter the full network path to your shared folder
6. Select **Add catalog** and check **Show in Menu**
7. Click **OK** to close dialogs and restart Word
8. Go to **Home** > **Add-ins** > **Advanced**
9. Choose **SHARED FOLDER** at the top of the Office Add-ins dialog
10. Select your add-in and choose **Add**

*Note: Requires HTTPS web server for manifest URLs and proper network share setup*

#### Mac:
1. Follow the same steps as Windows, but use the Mac-specific sideload process
2. See: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac

### Option 3: Development Sideloading

For development and testing:
1. Ensure the add-in files are served from a web server (IIS recommended)
2. Update manifest URLs to point to your development server
3. Use Office Online sideloading for quickest testing
4. Clear Office cache when making manifest changes

## Usage

1. Open Microsoft Word (Online or Desktop)
2. Click **Show Debug Info** button from the **Home** tab ribbon
3. The task pane opens with tabbed diagnostic information:
   - **Document Info**: Properties, custom properties, security settings
   - **Office Environment**: Version, platform, API support, license
   - **System Info**: Browser, security context, performance metrics
   - **Add-in Info**: Project details and troubleshooting resources
   - **Debug**: Advanced logging (restricted to @bighand.services users)

## Debug Information Tabs

### Document Info
- Document properties (title, author, dates, statistics)
- **Custom Properties**: Full CRUD management with green highlighting
- Security information (document mode, protection status)
- URL and basic document metadata

### Office Environment
- Office version, license, and API capability detection
- Platform information (Windows, Mac, Web)
- Network connectivity status
- Essential diagnostics for customer support

### System Info
- Browser and platform details
- Security context (HTTPS validation, storage availability)
- Performance metrics (memory usage, screen resolution)
- Critical system diagnostics

### Add-in Info
- Project information with clickable GitHub repository links
- Support resources and troubleshooting guides
- Add-in performance and load status
- Developer contact information

### Debug (Restricted Access)
- Real-time debug logging with window.logDebug() function
- Advanced diagnostics for technical troubleshooting
- Access restricted to @bighand.services domain users
- Multiple user detection methods with graceful fallbacks

## File Structure

```
TN-OfficeDebug/
├── src/
│   ├── taskpane.html      # Main UI with tabbed interface
│   ├── taskpane.css       # Card-based styling and layout
│   ├── taskpane.js        # Core logic and Office.js integration
│   ├── commands.html      # Ribbon command handlers
│   └── test-standalone.html # UI testing outside Office
├── config/
│   ├── manifest.xml       # Office Add-in manifest
│   └── web.config         # IIS web server configuration
├── assets/
│   └── gears.ico          # Custom add-in icon
├── docs/
│   ├── instructions.md    # Complete setup guide
│   └── USAGE.md          # Quick usage reference
├── .github/
│   └── copilot-instructions.md # AI coding guidelines
├── AGENTS.md             # AI agent context and patterns
├── PROJECT_STRUCTURE.md  # Detailed project organization
└── LICENSE               # MIT License
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

# TN-OfficeDebug - Project Structure

```
TN-OfficeDebug/
│
├── 📁 src/                     # Source code files
│   ├── taskpane.html          # Main add-in UI interface
│   ├── taskpane.js           # Core application logic
│   ├── taskpane.css          # UI styling and layout
│   ├── commands.html         # Office ribbon command handlers
│   └── test-standalone.html   # Standalone testing page
│
├── 📁 config/                  # Configuration files
│   ├── manifest.xml          # Office add-in manifest
│   └── web.config           # IIS web server configuration
│
├── 📁 docs/                    # Documentation
│   ├── instructions.md       # Complete setup and usage guide
│   └── USAGE.md             # Quick usage reference
│
├── 📁 assets/                  # Static assets (empty - for future use)
│   └── (icons, images, etc.)
│
├── 📁 .github/                # GitHub configuration
│   └── COMMIT_STRATEGY.md    # Git workflow guidelines
│
├── 📁 .vscode/                # VS Code settings
│   └── (workspace configuration)
│
├── index.html                 # Project landing page
├── README.md                 # Project overview
├── LICENSE                   # MIT license
├── .gitignore               # Git ignore rules
└── package.json             # Node.js dependencies (optional)
```

## Directory Purpose

### `/src` - Source Code
Contains all the active add-in files that are served to users:
- **taskpane.html**: Main user interface with tabbed layout
- **taskpane.js**: Core functionality and Office.js integration  
- **taskpane.css**: Styling with card-based responsive layout
- **commands.html**: Handlers for Office ribbon commands
- **test-standalone.html**: For testing UI outside of Office

### `/config` - Configuration
Contains deployment and server configuration:
- **manifest.xml**: Office add-in definition and permissions
- **web.config**: IIS server settings and CORS configuration

### `/docs` - Documentation
Complete project documentation:
- **instructions.md**: Comprehensive setup, troubleshooting, and development guide
- **USAGE.md**: Quick reference for common tasks

### `/assets` - Static Assets
Reserved for future static files like icons, images, or additional stylesheets

### `/.github` - GitHub Configuration
Repository management files:
- **COMMIT_STRATEGY.md**: Git workflow and commit guidelines

## File Dependencies

```
manifest.xml → References src/taskpane.html and src/commands.html
taskpane.html → Loads src/taskpane.css and src/taskpane.js  
web.config → Configures serving of all src/ files
```

## Development Workflow

1. **Edit source files** in `/src` directory
2. **Update configuration** in `/config` as needed
3. **Document changes** in `/docs`
4. **Test locally** using files in `/src`
5. **Deploy** by serving entire project structure

## Deployment Notes

- **Web server document root**: Point to project root directory
- **Add-in manifest**: Located at `/config/manifest.xml`
- **Entry points**: `/src/taskpane.html` and `/src/commands.html`
- **Configuration**: `/config/web.config` for IIS hosting
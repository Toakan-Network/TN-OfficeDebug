# TN-OfficeDebug - Office.js Diagnostic Tool

A comprehensive diagnostic and debugging tool for Microsoft Office Add-ins, specifically designed for Word integration with customer support capabilities.

## 📋 Project Overview

**TN-OfficeDebug** is a web-based Office.js add-in that provides detailed diagnostic information about the Office environment, document properties, system configuration, and add-in status. This tool is invaluable for customer support engineers and developers working with Office add-ins.

### Key Features

- **Document Information**: Document properties, statistics, custom properties, and security status
- **Office Environment**: Host details, API capabilities, licensing information, and version data
- **System Diagnostics**: Browser information, network status, security context, and performance metrics
- **Add-in Information**: Developer details, project information, troubleshooting data, and support resources
- **Domain-Based Security**: Debug tab visibility restricted to authorized domains
- **Real-time Debugging**: Comprehensive logging system for troubleshooting

## 🚀 Quick Start

### Prerequisites

- Microsoft Word (Desktop or Online)
- Web server (IIS, Apache, Node.js, or similar)
- HTTPS hosting for production deployment

### Installation Steps

1. **Clone the repository**
   ```bash
   git clone https://github.com/Toakan-Network/TN-OfficeDebug.git
   cd TN-OfficeDebug
   ```

2. **Configure hosting**
   
   **Option A: IIS (Recommended for Windows)**
   ```powershell
   # Enable IIS and required features
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole, IIS-WebServer, IIS-CommonHttpFeatures, IIS-HttpErrors, IIS-HttpLogging, IIS-RequestFiltering, IIS-StaticContent, IIS-ASPNET45
   
   # Create IIS site
   Import-Module WebAdministration
   New-Website -Name "TN-OfficeDebug" -Port 8080 -PhysicalPath "C:\dev\git\TN-OfficeDebug"
   ```
   
   **Option B: Node.js (Development)**
   ```bash
   npx http-server . -p 8080 --cors
   ```

3. **Update manifest URLs**
   - Edit `manifest.xml`
   - Replace placeholder URLs with your hosting domain
   - For local development: `http://localhost:8080`
   - For production: `https://your-domain.com`

4. **Install the add-in**
   - Open Microsoft Word
   - Go to Insert → Add-ins → Upload My Add-in
   - Select `manifest.xml`
   - The add-in will appear in the taskpane

## 📁 Project Structure

```
TN-OfficeDebug/
├── manifest.xml           # Office add-in manifest configuration
├── taskpane.html         # Main UI with tabbed interface
├── taskpane.js          # Core application logic and Office.js integration
├── taskpane.css         # Styling and responsive layout
├── commands.html        # Ribbon command handlers
├── index.html           # Landing page
├── web.config          # IIS configuration
├── package.json        # Node.js dependencies (optional)
├── instructions.md     # This documentation
└── README.md          # Project overview
```

## 🔧 Configuration

### Manifest Configuration

The `manifest.xml` file contains critical configuration:

```xml
<!-- Update these URLs to match your hosting -->
<bt:Url id="Commands.Url" DefaultValue="https://your-domain.com/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://your-domain.com/taskpane.html"/>
```

### Security Configuration

- **Domain Restriction**: Debug tab is only visible to users from `bighand.services` domain
- **HTTPS Requirement**: Production deployment requires HTTPS for security
- **CORS Headers**: Configured in `web.config` for cross-origin requests

### Web Server Configuration

**IIS (web.config)**:
```xml
<system.webServer>
  <httpProtocol>
    <customHeaders>
      <add name="Access-Control-Allow-Origin" value="*" />
      <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, OPTIONS" />
      <add name="Access-Control-Allow-Headers" value="Content-Type" />
    </customHeaders>
  </httpProtocol>
</system.webServer>
```

## 📊 Features Deep Dive

### Document Info Tab

**Core Document Properties**:
- Title, Subject, Author, Keywords, Comments
- Creation Date, Last Save Time, Last Print Date
- Application Name, Template, Revision Number

**Document Statistics**:
- Paragraph count
- Character count (approximate)
- Content controls detection

**Custom Properties**:
- Lists all custom document properties
- Shows property names and values
- Highlighted with green styling for easy identification

**Security Information**:
- Document URL and mode
- Content controls status
- Protection indicators

### Office Environment Tab

**Host Information**:
- Office host application (Word)
- Platform (Windows, Mac, Online)
- Display and content language

**Version & Licensing**:
- Office version and build numbers
- License type (subscription vs standalone)
- API capability detection

**API Support Matrix**:
- WordApi version support (1.1, 1.2, 1.3, 1.4+)
- DialogAPI availability
- SharedRuntime support
- RibbonAPI capabilities

### System Info Tab

**Browser Environment**:
- User agent string
- Platform and language detection
- Online/offline status

**Security Context**:
- HTTPS vs HTTP protocol
- Secure context validation
- Local storage availability
- Cookie support

**Performance Metrics**:
- JavaScript memory usage
- Screen resolution and window size
- Hardware capabilities

### Add-in Info Tab

**Developer Information**:
- Project details and version
- Developer contact information
- Repository links (clickable)

**Technical Stack**:
- Built with Office.js, HTML5, CSS3, JavaScript
- Target Office applications
- API requirements

**Troubleshooting Data**:
- Add-in load status and performance
- Security context validation
- Error monitoring status
- Common issue detection

**Support Resources**:
- Direct links to GitHub repository
- Issue reporting links
- Documentation references

### Debug Tab (Restricted Access)

**Access Control**:
- Only visible to users with `@bighand.services` email addresses
- Multiple detection methods for user authentication
- Fallback to manual domain verification
- Automatic hiding for unauthorized users

**Debug Features**:
- Real-time console logging with `window.logDebug()`
- Comprehensive error tracking
- Performance monitoring
- API call debugging

## 🛠️ Development & Customization

### Adding New Diagnostic Information

1. **Identify the appropriate tab** for your information
2. **Add to the relevant function** (`loadDocumentInfo`, `loadOfficeInfo`, etc.)
3. **Use the `createInfoRow()` helper** for consistent formatting:

```javascript
container.appendChild(createInfoRow('Label', 'Value'));
container.appendChild(createInfoRow('Long Label', 'Value', true)); // true for long labels
```

### Creating New Sections

```javascript
// Add a visual separator
const separator = document.createElement('div');
separator.style.marginTop = '20px';
separator.style.marginBottom = '15px';
separator.style.fontWeight = 'bold';
separator.style.borderTop = '1px solid #ccc';
separator.style.paddingTop = '10px';
separator.textContent = 'Section Title';
container.appendChild(separator);
```

### Error Handling Best Practices

```javascript
try {
  // Your Office API calls here
  const properties = context.document.properties;
  properties.load(['title', 'author']);
  await context.sync();
  
  container.appendChild(createInfoRow('Title', properties.title));
} catch (error) {
  window.logDebug('Error loading properties', { error: error.message });
  container.appendChild(createInfoRow('Title', 'Could not load'));
}
```

## 🔍 Troubleshooting Guide

### Common Issues & Solutions

#### 1. "Office is unable to start this add-in because it isn't set up properly"

**Cause**: Incorrect manifest URLs or web server not running

**Solution**:
1. Verify web server is running and accessible
2. Check manifest.xml URLs match your hosting
3. Ensure HTTPS for production deployment
4. Validate manifest XML syntax

#### 2. Document Info Tab Shows Errors

**Cause**: Using unavailable Office.js properties

**Solution**:
- Avoid `context.document.saved` and `context.document.isDirty`
- Always load properties before accessing: `properties.load(['propertyName'])`
- Use proper error handling with try/catch blocks

#### 3. Debug Tab Not Visible

**Cause**: Domain-based access control

**Solution**:
1. Check user email domain (must be `bighand.services`)
2. For development, ensure localhost hosting is detected
3. Use manual domain prompt as fallback

#### 4. Custom Properties Not Loading

**Cause**: Properties not properly loaded or empty

**Solution**:
```javascript
const customProperties = context.document.properties.customProperties;
customProperties.load(['items']);
await context.sync();

if (customProperties.items && customProperties.items.length > 0) {
  // Process properties
} else {
  // Handle empty case
}
```

#### 5. Network/CORS Issues

**Cause**: Cross-origin restrictions

**Solution**:
1. Configure proper CORS headers in web.config
2. Ensure HTTPS for production
3. Check browser developer tools for specific errors

### Development Tips

1. **Always test locally first**: Use localhost hosting for development
2. **Use browser developer tools**: Check console for JavaScript errors
3. **Enable debug logging**: Use `window.logDebug()` extensively
4. **Test in multiple environments**: Desktop Word, Word Online, different browsers
5. **Validate manifest**: Use Office Add-in Validator tools

### Performance Optimization

1. **Minimize API calls**: Load multiple properties in single calls
2. **Use efficient selectors**: Avoid repeated DOM queries
3. **Implement proper error boundaries**: Prevent single failures from breaking entire tab
4. **Cache static information**: Don't reload unchanging data

## 🔒 Security Considerations

### Production Deployment

1. **HTTPS Required**: Office Online requires HTTPS for add-ins
2. **Domain Validation**: Implement proper domain restrictions for sensitive features
3. **Content Security Policy**: Consider implementing CSP headers
4. **Input Sanitization**: Sanitize any user-provided data

### Access Control

- Debug features restricted to authorized domains
- Fallback security measures for unknown contexts
- Audit logging for access attempts

## 📝 Development History & Lessons Learned

### Key Development Challenges Resolved

#### 1. Manifest URL Configuration
**Problem**: Placeholder URLs in manifest caused load failures
**Solution**: Systematic replacement with proper hosting URLs
**Lesson**: Always validate manifest URLs match deployment environment

#### 2. Office.js API Compatibility
**Problem**: Properties like `context.document.saved` don't exist in Word API
**Solution**: Comprehensive API validation and error handling
**Lesson**: Thoroughly test all Office.js API calls, don't assume availability

#### 3. Web Server Configuration
**Problem**: CORS issues and SSL certificate conflicts with Node.js
**Solution**: Migration to IIS with proper CORS configuration
**Lesson**: IIS provides more stable hosting for enterprise environments

#### 4. Domain-Based Security
**Problem**: Auth API not available in Word add-ins
**Solution**: Multi-method user detection with graceful fallbacks
**Lesson**: Office.js APIs vary significantly between host applications

#### 5. UI Information Overload
**Problem**: Too much diagnostic information created poor user experience
**Solution**: Strategic organization and elimination of duplicates
**Lesson**: Customer support tools need scannable, organized information

### Technical Architecture Decisions

1. **Card-based Layout**: Prevents text overflow and improves readability
2. **Tab-based Organization**: Logical separation of different diagnostic areas
3. **Comprehensive Error Handling**: Ensures partial failures don't break entire tool
4. **Modular Function Design**: Each tab has dedicated load function for maintainability

### Best Practices Established

1. **Always validate Office.js API availability** before use
2. **Implement graceful fallbacks** for missing features
3. **Use consistent styling patterns** across all UI components
4. **Maintain comprehensive logging** for debugging
5. **Organize information logically** for customer support workflows

## 🚀 Deployment

### Local Development
```bash
# Start web server
npx http-server . -p 8080 --cors

# Access at http://localhost:8080
```

### Production Deployment
1. Upload files to web server
2. Configure HTTPS
3. Update manifest.xml with production URLs
4. Test manifest installation in Office

### Enterprise Deployment
1. Use Office 365 Admin Center for organization-wide deployment
2. Configure centralized hosting
3. Implement proper security policies
4. Set up monitoring and logging

## 📞 Support & Contributing

### Getting Help
- **Issues**: [GitHub Issues](https://github.com/Toakan-Network/TN-OfficeDebug/issues)
- **Documentation**: This file and README.md
- **AI Development**: AGENTS.md for AI agent context and patterns
- **Copilot Guidelines**: .github/copilot-instructions.md for GitHub Copilot
- **Discussions**: GitHub Discussions for questions and ideas

### Contributing
1. Fork the repository
2. Create feature branch
3. Make changes with proper testing
4. Submit pull request with clear description

### Reporting Issues
When reporting issues, please include:
- Office version and platform
- Browser information (if applicable)
- Steps to reproduce
- Error messages or screenshots
- Console logs if available

## 📄 License

MIT License - see LICENSE file for details

---

**Developed by Toakan Network**  
**Project Repository**: https://github.com/Toakan-Network/TN-OfficeDebug  
**Last Updated**: October 2025
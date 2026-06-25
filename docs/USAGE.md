# Office Debug Tool - Usage Guide

## Quick Start

### For Testing (Without Office)
1. Open `test-standalone.html` in your browser
2. View system information and basic diagnostics
3. This mode doesn't require Office or Office.js

### For Production (In Office Word)

#### Method 1: Office Online (Recommended for Quick Testing)
1. Go to https://office.com and sign in
2. Open Word Online
3. Create a new blank document
4. Click **Insert** → **Add-ins** → **Upload My Add-in**
5. Upload the `config/manifest.xml` file from this repository
6. Click **Show Debug Info** in the ribbon (Home tab)
7. View all diagnostic information in the task pane

#### Method 2: Office Desktop (Windows/Mac)
1. **Host the files**: You need to host the HTML/CSS/JS files on a web server
   - For local testing: Use `python3 -m http.server 8080` or any local server
   - For production: Deploy to a web hosting service
   
2. **Update `config/manifest.xml`**: Replace `~remoteAppUrl` with your server URL
   ```xml
   <SourceLocation DefaultValue="http://localhost:8080/taskpane.html"/>
   ```

3. **Sideload the add-in**:
   - **Windows**: Follow [Microsoft's sideloading guide for Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
   - **Mac**: Follow [Microsoft's sideloading guide for Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)

4. **Open Word** and look for the add-in in the ribbon

## Features Overview

### Office Environment Information
- Host application (Word)
- Platform (Windows, Mac, Web, etc.)
- Office version and build
- Display and content languages
- Touch capability
- Office theme colors

### Document Information
- Document URL and mode
- Title, author, subject, keywords
- Creation date, last save time
- Revision number
- Paragraph and character counts
- Template information

### Add-in Information
- Current add-in settings
- Roaming settings
- License information
- Manifest ID and version
- **Note**: Cannot enumerate other installed add-ins (security limitation)

### System Information
- Browser user agent
- Platform and language
- Online/offline status
- Screen and window dimensions
- JavaScript memory usage (Chrome/Edge only)
- Page load timing
- Current time and timezone

### Context Details
- Complete Office.js context object
- All available API information
- Diagnostic data

## Troubleshooting

### Add-in doesn't appear in Word
- Verify the manifest.xml is valid
- Check that the SourceLocation URLs are accessible
- Clear Office cache (AppData folder on Windows, ~/Library on Mac)
- Try uploading again

### "Loading..." persists
- Check browser console for errors (F12)
- Verify Office.js library is loading correctly
- Check network connectivity

### Memory info shows "Not available"
- This is normal for non-Chrome browsers
- Memory info only works in Chrome/Edge due to browser limitations

### Document info shows errors
- Ensure document is saved (not untitled)
- Some properties require the document to be saved to disk
- Try refreshing the information

## Development

### File Structure
```
manifest.xml          - Office Add-in manifest
taskpane.html         - Main UI
taskpane.css          - Styling
taskpane.js           - Logic and Office.js API calls
commands.html         - Ribbon commands support
test-standalone.html  - Standalone test (no Office required)
```

### Testing Changes
1. Edit HTML/CSS/JS files
2. Test in standalone mode: Open `test-standalone.html`
3. Test in Office: Reload the add-in or restart Office
4. Check browser console for errors

### Adding New Debug Information
1. Open `taskpane.js`
2. Find the relevant function (e.g., `loadOfficeInfo()`)
3. Add new `createInfoRow()` calls with your data
4. Test in both standalone and Office modes

## Security Notes

- The add-in runs in a sandboxed environment
- Cannot access other installed add-ins (by design)
- Cannot execute arbitrary system commands
- Only displays information available through Office.js APIs
- No data is sent to external servers

## Browser Compatibility

**Recommended Browsers:**
- Microsoft Edge (Chromium) - Full support
- Google Chrome - Full support
- Safari (Mac) - Good support
- Internet Explorer 11 - Limited support (Office Desktop only)

## Additional Resources

- [Office.js API Documentation](https://learn.microsoft.com/en-us/javascript/api/office)
- [Word JavaScript API](https://learn.microsoft.com/en-us/javascript/api/word)
- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Sideloading Guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins)

## License

MIT License - See LICENSE file for details

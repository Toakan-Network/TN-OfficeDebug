# PowerShell script to set up IIS for Office Add-in development
# Run as Administrator

Write-Host "Setting up IIS for TN-OfficeDebug Office Add-in..." -ForegroundColor Green

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

# Enable IIS features if not already enabled
Write-Host "Checking IIS installation..." -ForegroundColor Cyan

$iisFeatures = @(
    "IIS-WebServerRole",
    "IIS-WebServer", 
    "IIS-CommonHttpFeatures",
    "IIS-HttpErrors",
    "IIS-HttpLogging",
    "IIS-RequestFiltering",
    "IIS-StaticContent",
    "IIS-DefaultDocument",
    "IIS-DirectoryBrowsing",
    "IIS-ASPNET45",
    "IIS-NetFxExtensibility45",
    "IIS-ISAPIExtensions",
    "IIS-ISAPIFilter",
    "IIS-ManagementConsole"
)

foreach ($feature in $iisFeatures) {
    $status = Get-WindowsOptionalFeature -Online -FeatureName $feature
    if ($status.State -eq "Disabled") {
        Write-Host "Enabling $feature..." -ForegroundColor Yellow
        Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart
    }
}

# Import WebAdministration module
Import-Module WebAdministration -ErrorAction SilentlyContinue

# Variables
$siteName = "TN-OfficeDebug"
$sitePort = 8080
$sitePath = "C:\dev\git\TN-OfficeDebug"
$appPoolName = "TN-OfficeDebug-AppPool"

# Remove existing site if it exists
if (Get-Website -Name $siteName -ErrorAction SilentlyContinue) {
    Write-Host "Removing existing site: $siteName" -ForegroundColor Yellow
    Remove-Website -Name $siteName
}

# Remove existing app pool if it exists
if (Get-IISAppPool -Name $appPoolName -ErrorAction SilentlyContinue) {
    Write-Host "Removing existing app pool: $appPoolName" -ForegroundColor Yellow
    Remove-WebAppPool -Name $appPoolName
}

# Create new application pool
Write-Host "Creating application pool: $appPoolName" -ForegroundColor Cyan
New-WebAppPool -Name $appPoolName
Set-ItemProperty -Path "IIS:\AppPools\$appPoolName" -Name "processModel.identityType" -Value "ApplicationPoolIdentity"
Set-ItemProperty -Path "IIS:\AppPools\$appPoolName" -Name "managedRuntimeVersion" -Value ""

# Create new website
Write-Host "Creating website: $siteName" -ForegroundColor Cyan
New-Website -Name $siteName -Port $sitePort -PhysicalPath $sitePath -ApplicationPool $appPoolName

# Set MIME types for Office add-in files
Write-Host "Configuring MIME types..." -ForegroundColor Cyan
Add-WebConfigurationProperty -Filter "system.webServer/staticContent" -Name "." -Value @{fileExtension=".json"; mimeType="application/json"} -PSPath "IIS:\Sites\$siteName" -ErrorAction SilentlyContinue
Add-WebConfigurationProperty -Filter "system.webServer/staticContent" -Name "." -Value @{fileExtension=".xml"; mimeType="application/xml"} -PSPath "IIS:\Sites\$siteName" -ErrorAction SilentlyContinue

# Configure default document
Write-Host "Setting default documents..." -ForegroundColor Cyan
Set-WebConfiguration -Filter "system.webServer/defaultDocument/files" -Value @{value="index.html"} -PSPath "IIS:\Sites\$siteName"

# Enable directory browsing (helpful for development)
Set-WebConfiguration -Filter "system.webServer/directoryBrowse" -Value @{enabled="true"} -PSPath "IIS:\Sites\$siteName"

Write-Host "`n✅ IIS setup complete!" -ForegroundColor Green
Write-Host "📂 Site: $siteName" -ForegroundColor White
Write-Host "🌐 URL: http://localhost:$sitePort" -ForegroundColor White
Write-Host "📁 Path: $sitePath" -ForegroundColor White
Write-Host "`n📋 Next steps:" -ForegroundColor Yellow
Write-Host "1. Update your manifest.xml URLs to use http://localhost:$sitePort" -ForegroundColor White
Write-Host "2. Test the site: http://localhost:$sitePort/taskpane.html" -ForegroundColor White
Write-Host "3. Upload manifest.xml to Word (Insert > My Add-ins > Upload My Add-in)" -ForegroundColor White

# Test if the site is accessible
Write-Host "`n🧪 Testing site accessibility..." -ForegroundColor Cyan
try {
    $response = Invoke-WebRequest -Uri "http://localhost:$sitePort" -UseBasicParsing -TimeoutSec 5
    Write-Host "✅ Site is accessible (Status: $($response.StatusCode))" -ForegroundColor Green
} catch {
    Write-Host "⚠️  Could not test site accessibility: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   This is normal if the site takes a moment to start up" -ForegroundColor Gray
}

Write-Host "`n🎯 Ready to test your Office Add-in!" -ForegroundColor Green
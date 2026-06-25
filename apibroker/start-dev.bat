@echo off
echo Setting up NetDocuments Proxy API for development...

echo.
echo [1/4] Restoring NuGet packages...
dotnet restore

echo.
echo [2/4] Configuring user secrets...
dotnet user-secrets set "NetDocuments:ClientSecret" "iJ5RGF9eXXr86CFzl5kGbqEfT2RUIZf7GhXI8eHld1Km4gPY"
dotnet user-secrets set "Jwt:Key" "ThisIsAVeryLongSecureKeyForJWTTokenSigningPleaseChangeInProduction123456789"

echo.
echo [3/4] Building application...
dotnet build

echo.
echo [4/4] Starting development server...
echo.
echo API will be available at:
echo   - https://localhost:7001 (HTTPS)
echo   - http://localhost:5001 (HTTP)
echo   - Swagger UI: https://localhost:7001/swagger
echo.
echo Press Ctrl+C to stop the server
echo.

dotnet run
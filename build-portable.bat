@echo off
setlocal
cd /d "%~dp0"
where dotnet >nul 2>nul
if errorlevel 1 (
  echo .NET SDK 10 not found. This script requires .NET 10 SDK.
  pause
  exit /b 1
)
dotnet publish src\EmbodimentApiProbe\EmbodimentApiProbe.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)
echo.
echo Published to: %cd%\publish\EmbodimentApiProbe.exe
pause

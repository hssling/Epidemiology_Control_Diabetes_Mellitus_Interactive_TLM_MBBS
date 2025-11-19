@echo off
setlocal enabledelayedexpansion

echo Generating base64 data URIs for PNG images...

echo # Embedded image data URIs > embedded_images.py
echo embedded_images = { >> embedded_images.py

for %%f in (pathophysiology_diagram.png epidemiology_chart.png treatment_algorithm.png prevention_flowchart.png control_strategies_diagram.png national_program_diagram.png risk_factor_diagram.png) do (
    if exist "visualizations\%%f" (
        echo Processing %%f...

        REM Use PowerShell to convert to base64
        powershell -Command "$bytes = [System.IO.File]::ReadAllBytes('visualizations\%%f'); $b64 = [Convert]::ToBase64String($bytes); Write-Output \"    '../visualizations/%%f': 'data:image/png;base64,$b64',\"" >> embedded_images.py
    )
)

echo } >> embedded_images.py

echo.
echo Base64 encoding completed!
echo File created: embedded_images.py
echo.
echo Now open the HTML file and update image src attributes.
pause

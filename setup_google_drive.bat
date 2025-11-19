@echo off
echo 🚀 Setting up Google Drive Image Hosting for Diabetes TLM
echo ========================================================

echo 📝 Generating setup files...
echo.

if exist "C:\Python3*" (
    echo ✅ Python detected
) else (
    echo ⚠️  Python not found in PATH. Make sure Python 3 is installed and in PATH.
)

echo 📋 Creating Google Drive URL template...
echo # Google Drive URLs - Replace with your actual URLs > google_drive_urls.py
echo GOOGLE_DRIVE_URLS = { >> google_drive_urls.py
echo     "pathophysiology_diagram.png": "https://drive.google.com/uc?export=view&id=YOUR_FILE_ID_1", >> google_drive_urls.py
echo     "epidemiology_chart.png": "https://drive.google.com/uc?export=view&id=YOUR_FILE_ID_2", >> google_drive_urls.py
echo     "treatment_algorithm.png": "https://drive.google.com/uc?export=view&id=YOUR_FILE_ID_3", >> google_drive_urls.py
echo     "prevention_flowchart.png": "https://drive.google.com/uc?export=view&id=YOUR_FILE_ID_4", >> google_drive_urls.py
echo     "national_program_diagram.png": "https://drive.google.com/uc?export=view&id=YOUR_FILE_ID_5", >> google_drive_urls.py
echo     "control_strategies_diagram.png": "https://drive.google.com/uc?export=view&id=YOUR_FILE_ID_6" >> google_drive_urls.py
echo } >> google_drive_urls.py

echo ✅ Created: google_drive_urls.py

echo 📋 Creating batch file for HTML update...
echo echo Processing Google Drive URLs... > update_html_with_google_drive.bat
echo python google_drive_helper.py --update-html >> update_html_with_google_drive.bat
echo echo Done! Check for *_googledrive.html file >> update_html_with_google_drive.bat
echo pause >> update_html_with_google_drive.bat

echo ✅ Created: update_html_with_google_drive.bat

echo.
echo 🔗 To complete setup:
echo 1. Upload images in visualizations\ folder to Google Drive
echo 2. Get shareable links for each image
echo 3. Convert to direct URLs (replace /file/d/ID/view with /uc?export=view^&id=ID)
echo 4. Edit google_drive_urls.py with your direct URLs
echo 5. Run: update_html_with_google_drive.bat
echo.
echo 📂 Files created:
echo    - google_drive_urls.py (template to edit)
echo    - update_html_with_google_drive.bat (run after editing)
echo    - google_drive_helper.py (main processing script)
echo.
echo 🌐 Opening browser with setup instructions...
start google_drive_solution.html

echo 🎯 Setup complete! Follow the instructions in the opened browser window.
echo.
pause

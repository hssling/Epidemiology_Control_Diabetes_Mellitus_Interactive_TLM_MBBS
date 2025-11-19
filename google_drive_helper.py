#!/usr/bin/env python3
"""
Google Drive Image Link Generator
Automatically create Google Drive shareable links and update HTML
"""
import os
import re
from pathlib import Path
import json

class GoogleDriveHelper:
    def __init__(self, images_dir="visualizations", html_file="interactive/diabetes_interactive_tlm.html"):
        self.images_dir = Path(images_dir)
        self.html_file = Path(html_file)
        self.image_files = [
            "pathophysiology_diagram.png",
            "epidemiology_chart.png",
            "treatment_algorithm.png",
            "prevention_flowchart.png",
            "national_program_diagram.png",
            "control_strategies_diagram.png"
        ]

    def convert_drive_link_to_direct(self, shareable_link):
        """Convert Google Drive shareable link to direct image URL"""
        # Extract file ID from shareable link
        match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', shareable_link)
        if match:
            file_id = match.group(1)
            return f"https://drive.google.com/uc?export=view&id={file_id}"
        return shareable_link

    def generate_url_template(self):
        """Generate template for user to fill with their Google Drive links"""
        template = """# Copy this template and replace the FILE_IDs with your actual Google Drive file IDs
# Get file IDs from your Google Drive shareable links:
# https://drive.google.com/file/d/FILE_ID_HERE/view?usp=sharing

GOOGLE_DRIVE_URLS = {
    "pathophysiology_diagram.png": "https://drive.google.com/uc?export=view&id=YOUR_PATHOPHYSIOLOGY_FILE_ID",
    "epidemiology_chart.png": "https://drive.google.com/uc?export=view&id=YOUR_EPIGENOMICS_FILE_ID",
    "treatment_algorithm.png": "https://drive.google.com/uc?export=view&id=YOUR_TREATMENT_FILE_ID",
    "prevention_flowchart.png": "https://drive.google.com/uc?export=view&id=YOUR_PREVENTION_FILE_ID",
    "national_program_diagram.png": "https://drive.google.com/uc?export=view&id=YOUR_NPCDCS_FILE_ID",
    "control_strategies_diagram.png": "https://drive.google.com/uc?export=view&id=YOUR_CONTROL_FILE_ID"
}

# Example: If your shareable link is:
# https://drive.google.com/file/d/1ABC123def456/view?usp=sharing
# Then your direct URL becomes:
# https://drive.google.com/uc?export=view&id=1ABC123def456
"""
        return template

    def update_html_with_drive_urls(self, url_mapping):
        """Update HTML file with Google Drive URLs"""
        if not self.html_file.exists():
            raise FileNotFoundError(f"HTML file not found: {self.html_file}")

        with open(self.html_file, 'r', encoding='utf-8') as f:
            html_content = f.read()

        replacements_made = 0

        # Convert Google Drive shareable links to direct links if needed
        processed_urls = {}
        for image_name, url in url_mapping.items():
            if url and url.strip():
                processed_urls[image_name] = self.convert_drive_link_to_direct(url.strip())

        print("🔄 Updating HTML with Google Drive URLs...")
        print("Images to process:", len(processed_urls))

        # Replace standard img tags
        for image_name, drive_url in processed_urls.items():
            # Pattern for src attributes
            src_pattern = re.compile(rf'src="[^"]*{re.escape(image_name)}"', re.IGNORECASE)
            html_content = src_pattern.sub(f'src="{drive_url}"', html_content)

            # Count replacements
            matches = src_pattern.findall(html_content)
            replacements_made += len(matches) if matches else 0

            if matches:
                print(f"✅ Updated src for: {image_name}")

        # Replace placeholder divs (specific to Diabetes TLM structure)
        for image_name, drive_url in processed_urls.items():
            placeholder_img = f'<img src="{drive_url}" alt="{image_name.replace("_", " ").replace(".png", "").title()}" style="width: 100%; border-radius: 8px;">'

            if image_name == "pathophysiology_diagram.png":
                pattern = r'<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#9b59b6[^"]*">[\s\S]*?<\/div>'
                html_content = re.sub(pattern, f'<div style="margin-bottom: 30px;">{placeholder_img}</div>', html_content, flags=re.DOTALL | re.IGNORECASE)
                replacements_made += 1
                print(f"🔄 Replaced pathophysiology placeholder")

            elif image_name == "epidemiology_chart.png":
                pattern = r'<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">(?![\s\S]*?(?:Prevention|NPCDCS|Control))[\s\S]*?<\/div>'
                html_content = re.sub(pattern, f'<div style="margin-bottom: 20px;">{placeholder_img}</div>', html_content, flags=re.DOTALL | re.IGNORECASE)
                replacements_made += 1
                print(f"🔄 Replaced epidemiology placeholder")

            elif image_name == "treatment_algorithm.png":
                pattern = r'<div[^>]*>[\s\S]*?Management Tab Diagram[\s\S]*?<\/div>\s*<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">[\s\S]*?<\/div>'
                html_content = re.sub(pattern, f'<div>Management Tab Diagram</div><div style="margin-bottom: 20px;">{placeholder_img}</div>', html_content, flags=re.DOTALL | re.IGNORECASE)
                replacements_made += 1
                print(f"🔄 Replaced treatment algorithm placeholder")

            elif image_name == "prevention_flowchart.png":
                pattern = r'<div[^>]*style="[^"]*margin-top:\s*20px[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*"[^>]*>[\s\S]*?Prevention Flowchart[\s\S]*?<\/div>'
                html_content = re.sub(pattern, f'<div style="margin-top: 20px;">{placeholder_img}</div>', html_content, flags=re.DOTALL | re.IGNORECASE)
                replacements_made += 1
                print(f"🔄 Replaced prevention flowchart placeholder")

            elif image_name == "national_program_diagram.png":
                pattern = r'<div[^>]*style="[^"]*margin-top:\s*20px[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*"[^>]*>[\s\S]*?NPCDCS[\s\S]*?<\/div>'
                html_content = re.sub(pattern, f'<div style="margin-top: 20px;">{placeholder_img}</div>', html_content, flags=re.DOTALL | re.IGNORECASE)
                replacements_made += 1
                print(f"🔄 Replaced NPCDCS diagram placeholder")

            elif image_name == "control_strategies_diagram.png":
                pattern = r'<div[^>]*>[\s\S]*?Control Strategies[\s\S]*?<\/div>\s*<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">[\s\S]*?<\/div>'
                html_content = re.sub(pattern, f'<div>Control Strategies</div><div style="margin-bottom: 20px;">{placeholder_img}</div>', html_content, flags=re.DOTALL | re.IGNORECASE)
                replacements_made += 1
                print(f"🔄 Replaced control strategies placeholder")

        # Save updated HTML
        output_file = self.html_file.parent / f"{self.html_file.stem}_googledrive{self.html_file.suffix}"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

        print(f"📊 Replacements made: {replacements_made}")
        print("🎉 Successfully updated HTML with Google Drive URLs!")
        print(f"Output file: {output_file}")

        return output_file, replacements_made

    def generate_instructions_html(self):
        """Generate HTML instructions for Google Drive setup"""
        instructions = f"""<!DOCTYPE html>
<html>
<head>
    <title>Google Drive Setup Instructions - Diabetes TLM</title>
    <style>
        body {{ font-family: Arial; margin: 40px; background: #f5f5f5; }}
        .container {{ max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; }}
        .step {{ margin: 20px 0; padding: 20px; border-left: 4px solid #4285f4; }}
        .code {{ background: #f0f0f0; padding: 10px; border-radius: 4px; font-family: monospace; }}
        .image-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin: 20px 0; }}
        .image-card {{ border: 1px solid #ddd; padding: 10px; border-radius: 5px; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>🚀 Google Drive Image Hosting Setup</h1>
        <p><strong>Total Images:</strong> {len(self.image_files)}</p>
        <p><strong>Total Size:</strong> ~2.7 MB</p>

        <div class="step">
            <h2>1. Create Google Drive Folder</h2>
            <ul>
                <li>Go to <a href="https://drive.google.com" target="_blank">Google Drive</a></li>
                <li>Create new folder: "diabetes_visualizations"</li>
                <li>Upload these {len(self.image_files)} PNG files:
                    <ul>
"""

        for img in self.image_files:
            instructions += f"                        <li>{img}</li>\n"

        instructions += """
                    </ul>
                </li>
            </ul>
        </div>

        <div class="step">
            <h2>2. Get Shareable Links</h2>
            <p>For each uploaded image:</p>
            <ol>
                <li>Right-click image → "Get shareable link"</li>
                <li>Click "Change" → "Anyone with the link"</li>
                <li>Copy the link that looks like: <code>https://drive.google.com/file/d/FILE_ID/view?usp=sharing</code></li>
            </ol>
        </div>

        <div class="step">
            <h2>3. Convert to Direct URLs</h2>
            <p>Replace <code>/file/d/FILE_ID/view?usp=sharing</code> with <code>/uc?export=view&id=FILE_ID</code></p>

            <div class="code">
Example:<br>
❌ Shareable: https://drive.google.com/file/d/1ABC123def456/view?usp=sharing<br>
✅ Direct: https://drive.google.com/uc?export=view&id=1ABC123def456
            </div>
        </div>

        <div class="step">
            <h2>4. Update Python Script</h2>
            <p>Edit <code>google_drive_urls.py</code> with your direct URLs:</p>
            <div class="code">
GOOGLE_DRIVE_URLS = {
    "pathophysiology_diagram.png": "YOUR_DIRECT_URL_HERE",
    "epidemiology_chart.png": "YOUR_DIRECT_URL_HERE",
    # ... add all 6 URLs
}
            </div>
        </div>

        <div class="step">
            <h2>5. Run Update Script</h2>
            <div class="code">
python google_drive_helper.py --update-html
            </div>
        </div>

        <div class="image-grid">
"""

        for img in self.image_files:
            if (self.images_dir / img).exists():
                size = (self.images_dir / img).stat().st_size / 1024  # KB
                instructions += f"""            <div class="image-card">
                <img src="visualizations/{img}" style="width: 100%; border-radius: 5px;">
                <strong>{img.replace('_', ' ').replace('.png', '').title()}</strong>
                <br><small>{size:.0f} KB</small>
            </div>
"""

        instructions += """        </div>
    </div>
</body>
</html>"""

        with open('google_drive_setup_instructions.html', 'w', encoding='utf-8') as f:
            f.write(instructions)

        print("📋 Generated setup instructions: google_drive_setup_instructions.html")

def main():
    print("🚀 Google Drive Image Helper for Diabetes TLM")
    print("=" * 50)

    helper = GoogleDriveHelper()

    # Generate template for user
    template = helper.generate_url_template()

    with open('google_drive_urls_template.py', 'w', encoding='utf-8') as f:
        f.write(template)

    print("📝 Generated template file: google_drive_urls_template.py")

    # Generate setup instructions
    helper.generate_instructions_html()

    print("\n📋 Next steps:")
    print("1. Upload images to Google Drive")
    print("2. Get shareable links for each image")
    print("3. Convert links to direct URLs")
    print("4. Edit google_drive_urls_template.py")
    print("5. Run: python google_drive_helper.py --update-html")

if __name__ == "__main__":
    main()

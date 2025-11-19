#!/usr/bin/env python3
"""
Auto Embed Images Tool
Automatically embeds PNG images as base64 data URIs into HTML files
"""
import os
import base64
import re
from pathlib import Path

class ImageEmbedder:
    def __init__(self, images_dir="visualizations", html_file="interactive/diabetes_interactive_tlm.html"):
        self.images_dir = Path(images_dir)
        self.html_file = Path(html_file)
        self.base64_cache = {}

    def encode_image_to_base64(self, image_path):
        """Convert image to base64 data URI"""
        if image_path in self.base64_cache:
            return self.base64_cache[image_path]

        try:
            with open(image_path, 'rb') as f:
                image_data = f.read()
            encoded = base64.b64encode(image_data).decode('utf-8')
            data_uri = f"data:image/png;base64,{encoded}"
            self.base64_cache[image_path] = data_uri
            return data_uri
        except Exception as e:
            print(f"Error encoding {image_path}: {e}")
            return None

    def get_image_mappings(self):
        """Get mapping of image filenames to their data URIs"""
        mappings = {}
        if self.images_dir.exists():
            for png_file in self.images_dir.glob("*.png"):
                data_uri = self.encode_image_to_base64(png_file)
                if data_uri:
                    mappings[png_file.name] = data_uri
        return mappings

    def create_embedded_html(self, output_file=None):
        """Create HTML with embedded images"""
        if not self.html_file.exists():
            print(f"HTML file not found: {self.html_file}")
            return False

        # Read HTML content
        with open(self.html_file, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # Get image mappings
        image_mappings = self.get_image_mappings()
        replacements_made = 0

        print("Processing HTML file...")

        # Pattern to match image references in HTML
        # This will match various formats of image references
        patterns = [
            # Standard img tags
            (r'<img[^>]*src="([^"]*\.png)"[^>]*>', 'src'),
            # CSS background-image
            (r'background-image:\s*url\("([^"]*\.png)"\)', 'url'),
            # Divs with data-image attribute or image file references
        ]

        for pattern, attr in patterns:
            def replace_match(match):
                nonlocal replacements_made
                full_match = match.group(0)
                image_path = match.group(1)

                # Extract just the filename if full path
                image_name = os.path.basename(image_path)

                if image_name in image_mappings:
                    if attr == 'src':
                        new_match = full_match.replace(f'src="{image_path}"', f'src="{image_mappings[image_name]}"')
                    elif attr == 'url':
                        new_match = full_match.replace(f'url("{image_path}")', f'url("{image_mappings[image_name]}")')

                    replacements_made += 1
                    print(f"Embedded: {image_name}")
                    return new_match
                return full_match

            html_content = re.sub(pattern, replace_match, html_content, flags=re.IGNORECASE)

        # Handle placeholder divs that reference images by name
        placeholder_patterns = [
            # Pathophysiology diagram
            (r'<div[^>]*style="[^"]*background: #f8f9fa[^"]*padding: 20px[^"]*border-radius: 8px[^"]*text-align: center[^"]*border: 2px dashed #9b59b6[^"]*">(.*?)</div>',
             'pathophysiology_diagram.png'),
            # Epidemiology chart
            (r'<div[^>]*style="[^"]*background: #f8f9fa[^"]*padding: 20px[^"]*border-radius: 8px[^"]*text-align: center[^"]*border: 2px dashed #3498db[^"]*">(.*?)</div>',
             'epidemiology_chart.png'),
            # Management/Treatment algorithm
            (r'<div[^>]*>Management Tab Diagram.*?</div>.*?<div[^>]*style="[^"]*background: #f8f9fa[^"]*padding: 20px[^"]*border-radius: 8px[^"]*text-align: center[^"]*border: 2px dashed #3498db[^"]*">(.*?)</div>',
             'treatment_algorithm.png'),
            # Prevention flowchart
            (r'<div[^>]*style="[^"]*margin-top: 20px[^"]*background: #f8f9fa[^"]*padding: 20px[^"]*border-radius: 8px[^"]*text-align: center[^"]*border: 2px dashed #3498db[^"]*"[^>]*>(.*?Prevention Flowchart.*?)</div>',
             'prevention_flowchart.png'),
            # NPCDCS diagram
            (r'<div[^>]*style="[^"]*margin-top: 20px[^"]*background: #f8f9fa[^"]*padding: 20px[^"]*border-radius: 8px[^"]*text-align: center[^"]*border: 2px dashed #3498db[^"]*"[^>]*>(.*?NPCDCS.*?)</div>',
             'national_program_diagram.png'),
            # Control strategies
            (r'<div[^>]*>Control Strategies.*?</div>.*?<div[^>]*style="[^"]*background: #f8f9fa[^"]*padding: 20px[^"]*border-radius: 8px[^"]*text-align: center[^"]*border: 2px dashed #3498db[^"]*">(.*?)</div>',
             'control_strategies_diagram.png'),
        ]

        for pattern, image_name in placeholder_patterns:
            def replace_placeholder(match):
                nonlocal replacements_made
                if image_name in image_mappings:
                    replacement = f'<div style="margin-bottom: 20px;"><img src="{image_mappings[image_name]}" alt="{image_name.replace("_", " ").replace(".png", "").title()}" style="width: 100%; border-radius: 8px;"></div>'
                    replacements_made += 1
                    print(f"Replaced placeholder with: {image_name}")
                    return replacement
                return match.group(0)

            html_content = re.sub(pattern, replace_placeholder, html_content, flags=re.DOTALL | re.IGNORECASE)

        # Output file
        if output_file is None:
            output_file = self.html_file.parent / f"{self.html_file.stem}_embedded{self.html_file.suffix}"

        # Write the modified HTML
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

        file_size = len(html_content.encode('utf-8')) / (1024 * 1024)  # Size in MB
        print(f"\n{'='*50}")
        print("EMBEDDING COMPLETE!")
        print(f"✅ Processed: {len(image_mappings)} images")
        print(f"✅ Replacements made: {replacements_made}")
        print(f"✅ Output file: {output_file}")
        print(f"📁 File size: {file_size:.2f} MB")
        print("✅ Ready for offline use and Google Sites upload!"
        print(f"{'='*50}\n")

        return True

def main():
    print("🖼️  DIABETES MELLITUS - AUTO IMAGE EMBEDDER")
    print("=" * 50)

    embedder = ImageEmbedder()
    embedder.create_embedded_html()

if __name__ == "__main__":
    main()

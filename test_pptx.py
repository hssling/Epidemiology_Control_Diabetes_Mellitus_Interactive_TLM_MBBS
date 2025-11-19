#!/usr/bin/env python3
"""
Script to embed PNG images as base64 data URIs in HTML
"""

import base64
import os

def png_to_data_uri(png_path):
    """Convert PNG file to base64 data URI"""
    try:
        with open(png_path, 'rb') as file:
            png_data = file.read()
            b64_data = base64.b64encode(png_data).decode('utf-8')
            return f'data:image/png;base64,{b64_data}'
    except Exception as e:
        print(f"Error converting {png_path}: {e}")
        return None

def main():
    print("=== Converting PNG Images to Base64 Data URIs ===")

    # List of PNG files to convert
    png_files = [
        'visualizations/pathophysiology_diagram.png',
        'visualizations/epidemiology_chart.png',
        'visualizations/treatment_algorithm.png',
        'visualizations/prevention_flowchart.png',
        'visualizations/control_strategies_diagram.png',
        'visualizations/national_program_diagram.png',
        'visualizations/risk_factor_diagram.png'
    ]

    # Dictionary to store the data URIs
    embedded_images = {}

    for png_file in png_files:
        if os.path.exists(png_file):
            print(f"✅ Converting {png_file}...")
            data_uri = png_to_data_uri(png_file)
            if data_uri:
                # Store with the HTML image path
                html_path = f"../{png_file}"
                embedded_images[html_path] = data_uri
                print(f"   Size: {len(data_uri)} characters")
        else:
            print(f"❌ File not found: {png_file}")

    # Generate JavaScript object for HTML
    print(f"\n🎯 Generated {len(embedded_images)} embedded images")

    # Create a JavaScript file for use in HTML
    js_content = """// Embedded image data URIs for standalone HTML
const embeddedImages = {
"""

    for html_path, data_uri in embedded_images.items():
        js_content += f'    "{html_path}": "{data_uri}",\n'

    js_content += """};\n\n// Function to replace image src with embedded data
function embedImages() {
    const images = document.querySelectorAll('img');
    images.forEach(img => {
        const src = img.getAttribute('src');
        if (embeddedImages[src]) {
            img.src = embeddedImages[src];
        }
    });
}

// Embed images when page loads
document.addEventListener('DOMContentLoaded', embedImages);
"""

    with open('embedded_images.js', 'w', encoding='utf-8') as f:
        f.write(js_content)

    print("✅ JavaScript file created: embedded_images.js")
    print("\n📋 To make HTML standalone:")
    print("1. Add: <script src=\"embedded_images.js\"></script>")
    print("2. The script will automatically replace image src with embedded data")
    print("\n🎉 Your HTML will now work as a single standalone file!")

if __name__ == "__main__":
    main()

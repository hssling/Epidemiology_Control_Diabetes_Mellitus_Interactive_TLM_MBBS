#!/usr/bin/env node

/**
 * Diabetes HTML Image Embedder - Node.js Version
 * Automatically embeds PNG images as base64 data URIs into HTML files
 */

const fs = require('fs');
const path = require('path');
const { promisify } = require('util');

const readFile = promisify(fs.readFile);
const writeFile = promisify(fs.writeFile);
const readdir = promisify(fs.readdir);
const stat = promisify(fs.stat);

class ImageEmbedder {
    constructor(imagesDir = 'visualizations', htmlFile = 'interactive/diabetes_interactive_tlm.html') {
        this.imagesDir = path.resolve(imagesDir);
        this.htmlFile = path.resolve(htmlFile);
        this.base64Cache = new Map();
    }

    async encodeImageToBase64(imagePath) {
        if (this.base64Cache.has(imagePath)) {
            return this.base64Cache.get(imagePath);
        }

        try {
            const imageBuffer = await readFile(imagePath);
            const base64Data = imageBuffer.toString('base64');
            const dataUri = `data:image/png;base64,${base64Data}`;
            this.base64Cache.set(imagePath, dataUri);
            return dataUri;
        } catch (error) {
            console.error(`❌ Error encoding ${imagePath}:`, error.message);
            return null;
        }
    }

    async getImageMappings() {
        const mappings = {};

        try {
            const files = await readdir(this.imagesDir);
            const pngFiles = files.filter(file => file.endsWith('.png'));

            console.log(`🔍 Found ${pngFiles.length} PNG files in ${this.imagesDir}`);

            for (const pngFile of pngFiles) {
                const fullPath = path.join(this.imagesDir, pngFile);
                const dataUri = await this.encodeImageToBase64(fullPath);
                if (dataUri) {
                    mappings[pngFile] = dataUri;
                    console.log(`✅ Encoded: ${pngFile}`);
                }
            }
        } catch (error) {
            console.error(`❌ Error reading images directory: ${error.message}`);
        }

        return mappings;
    }

    async getFileSize(filePath) {
        try {
            const stats = await stat(filePath);
            return (stats.size / (1024 * 1024)).toFixed(2);
        } catch (error) {
            return '0.00';
        }
    }

    replacePlaceholderDivs(content, imageMappings) {
        let replacementsMade = 0;

        // Define placeholder patterns with their corresponding image filenames
        const placeholderPatterns = [
            // Pathophysiology diagram
            {
                pattern: /<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#9b59b6[^"]*">[\s\S]*?<\/div>/g,
                image: 'pathophysiology_diagram.png'
            },
            // Epidemiology chart
            {
                pattern: /<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">(?![\s\S]*?(?:Prevention|NPCDCS|Control))[\s\S]*?<\/div>/g,
                image: 'epidemiology_chart.png'
            },
            // Management/Treatment algorithm
            {
                pattern: /<div[^>]*>[\s\S]*?Management Tab Diagram[\s\S]*?<\/div>\s*<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">[\s\S]*?<\/div>/g,
                image: 'treatment_algorithm.png'
            },
            // Prevention flowchart
            {
                pattern: /<div[^>]*style="[^"]*margin-top:\s*20px[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*"[^>]*>[\s\S]*?Prevention Flowchart[\s\S]*?<\/div>/g,
                image: 'prevention_flowchart.png'
            },
            // NPCDCS diagram
            {
                pattern: /<div[^>]*style="[^"]*margin-top:\s*20px[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*"[^>]*>[\s\S]*?NPCDCS[\s\S]*?<\/div>/g,
                image: 'national_program_diagram.png'
            },
            // Control strategies
            {
                pattern: /<div[^>]*>[\s\S]*?Control Strategies[\s\S]*?<\/div>\s*<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">[\s\S]*?<\/div>/g,
                image: 'control_strategies_diagram.png'
            }
        ];

        for (const { pattern, image } of placeholderPatterns) {
            if (image in imageMappings) {
                const replacement = `<div style="margin-bottom: 20px;"><img src="${imageMappings[image]}" alt="${image.replace(/_/g, ' ').replace('.png', '').replace(/\b\w/g, l => l.toUpperCase())}" style="width: 100%; border-radius: 8px;"></div>`;
                content = content.replace(pattern, replacement);
                replacementsMade++;
                console.log(`🔄 Replaced placeholder: ${image}`);
            }
        }

        return { content, replacementsMade };
    }

    replaceImageTags(content, imageMappings) {
        let replacementsMade = 0;

        // Replace standard img tags
        const imgTagPattern = /<img[^>]*src="([^"]*\.png)"[^>]*>/gi;
        content = content.replace(imgTagPattern, (match, src) => {
            const filename = path.basename(src);
            if (filename in imageMappings) {
                const newMatch = match.replace(src, imageMappings[filename]);
                replacementsMade++;
                console.log(`🖼️  Embedded: ${filename}`);
                return newMatch;
            }
            return match;
        });

        // Replace CSS background-image
        const cssBgPattern = /background-image:\s*url\("([^"]*\.png)"\)/gi;
        content = content.replace(cssBgPattern, (match, src) => {
            const filename = path.basename(src);
            if (filename in imageMappings) {
                const newMatch = match.replace(src, imageMappings[filename]);
                replacementsMade++;
                console.log(`🎨 Embedded background: ${filename}`);
                return newMatch;
            }
            return match;
        });

        return { content, replacementsMade };
    }

    async createEmbeddedHtml(outputFile = null) {
        try {
            // Check if HTML file exists
            if (!fs.existsSync(this.htmlFile)) {
                console.error(`❌ HTML file not found: ${this.htmlFile}`);
                return false;
            }

            const originalSize = await this.getFileSize(this.htmlFile);
            console.log(`📄 Original HTML size: ${originalSize} MB`);

            // Read HTML content
            let htmlContent = await readFile(this.htmlFile, 'utf8');

            // Get image mappings
            const imageMappings = await this.getImageMappings();

            if (Object.keys(imageMappings).length === 0) {
                console.log('⚠️  No images found to embed');
                return false;
            }

            // Process the content
            const { content: contentWithImages, replacementsMade: imgReplacements } = this.replaceImageTags(htmlContent, imageMappings);
            const { content: finalContent, replacementsMade: placeholderReplacements } = this.replacePlaceholderDivs(contentWithImages, imageMappings);

            const totalReplacements = imgReplacements + placeholderReplacements;

            // Determine output file
            if (!outputFile) {
                const parsedPath = path.parse(this.htmlFile);
                outputFile = path.join(parsedPath.dir, `${parsedPath.name}_embedded${parsedPath.ext}`);
            }

            // Write the modified HTML
            await writeFile(outputFile, finalContent, 'utf8');

            // Get final file size
            const finalSize = await this.getFileSize(outputFile);

            // Results
            console.log(`\n${'='.repeat(60)}`);
            console.log('🎉 EMBEDDING COMPLETE!');
            console.log(`✅ Images processed: ${Object.keys(imageMappings).length}`);
            console.log(`✅ Replacements made: ${totalReplacements}`);
            console.log(`✅ Output file: ${outputFile}`);
            console.log(`📁 Size: ${originalSize} MB → ${finalSize} MB`);
            console.log('✅ Ready for offline use and Google Sites upload!');
            console.log(`${'='.repeat(60)}\n`);

            return true;

        } catch (error) {
            console.error('❌ Error processing HTML file:', error.message);
            return false;
        }
    }

    // Watch mode for automatic rebuilding
    async watch() {
        const chokidar = require('chokidar');

        console.log('👀 Watch mode enabled. Monitoring for changes...');
        console.log('Press Ctrl+C to stop watching\n');

        const watcher = chokidar.watch([this.htmlFile, this.imagesDir], {
            ignored: /(^|[\/\\])\../,
            persistent: true
        });

        watcher.on('change', async (changedPath) => {
            const timestamp = new Date().toLocaleTimeString();
            console.log(`\n[${timestamp}] File changed: ${changedPath}`);
            await this.createEmbeddedHtml();
        });

        // Keep the process running
        process.on('SIGINT', () => {
            console.log('\n👋 Stopping watch mode...');
            watcher.close();
            process.exit(0);
        });
    }
}

// CLI Interface
async function main() {
    const args = process.argv.slice(2);
    const watch = args.includes('--watch');
    const compress = args.includes('--compress');

    console.log('🖼️  DIABETES HTML IMAGE EMBEDDER (Node.js)');
    console.log('=' .repeat(50));

    if (watch) {
        console.log('📋 Mode: Watch for changes');
    } else if (compress) {
        console.log('📋 Mode: Embed with compression');
    } else {
        console.log('📋 Mode: One-time embedding');
    }

    const embedder = new ImageEmbedder();

    if (watch) {
        await embedder.watch();
    } else {
        await embedder.createEmbeddedHtml();
    }
}

// Error handling
process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
    process.exit(1);
});

if (require.main === module) {
    main().catch(error => {
        console.error('❌ Fatal error:', error.message);
        process.exit(1);
    });
}

module.exports = ImageEmbedder;

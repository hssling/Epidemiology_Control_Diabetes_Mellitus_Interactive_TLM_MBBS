/**
 * Diabetes HTML Image Embedder - Gulp Build Tool
 * Automated workflow for embedding images and building HTML
 */

const gulp = require('gulp');
const fs = require('fs');
const path = require('path');
const through = require('through2');

function embedImages() {
    return gulp.src('interactive/diabetes_interactive_tlm.html')
        .pipe(embedImageData())
        .pipe(gulp.dest('dist'));
}

function embedImageData() {
    return through.obj(function(file, enc, cb) {
        if (file.isNull()) {
            return cb(null, file);
        }

        if (file.isStream()) {
            return cb(new Error('Streaming not supported'));
        }

        let content = file.contents.toString();
        const visualizationsDir = path.resolve('visualizations');

        // Get list of PNG files
        let imageCount = 0;
        let replacementsMade = 0;

        try {
            const files = fs.readdirSync(visualizationsDir);
            const pngFiles = files.filter(f => f.endsWith('.png'));

            console.log(`📁 Processing ${pngFiles.length} images from ${visualizationsDir}`);

            for (const pngFile of pngFiles) {
                const fullPath = path.join(visualizationsDir, pngFile);
                const imageBuffer = fs.readFileSync(fullPath);
                const base64Data = imageBuffer.toString('base64');
                const dataUri = `data:image/png;base64,${base64Data}`;

                // Replace various patterns
                const patterns = [
                    // Standard img src
                    new RegExp(`src="([^"]*${pngFile})"`, 'g'),
                    // CSS background-image
                    new RegExp(`background-image:\\s*url\\("([^"]*${pngFile})"\\)`, 'g'),
                ];

                for (const pattern of patterns) {
                    content = content.replace(pattern, (match, p1) => {
                        replacementsMade++;
                        return match.replace(p1, dataUri);
                    });
                }

                imageCount++;
                console.log(`✅ Embedded: ${pngFile}`);
            }

            // Handle placeholder divs
            const placeholderReplacements = replacePlaceholderDivs(content, pngFiles);
            content = placeholderReplacements.content;
            replacementsMade += placeholderReplacements.replacements;

            file.contents = Buffer.from(content);

            const originalSize = file.contents.length / (1024 * 1024);
            const newSize = Buffer.from(content).length / (1024 * 1024);

            console.log(`\n🎉 Gulp build complete!`);
            console.log(`📊 Statistics:`);
            console.log(`   Images: ${imageCount}`);
            console.log(`   Replacements: ${replacementsMade}`);
            console.log(`   Size: ${originalSize.toFixed(2)} MB → ${newSize.toFixed(2)} MB`);
            console.log(`   Output: ${path.relative(process.cwd(), file.path.replace(path.extname(file.path), '_embedded' + path.extname(file.path)))}\n`);

        } catch (error) {
            console.error('❌ Error processing file:', error.message);
        }

        cb(null, file);
    });
}

function replacePlaceholderDivs(content, pngFiles) {
    let replacements = 0;

    // Define placeholder patterns
    const placeholders = [
        {
            pattern: /<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#9b59b6[^"]*">[\s\S]*?<\/div>/g,
            image: 'pathophysiology_diagram.png'
        },
        {
            pattern: /<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">(?![\s\S]*?(?:Prevention|NPCDCS|Control))[\s\S]*?<\/div>/g,
            image: 'epidemiology_chart.png'
        },
        {
            pattern: /<div[^>]*>[\s\S]*?Management Tab Diagram[\s\S]*?<\/div>\s*<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">[\s\S]*?<\/div>/g,
            image: 'treatment_algorithm.png'
        },
        {
            pattern: /<div[^>]*style="[^"]*margin-top:\s*20px[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*"[^>]*>[\s\S]*?Prevention Flowchart[\s\S]*?<\/div>/g,
            image: 'prevention_flowchart.png'
        },
        {
            pattern: /<div[^>]*style="[^"]*margin-top:\s*20px[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*"[^>]*>[\s\S]*?NPCDCS[\s\S]*?<\/div>/g,
            image: 'national_program_diagram.png'
        },
        {
            pattern: /<div[^>]*>[\s\S]*?Control Strategies[\s\S]*?<\/div>\s*<div[^>]*style="[^"]*background:\s*#f8f9fa[^"]*padding:\s*20px[^"]*border-radius:\s*8px[^"]*text-align:\s*center[^"]*border:\s*2px\s*dashed\s*#3498db[^"]*">[\s\S]*?<\/div>/g,
            image: 'control_strategies_diagram.png'
        }
    ];

    // Create image data URIs map
    const imageUris = {};
    for (const pngFile of pngFiles) {
        const fullPath = path.join('visualizations', pngFile);
        const imageBuffer = fs.readFileSync(fullPath);
        const base64Data = imageBuffer.toString('base64');
        imageUris[pngFile] = `data:image/png;base64,${base64Data}`;
    }

    // Replace placeholders
    for (const { pattern, image } of placeholders) {
        if (image in imageUris) {
            const replacement = `<div style="margin-bottom: 20px;"><img src="${imageUris[image]}" alt="${image.replace(/_/g, ' ').replace('.png', '').replace(/\b\w/g, l => l.toUpperCase())}" style="width: 100%; border-radius: 8px;"></div>`;
            content = content.replace(pattern, replacement);
            replacements++;
            console.log(`🔄 Gulp replaced placeholder: ${image}`);
        }
    }

    return { content, replacements };
}

function renameOutput() {
    return through.obj(function(file, enc, cb) {
        const parsed = path.parse(file.path);
        file.path = path.join(parsed.dir, `${parsed.name}_embedded${parsed.ext}`);
        cb(null, file);
    });
}

function clean() {
    return require('del')(['dist/**/*']);
}

function watch() {
    gulp.watch(['interactive/*.html', 'visualizations/*.png'], embedImages);
}

function build() {
    return gulp.series(clean, embedImages);
}

function defaultTask() {
    console.log('\n🖼️  Diabetes HTML Image Embedder - Gulp Version');
    console.log('=' .repeat(50));
    console.log('Available tasks:');
    console.log('  gulp build     - Build embedded HTML to dist/');
    console.log('  gulp embed     - Embed images into HTML');
    console.log('  gulp watch     - Watch for changes and auto-build');
    console.log('  gulp clean     - Clean dist/ directory');
    console.log('  gulp           - Show this help message\n');
}

// Exports
exports.embed = embedImages;
exports.build = build();
exports.watch = gulp.series(build(), watch);
exports.clean = clean;
exports.default = defaultTask;

// Also support old-style gulp tasks for compatibility
gulp.task('embed', embedImages);
gulp.task('build', build());
gulp.task('watch', gulp.series(build(), watch));
gulp.task('clean', clean);
gulp.task('default', defaultTask);

    // --- REPLACEMENT ENGINE (Block-Centering Logic) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');

        return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));

                // 1. Split into lines and find the longest line in the section
                let lines = (replacement || '').split(/\r?\n/);
                const maxLength = Math.max(...lines.map(l => l.length));

                // 2. Block Centering Calculation
                // Standard PPT slide width at ~24-28pt font is roughly 60-70 characters wide.
                const CANVAS_WIDTH = 65; 
                const padCount = Math.max(0, Math.floor((CANVAS_WIDTH - maxLength) / 2));
                const blockPadding = " ".repeat(padCount);

                // 3. Process lines: Apply uniform padding to the left of EVERY line
                const processedLines = lines.map(l => this.escXml(blockPadding + l));

                phRegex.lastIndex = 0;
                let injected = '';
                processedLines.forEach((line, idx) => {
                    // xml:space="preserve" is CRITICAL to stop PPT from deleting your spaces
                    if (idx > 0) injected += `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`;
                    injected += line;
                });

                // Auto-shrink for long sections
                if (placeholder === '[Lyrics and Chords]' && processedLines.length > 10) {
                    const szMatch = style.match(/sz=\"(\d+)\"/);
                    if (szMatch) {
                        const scale = Math.max(0.6, 1 - (processedLines.length - 10) * 0.05);
                        style = style.replace(/sz=\"\d+\"/, `sz="${Math.floor(parseInt(szMatch[1]) * scale)}"`);
                    }
                }

                // 4. Inject and FORCE Left Alignment
                // We must use Left Align ('l') because our padding handles the centering.
                let result = shapeXml.replace(phRegex, () => {
                    return `</a:t></a:r><a:r>${style}<a:t xml:space="preserve">${injected}</a:t></a:r><a:r>${style}<a:t xml:space="preserve">`;
                })
                .replace(/algn="ctr"/g, 'algn="l"') // Change center to left
                .replace(/<a:t xml:space="preserve"><\/a:t>/g, '')
                .replace(/<a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r>/g, '');

                // Standard Autofit Logic
                if (result.includes('<a:noAutofit/>')) {
                    result = result.replace('<a:noAutofit/>', '<a:normAutofit fontScale="75000" lnSpcReduction="15000"/>');
                } else if (!result.includes('Autofit')) {
                    result = result.replace('</a:bodyPr>', '<a:normAutofit fontScale="75000" lnSpcReduction="15000"/></a:bodyPr>');
                }
                return result;
            }
            return shapeXml;
        });
    },
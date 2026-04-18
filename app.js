    // --- REPLACEMENT ENGINE (Updated for Centered Alignment) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');

        return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));

                // --- NEW ALIGNMENT LOGIC ---
                let lines = (replacement || '').split(/\r?\n/);
                
                // 1. Find the longest line to determine the block width
                const maxLength = Math.max(...lines.map(l => l.length));
                
                // 2. Define a "Canvas Width" (Characters per line). 
                // 60 is a safe average for standard PPT slides.
                const CANVAS_WIDTH = 60; 
                const paddingAmount = Math.max(0, Math.floor((CANVAS_WIDTH - maxLength) / 2));
                const leadingSpaces = " ".repeat(paddingAmount);

                // 3. Apply padding to every line and escape XML
                const processedLines = lines.map(l => this.escXml(leadingSpaces + l));
                
                phRegex.lastIndex = 0;
                let injected = '';
                processedLines.forEach((line, idx) => {
                    // Note: We use xml:space="preserve" to ensure PPT doesn't trim our padding
                    if (idx > 0) injected += `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`;
                    injected += line;
                });

                // Manual shrink for long lyrics
                if (placeholder === '[Lyrics and Chords]' && processedLines.length > 10) {
                    const szMatch = style.match(/sz=\"(\d+)\"/);
                    if (szMatch) {
                        const scale = Math.max(0.6, 1 - (processedLines.length - 10) * 0.05);
                        style = style.replace(/sz=\"\d+\"/, `sz="${Math.floor(parseInt(szMatch[1]) * scale)}"`);
                    }
                }

                let result = shapeXml.replace(phRegex, () => {
                    return `</a:t></a:r><a:r>${style}<a:t xml:space="preserve">${injected}</a:t></a:r><a:r>${style}<a:t xml:space="preserve">`;
                }).replace(/<a:t xml:space="preserve"><\/a:t>/g, '').replace(/<a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r>/g, '');

                // Force Left Alignment in the XML so our manual space-padding works
                // If the template was centered, our spaces would double-center and break it.
                result = result.replace(/algn="ctr"/g, 'algn="l"');

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
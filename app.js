/* LyricSlide Pro - Version 3.1.0 (MS Compliance Engine) */

const App = {
    // ... (Keep elements, musical, chordRegex, originalSlides as they were)
    version: "Version 3.1.0-Compliance",
    elements: {
        songTitle: document.getElementById('songTitle'),
        lyricsInput: document.getElementById('lyricsInput'),
        copyrightInfo: document.getElementById('copyrightInfo'),
        generateBtn: document.getElementById('generateBtn'),
        transFileInput: document.getElementById('transFileInput'),
        transposeBtn: document.getElementById('transposeBtn'),
        semitoneDisplay: document.getElementById('semitoneDisplay'),
        loadingOverlay: document.getElementById('loadingOverlay'),
        loadingText: document.getElementById('loadingText')
    },
    musical: {
        keys: ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'],
        flats: ['C', 'Db', 'D', 'Eb', 'E', 'F', 'Gb', 'G', 'Ab', 'A', 'Bb', 'B'],
        preferred: ['C', 'C#', 'D', 'Eb', 'E', 'F', 'F#', 'G', 'Ab', 'A', 'Bb', 'B']
    },
    chordRegex: /(?:\[)?\b([A-G][b#]?)((?:m|maj|dim|aug|sus|add|[245679]|11|13|[\(\)])*)(\/[A-G][b#]?)?\b(?:\])?/g,
    originalSlides: [],   
    selectedTemplateFile: null, 

    init() {
        if (this.elements.generateBtn) this.elements.generateBtn.addEventListener('click', () => this.generate());
        if (this.elements.transposeBtn) this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        const alignSelect = document.getElementById('alignmentSelect');
        if (alignSelect) alignSelect.addEventListener('change', () => { if (this.originalSlides.length > 0) this.updatePreview(0); });
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        const versionEl = document.getElementById('appVersion');
        if (versionEl) versionEl.textContent = this.version;
    },

    // --- NEW: COMPLIANCE HELPERS ---

    async updateContentTypes(zip, slideNames, notesNames) {
        const path = '[Content_Types].xml';
        let content = await zip.file(path).async('string');
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(content, 'application/xml');
        const types = xmlDoc.documentElement;

        const addOverride = (part, type) => {
            if (!content.includes(`PartName="${part}"`)) {
                const el = xmlDoc.createElement('Override');
                el.setAttribute('PartName', part);
                el.setAttribute('ContentType', type);
                types.appendChild(el);
            }
        };

        slideNames.forEach(name => addOverride(`/ppt/slides/${name}`, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'));
        notesNames.forEach(name => addOverride(`/ppt/notesSlides/${name}`, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'));

        zip.file(path, new XMLSerializer().serializeToString(xmlDoc));
    },

    // --- CORE GENERATION (Now with Compliance) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Select a template and input lyrics.');

        try {
            this.showLoading('Generating Compliant PPTX...');
            const zip = await JSZip.loadAsync(file);
            
            // 1. Get Template Structure
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid]; // e.g. slides/slide1.xml
            const templateFileName = templateRelPath.split('/').pop();

            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templateFileName}.rels`).async('string');
            
            // Extract Layout ID (Critical for MS compliance)
            const layoutMatch = templateRelsXml.match(/Type="[^"]*?slideLayout"[^>]*?Target="([^"]+)"/);
            const layoutTarget = layoutMatch ? layoutMatch[1] : '../slideLayouts/slideLayout1.xml';

            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            
            const generatedSlides = [];
            const slideFileNames = [];
            const notesFileNames = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionText = sections[i].trim();
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);

                const sName = `slide_gen_${i + 1}.xml`;
                const nName = `notes_gen_${i + 1}.xml`;
                
                zip.file(`ppt/slides/${sName}`, slideXml);
                slideFileNames.push(sName);

                // Create Slide-to-Layout relationship file (The Missing Link)
                let slideRelContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
                slideRelContent += `<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutTarget}"/>`;
                
                if (templateNotesXml) {
                    slideRelContent += `<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/${nName}"/>`;
                    
                    // Generate Notes Content
                    const styleMatch = templateNotesXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/);
                    const notesStyle = styleMatch ? styleMatch[0] : '<a:rPr lang="en-US" sz="1600"/>';
                    const noteLines = sectionText.split(/\n/).map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formattedNotes = this.escXml(noteLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r>${notesStyle}<a:t xml:space="preserve">`);
                    
                    zip.file(`ppt/notesSlides/${nName}`, templateNotesXml.replace(new RegExp(this.getPlaceholderRegexStr('[Presenter Note]'), 'gi'), formattedNotes));
                    zip.file(`ppt/notesSlides/_rels/${nName}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${sName}"/></Relationships>`);
                    notesFileNames.push(nName);
                }
                
                slideRelContent += `</Relationships>`;
                zip.file(`ppt/slides/_rels/${sName}.rels`, slideRelContent);

                generatedSlides.push({ id: 5000 + i, rid: `rIdG${i + 1}`, name: sName });
            }

            // Sync Registry and Content Types
            this.syncPresentationRegistry(zip, presXml, presRelsXml, generatedSlides);
            await this.updateContentTypes(zip, slideFileNames, notesFileNames);

            const blob = await zip.generateAsync({ type: 'blob' });
            saveAs(blob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { console.error(err); alert(err.message); this.hideLoading(); }
    },

    // --- UPDATED REPLACEMENT LOGIC (Prevents Illegal Tag Nesting) ---
    lockInStyleAndReplace(xml, ph, replacement, align = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(ph), 'gi');
        
        // We match the entire Paragraph <a:p> containing the placeholder
        // instead of just replacing text inside <a:t>.
        return xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (pMatch) => {
            if (!phRegex.test(pMatch)) return pMatch;

            // Extract the style from the first run in this paragraph
            const style = pMatch.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr lang="en-US"/>';
            const alignment = align === 'ctr' ? 'ctr' : 'l';

            if (ph !== '[Lyrics and Chords]') {
                const escaped = this.escXml(replacement);
                return `<a:p><a:pPr algn="${alignment}"/><a:r>${style}<a:t xml:space="preserve">${escaped}</a:t></a:r></a:p>`;
            }

            // Handle Lyrics & Chords Multi-line Injection
            let newParagraphs = "";
            const rawLines = replacement.split('\n');
            for (let i = 0; i < rawLines.length; i++) {
                let line = rawLines[i], next = rawLines[i+1];
                if (this.isChordLine(line) && next && !this.isChordLine(next) && !next.trim().startsWith('[')) {
                    const max = Math.max(line.length, next.length);
                    if (align === 'ctr') {
                        newParagraphs += this.makeGhostAlignmentLine(line.padEnd(max,' '), next.padEnd(max,' '), style, 'ctr');
                        newParagraphs += this.makePptLine(next.padEnd(max,' '), style, 'ctr');
                    } else {
                        newParagraphs += this.makePptLine(line, this.getChordStyle(style), 'l');
                        newParagraphs += this.makePptLine(next, style, 'l');
                    }
                    i++;
                } else {
                    const text = line.trim();
                    if (!text) {
                        newParagraphs += `<a:p><a:pPr algn="${alignment}"/></a:p>`;
                    } else {
                        const isTag = text.startsWith('[') && text.endsWith(']');
                        let curStyle = isTag ? style.replace(/sz="\d+"/, 'sz="2000"') : style;
                        newParagraphs += this.makePptLine(text, curStyle, alignment);
                    }
                }
            }
            return newParagraphs;
        });
    },

    // ... (Keep other helpers like shiftNote, isChordLine, getChordStyle, theme as they were)
    // ... (Keep transpose and transposeParagraphs logic)
    
    // Updated Registry Sync to be cleaner
    syncPresentationRegistry(zip, xml, rels, gen) {
        // Update presentation.xml
        const newSldIdLst = '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', xml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, newSldIdLst));

        // Update presentation.xml.rels
        const parser = new DOMParser();
        const doc = parser.parseFromString(rels, 'application/xml');
        const rNode = doc.documentElement;
        
        // Remove old slide relationships
        const oldRels = rNode.getElementsByTagName('Relationship');
        for (let i = oldRels.length - 1; i >= 0; i--) {
            if (oldRels[i].getAttribute('Type').endsWith('slide')) rNode.removeChild(oldRels[i]);
        }

        // Add new ones
        gen.forEach(s => {
            const e = doc.createElement('Relationship');
            e.setAttribute('Id', s.rid);
            e.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            e.setAttribute('Target', `slides/${s.name}`);
            rNode.appendChild(e);
        });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(doc));
    },

    // (Helper stubs to ensure full functionality)
    makePptLine(text, style, align) {
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr><a:r>${style}<a:t xml:space="preserve">${this.escXml(text)}</a:t></a:r></a:p>`;
    },

    makeGhostAlignmentLine(chord, lyric, style, align) {
        let ghost = style.replace('<a:rPr', '<a:rPr><a:noFill/>').replace(/<a:solidFill>.*?<\/a:solidFill>/g, '');
        let xml = "";
        for (let i = 0; i < chord.length; i++) {
            xml += (chord[i] === ' ') ? `<a:r>${ghost}<a:t xml:space="preserve">${this.escXml(lyric[i] || ' ')}</a:t></a:r>` 
                                     : `<a:r>${this.getChordStyle(style)}<a:t xml:space="preserve">${this.escXml(chord[i])}</a:t></a:r>`;
        }
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr>${xml}</a:p>`;
    },

    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(slideRelsXml) { const m = slideRelsXml?.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    isChordLine(lineStr) {
        if (!lineStr || typeof lineStr !== 'string') return false;
        const trimmed = lineStr.trim();
        if (trimmed === '' || /^(A|I|The|And|Then|They|We|He|She)\s+[a-zA-Z]{2,}/i.test(trimmed)) return false;
        const chords = trimmed.match(this.chordRegex) || [];
        return chords.length > 0;
    },
    getChordStyle(lyricStyle) {
        let s = lyricStyle.includes('sz=') ? lyricStyle.replace(/sz="\d+"/, 'sz="1800"') : lyricStyle.replace('<a:rPr', '<a:rPr sz="1800"');
        const greyFill = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>';
        return s.includes('<a:solidFill>') ? s.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, greyFill) : s.replace('</a:rPr>', greyFill + '</a:rPr>');
    },
    showLoading(text) { this.elements.loadingText.textContent = text; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    
    // Theme and Gallery placeholders (keep from original)
    theme: {
        defaults: {'--primary-color': '#334155', '--bg-start': '#f8fafc', '--bg-end': '#f8fafc', '--text-main': '#1e293b', '--card-accent': '#e2e8f0', '--preview-card-bg': '#ffffff', '--preview-chord-color': '#334155', '--preview-lyrics-color': '#1e293b'},
        init() { /* ... implementation from original ... */ },
        setVariable(name, val) { document.documentElement.style.setProperty(name, val); }
    },
    async loadDefaultTemplates() { /* ... implementation from original ... */ }
};

App.init();
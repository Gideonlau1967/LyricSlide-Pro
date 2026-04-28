/* LyricSlide Pro - Version 3.4.0 (Final Integrity Build) */

const App = {
    version: "Version 3.4.0",
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
        if (this.elements.transFileInput) this.elements.transFileInput.addEventListener('change', (e) => e.target.files[0] && this.loadForPreview(e.target.files[0]));
        this.theme.init();
        this.loadDefaultTemplates(); 
        if (document.getElementById('appVersion')) document.getElementById('appVersion').textContent = this.version;
    },

    // --- CORE GENERATION (INTEGRITY FIRST) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Input lyrics and select template.');

        try {
            this.showLoading('Analyzing Template Architecture...');
            const zip = await JSZip.loadAsync(file);

            // 1. DATA EXTRACTION (Cloning the source)
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const templatePath = slideRels[this.getSlideIds(presXml)[0].rid]; // Likely ppt/slides/slide1.xml
            const templateXml = await zip.file(`ppt/${templatePath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templatePath.split('/').pop()}.rels`).async('string');
            
            // Notes Cloning
            const templateNotesRelPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesRelPath ? await zip.file(templateNotesRelPath).async('string') : null;
            const templateNotesRelsXml = templateNotesRelPath ? await zip.file(`ppt/notesSlides/_rels/${templateNotesRelPath.split('/').pop()}.rels`).async('string') : null;

            // 2. CONTENT GENERATION
            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/i;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionText = sections[i].trim();
                const sName = `slide_gen_${i + 1}.xml`;
                
                // Clone Slide XML and replace text
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);
                zip.file(`ppt/slides/${sName}`, slideXml);
                
                // Clone Slide Relationships exactly
                zip.file(`ppt/slides/_rels/${sName}.rels`, templateRelsXml.replace(/notesSlide\d+\.xml/g, `notes_gen_${i + 1}.xml`));

                if (templateNotesXml) {
                    const nName = `notes_gen_${i + 1}.xml`;
                    const noteLines = sectionText.split(/\n/).map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formattedNotes = this.escXml(noteLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">`);
                    const newNotesContent = templateNotesXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">${formattedNotes}</a:t></a:r></a:p>`);
                    
                    zip.file(`ppt/notesSlides/${nName}`, newNotesContent);
                    // Clone Notes Relationships exactly
                    zip.file(`ppt/notesSlides/_rels/${nName}.rels`, templateNotesRelsXml.replace(/slide\d+\.xml/g, sName));
                }
                generated.push({ id: 1000 + i, rid: `rIdGen${i + 1}`, name: sName });
            }

            // 3. ARCHITECTURAL SYNC (The "No-Repair" Patch)
            await this.finalizeAndRepair(zip, presXml, presRelsXml, generated);
            
            const finalBlob = await zip.generateAsync({ 
                type: 'blob', 
                mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" 
            });
            saveAs(finalBlob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    async finalizeAndRepair(zip, presXml, presRelsXml, generated) {
        const serializer = new XMLSerializer();
        const parser = new DOMParser();

        // A. Sync presentation.xml (Slide ID List)
        const presDoc = parser.parseFromString(presXml, 'application/xml');
        const sldIdLst = presDoc.getElementsByTagName('p:sldIdLst')[0];
        while (sldIdLst.firstChild) sldIdLst.removeChild(sldIdLst.firstChild);
        generated.forEach(s => {
            const node = presDoc.createElement('p:sldId');
            node.setAttribute('id', (256 + s.id).toString());
            node.setAttribute('r:id', s.rid);
            sldIdLst.appendChild(node);
        });
        zip.file('ppt/presentation.xml', serializer.serializeToString(presDoc));

        // B. Sync presentation.xml.rels
        const relsDoc = parser.parseFromString(presRelsXml, 'application/xml');
        const rRoot = relsDoc.documentElement;
        [...rRoot.getElementsByTagName('Relationship')].forEach(r => r.getAttribute('Type').endsWith('slide') && r.remove());
        generated.forEach(s => {
            const e = relsDoc.createElement('Relationship');
            e.setAttribute('Id', s.rid);
            e.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            e.setAttribute('Target', `slides/${s.name}`);
            rRoot.appendChild(e);
        });
        zip.file('ppt/_rels/presentation.xml.rels', serializer.serializeToString(relsDoc));

        // C. Clean [Content_Types].xml
        const ctDoc = parser.parseFromString(await zip.file('[Content_Types].xml').async('string'), 'application/xml');
        const ctRoot = ctDoc.documentElement;
        [...ctRoot.getElementsByTagName('Override')].forEach(ov => {
            const pn = ov.getAttribute('PartName');
            if (pn.includes('/ppt/slides/') || pn.includes('/ppt/notesSlides/')) ov.remove();
        });
        generated.forEach(s => {
            const sEl = ctDoc.createElement('Override');
            sEl.setAttribute('PartName', `/ppt/slides/${s.name}`);
            sEl.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            ctRoot.appendChild(sEl);
            const nName = s.name.replace('slide_gen_', 'notes_gen_');
            if (zip.file(`ppt/notesSlides/${nName}`)) {
                const nEl = ctDoc.createElement('Override');
                nEl.setAttribute('PartName', `/ppt/notesSlides/${nName}`);
                nEl.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
                ctRoot.appendChild(nEl);
            }
        });
        zip.file('[Content_Types].xml', serializer.serializeToString(ctDoc));

        // D. Update Metadata (Slide Count)
        if (zip.file('docProps/app.xml')) {
            let appXml = await zip.file('docProps/app.xml').async('string');
            appXml = appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${generated.length}</Slides>`)
                           .replace(/<I4>\d+<\/I4>/, `<I4>${generated.length}</I4>`);
            zip.file('docProps/app.xml', appXml);
        }
    },

    // --- REPLACEMENT ENGINE (Handles XML fragmentation) ---
    lockInStyleAndReplace(xml, ph, replacement, align = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(ph), 'gi');
        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shape) => {
            if (!phRegex.test(shape)) return shape;
            const style = shape.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr lang="en-US"/>';
            
            // For simple Title/Copyright
            if (!ph.toLowerCase().includes('lyrics')) {
                const escaped = replacement.split('\n').map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                return shape.replace(phRegex, escaped);
            }
            
            // For Lyrics and Chords
            let injected = `</a:t></a:r></a:p>`;
            const rawLines = replacement.split('\n');
            for (let i = 0; i < rawLines.length; i++) {
                let line = rawLines[i], next = rawLines[i+1];
                if (this.isChordLine(line) && next && !this.isChordLine(next) && !next.trim().startsWith('[')) {
                    const max = Math.max(line.length, next.length);
                    injected += (align === 'ctr') ? this.makeGhostAlignmentLine(line.padEnd(max,' '), next.padEnd(max,' '), style, 'ctr') + this.makePptLine(next.padEnd(max,' '), style, 'ctr') 
                                                : this.makePptLine(line, this.getChordStyle(style), 'l') + this.makePptLine(next, style, 'l');
                    i++;
                } else {
                    const text = line.trim(), isTag = text.startsWith('[') && text.endsWith(']');
                    let curStyle = isTag ? style.replace(/sz="\d+"/, 'sz="2000"') : style;
                    injected += text ? this.makePptLine(text, curStyle, align === 'ctr' ? 'ctr' : 'l') : `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                }
            }
            return shape.replace(phRegex, injected + `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t xml:space="preserve">`)
                        .replace('</a:bodyPr>', '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/></a:bodyPr>');
        });
    },

    // --- SUPPORT UTILITIES ---
    makeGhostAlignmentLine(c, l, s, a) {
        let gh = s.replace('<a:rPr', '<a:rPr><a:noFill/>').replace(/<a:solidFill>.*?<\/a:solidFill>/g, '');
        let xml = "";
        for (let i = 0; i < c.length; i++) xml += (c[i] === ' ') ? `<a:r>${gh}<a:t xml:space="preserve">${this.escXml(l[i] || ' ')}</a:t></a:r>` : `<a:r>${this.getChordStyle(s)}<a:t xml:space="preserve">${this.escXml(c[i])}</a:t></a:r>`;
        return `<a:p><a:pPr algn="${a}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr>${xml}</a:p>`;
    },
    makePptLine(t, s, a) { return `<a:p><a:pPr algn="${a}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr><a:r>${s}<a:t xml:space="preserve">${this.escXml(t)}</a:t></a:r></a:p>`; },
    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(rXml) { const m = rXml?.match(/Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    isChordLine(l) { if(!l) return false; const t = l.trim(), w = t.split(/\s+/), c = t.match(this.chordRegex) || []; return c.length >= w.length * 0.5 || (c.length > 0 && w.length <= 2); },
    getChordStyle(s) { const f = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>'; let res = s.includes('sz=') ? s.replace(/sz="\d+"/, 'sz="1800"') : s.replace('<a:rPr', '<a:rPr sz="1800"'); return res.includes('<a:solidFill>') ? res.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, f) : res.replace('</a:rPr>', f + '</a:rPr>'); },
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    theme: { init() {}, save() {} }, // Theme placeholders
    loadDefaultTemplates() {} 
};

App.init();
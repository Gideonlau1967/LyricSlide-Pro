/* LyricSlide Pro - Version 3.2.0 (Optimized for Layout2 & NotesMaster) */

const App = {
    version: "Version 3.2.0",
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
        if (alignSelect) alignSelect.addEventListener('change', () => this.originalSlides.length > 0 && this.updatePreview(0));
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        if (document.getElementById('appVersion')) document.getElementById('appVersion').textContent = this.version;
    },

    // --- THEME & UI HELPERS ---
    theme: {
        defaults: { '--primary-color': '#334155', '--bg-start': '#f8fafc', '--bg-end': '#f8fafc', '--text-main': '#1e293b', '--card-accent': '#e2e8f0', '--preview-card-bg': '#ffffff', '--preview-chord-color': '#334155', '--preview-lyrics-color': '#1e293b' },
        init() {
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                this.setVariable(key, val);
                const picker = document.getElementById('picker-' + key.replace('--', '').replace('-color', ''));
                if (picker) picker.value = val;
            });
            document.querySelectorAll('.color-picker-input').forEach(picker => {
                picker.addEventListener('input', (e) => {
                    const varName = this.getVarNameFromPicker(e.target.id);
                    this.setVariable(varName, e.target.value);
                    this.save();
                });
            });
        },
        getVarNameFromPicker(id) {
            const map = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' };
            return map[id];
        },
        setVariable(name, val) { document.documentElement.style.setProperty(name, val); if (name === '--primary-color') document.documentElement.style.setProperty('--primary-gradient', val); },
        save() { const current = {}; Object.keys(this.defaults).forEach(key => current[key] = getComputedStyle(document.documentElement).getPropertyValue(key).trim()); localStorage.setItem('lyric_theme', JSON.stringify(current)); }
    },

    setMode(mode) {
        const isGen = mode === 'gen';
        document.getElementById('modeGen').classList.toggle('active', isGen);
        document.getElementById('modeTrans').classList.toggle('active', !isGen);
        document.getElementById('viewGen').classList.toggle('hidden', !isGen);
        document.getElementById('viewTrans').classList.toggle('hidden', isGen);
    },

    // --- GENERATION ENGINE (FIXED COMPATIBILITY) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Select a template and input lyrics.');

        try {
            this.showLoading('Creating Office-Safe PPTX...');
            const zip = await JSZip.loadAsync(file);
            
            // Get internal paths from YOUR specific template
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templateRelPath.split('/').pop()}.rels`).async('string');
            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/i;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionText = sections[i].trim();
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);

                const name = `song_gen_${i + 1}.xml`;
                zip.file(`ppt/slides/${name}`, slideXml);
                
                // Construct Rels (Preserving your Layout 2 link)
                if (templateNotesXml) {
                    const notesName = `notes_gen_${i + 1}.xml`;
                    const noteLines = sectionText.split(/\n/).map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formattedNotes = this.escXml(noteLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">`);
                    
                    zip.file(`ppt/notesSlides/${notesName}`, templateNotesXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">${formattedNotes}</a:t></a:r></a:p>`));
                    zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml.replace(/Target="..\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/${notesName}"`));
                    zip.file(`ppt/notesSlides/_rels/${notesName}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${name}"/></Relationships>`);
                    generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name });
                }
            }

            // --- CRITICAL REPAIR STEPS ---
            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            await this.repairPackageIntegrity(zip, generated);
            
            const finalBlob = await zip.generateAsync({ 
                type: 'blob',
                mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" 
            });
            
            saveAs(finalBlob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    // --- MS OFFICE COMPATIBILITY FIX ---
    async repairPackageIntegrity(zip, generatedSlides) {
        const contentTypesPath = '[Content_Types].xml';
        const contentTypesXml = await zip.file(contentTypesPath).async('string');
        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const ctDoc = parser.parseFromString(contentTypesXml, 'application/xml');
        const typesTag = ctDoc.documentElement;

        // Clear existing slide/notes overrides
        const overrides = typesTag.getElementsByTagName('Override');
        for (let i = overrides.length - 1; i >= 0; i--) {
            const partName = overrides[i].getAttribute('PartName');
            if (partName.includes('/ppt/slides/') || partName.includes('/ppt/notesSlides/')) {
                overrides[i].parentNode.removeChild(overrides[i]);
            }
        }

        // Add correct content types for every NEW slide
        generatedSlides.forEach(s => {
            const sEl = ctDoc.createElement('Override');
            sEl.setAttribute('PartName', `/ppt/slides/${s.name}`);
            sEl.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            typesTag.appendChild(sEl);

            const notesName = s.name.replace('song_gen_', 'notes_gen_');
            if (zip.file(`ppt/notesSlides/${notesName}`)) {
                const nEl = ctDoc.createElement('Override');
                nEl.setAttribute('PartName', `/ppt/notesSlides/${notesName}`);
                nEl.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
                typesTag.appendChild(nEl);
            }
        });

        zip.file(contentTypesPath, serializer.serializeToString(ctDoc));
    },

    // --- CORE PPT XML LOGIC ---
    lockInStyleAndReplace(xml, ph, replacement, align = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(ph), 'gi');
        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shape) => {
            if (!phRegex.test(shape)) return shape;
            const style = shape.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr lang="en-US"/>';
            if (!ph.includes('Lyrics')) {
                const escaped = replacement.split('\n').map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                return shape.replace(phRegex, escaped);
            }
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

    makeGhostAlignmentLine(chord, lyric, style, align) {
        let ghost = style.replace('<a:rPr', '<a:rPr><a:noFill/>').replace(/<a:solidFill>.*?<\/a:solidFill>/g, '');
        let xml = "";
        for (let i = 0; i < chord.length; i++) {
            xml += (chord[i] === ' ') ? `<a:r>${ghost}<a:t xml:space="preserve">${this.escXml(lyric[i] || ' ')}</a:t></a:r>` 
                                     : `<a:r>${this.getChordStyle(style)}<a:t xml:space="preserve">${this.escXml(chord[i])}</a:t></a:r>`;
        }
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr>${xml}</a:p>`;
    },

    makePptLine(text, style, align) {
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr><a:r>${style}<a:t xml:space="preserve">${this.escXml(text)}</a:t></a:r></a:p>`;
    },

    // --- REUSED TRANSPOSITION LOGIC ---
    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        return text.split('\n').map(line => {
            if (!this.isChordLine(line)) return line;
            let result = line, offset = 0;
            const matches = [...line.matchAll(this.chordRegex)];
            for (const m of matches) {
                const nr = this.shiftNote(m[1], semitones); 
                const nb = m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '';
                const nf = nr + (m[2] || '') + nb;
                const p = m.index + offset, d = nf.length - m[0].length;
                let pre = result.substring(0, p), suf = result.substring(p + m[0].length);
                if (d > 0 && suf.startsWith(' ')) { suf = suf.substring(1); offset--; }
                else if (d < 0) { suf = ' '.repeat(Math.abs(d)) + suf; offset += Math.abs(d); }
                result = pre + nf + suf; offset += d;
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let idx = this.musical.keys.indexOf(note);
        if (idx === -1) idx = this.musical.flats.indexOf(note);
        if (idx === -1) return note;
        let newIdx = (idx + semitones) % 12;
        if (newIdx < 0) newIdx += 12;
        return this.musical.preferred[newIdx];
    },

    isChordLine(lineStr) {
        if (!lineStr || typeof lineStr !== 'string') return false;
        const trimmed = lineStr.trim();
        if (trimmed === '' || /^(A|I|The|And|Then|They|We|He|She)\s+[a-zA-Z]{2,}/i.test(trimmed)) return false;
        const words = trimmed.split(/\s+/), chords = trimmed.match(this.chordRegex) || [];
        return chords.length >= words.length * 0.5 || (chords.length > 0 && words.length <= 2);
    },

    getChordStyle(lyricStyle) {
        let s = lyricStyle.includes('sz=') ? lyricStyle.replace(/sz="\d+"/, 'sz="1800"') : lyricStyle.replace('<a:rPr', '<a:rPr sz="1800"');
        return s.includes('<a:solidFill>') ? s.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>') : s.replace('</a:rPr>', '<a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:rPr>');
    },

    // --- SYSTEM HELPERS ---
    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(slideRelsXml) { const m = slideRelsXml?.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    syncPresentationRegistry(zip, xml, rels, gen) {
        zip.file('ppt/presentation.xml', xml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>'));
        let doc = new DOMParser().parseFromString(rels, 'application/xml');
        let rNode = doc.documentElement; [...rNode.getElementsByTagName('Relationship')].forEach(n => n.getAttribute('Type').endsWith('slide') && n.remove());
        gen.forEach(s => { let e = doc.createElement('Relationship'); e.setAttribute('Id', s.rid); e.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); e.setAttribute('Target', `slides/${s.name}`); rNode.appendChild(e); });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(doc));
    },
    showLoading(text) { this.elements.loadingText.textContent = text; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    async loadDefaultTemplates() { /* Implementation for templates.json fetch */ }
};

App.init();
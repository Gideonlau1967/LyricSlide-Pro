/* LyricSlide Pro - Version 3.4 (Auto-Repair Engine) */

const App = {
    version: "Version 3.4",
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
    },

    theme: {
        defaults: {
            '--primary-color': '#334155', '--bg-start': '#f8fafc', '--bg-end': '#f8fafc',
            '--text-main': '#1e293b', '--card-accent': '#e2e8f0', '--preview-card-bg': '#ffffff',
            '--preview-chord-color': '#334155', '--preview-lyrics-color': '#1e293b'
        },
        init() {
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                this.setVariable(key, val);
                const p = document.getElementById('picker-' + key.replace('--', '').replace('-color', ''));
                if (p) p.value = val;
            });
            document.querySelectorAll('.color-picker-input').forEach(p => {
                p.addEventListener('input', (e) => {
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
        setVariable(name, val) {
            document.documentElement.style.setProperty(name, val);
            if (name === '--primary-color') document.documentElement.style.setProperty('--primary-gradient', val);
        },
        save() {
            const current = {};
            Object.keys(this.defaults).forEach(key => { current[key] = getComputedStyle(document.documentElement).getPropertyValue(key).trim(); });
            localStorage.setItem('lyric_theme', JSON.stringify(current));
        }
    },

    // --- STEP 1: CONTENT GENERATION (Original Logic) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Select a template and input lyrics.');

        try {
            this.showLoading('Creating Slides...');
            const zip = await JSZip.loadAsync(file);

            // Fetch template foundations
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const templateRelsPath = `ppt/slides/_rels/${templateRelPath.split('/').pop()}.rels`;
            const templateRelsXml = await zip.file(templateRelsPath).async('string');

            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            let templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;
            let notesMasterRel = "";
            if (templateNotesPath) {
                const nRelsP = `ppt/notesSlides/_rels/${templateNotesPath.split('/').pop()}.rels`;
                if (zip.file(nRelsP)) {
                    notesMasterRel = (await zip.file(nRelsP).async('string')).match(/<Relationship[^>]+Type="[^"]+notesMaster"[^>]+Target="([^"]+)"[^>]*\/>/)?.[0] || "";
                }
            }

            const sections = ("\n" + lyrics).split(/\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/).filter(s => s.trim() !== '');
            const genTracker = [];

            for (let i = 0; i < sections.length; i++) {
                const sName = `slide_gen_${i + 1}.xml`;
                const nName = `notes_gen_${i + 1}.xml`;

                // Build Slide
                let sXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                sXml = this.lockInStyleAndReplace(sXml, '[Copyright Info]', copyright);
                sXml = this.lockInStyleAndReplace(sXml, '[Lyrics and Chords]', sections[i].trim(), userAlign);
                zip.file(`ppt/slides/${sName}`, sXml);

                // Build Slide Rels
                let sRels = templateRelsXml.replace(/Type="[^"]+notesSlide"[^>]+Target="[^"]+"/, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/${nName}"`);
                if (!sRels.includes(nName)) sRels = sRels.replace('</Relationships>', `<Relationship Id="rIdNotes99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/${nName}"/></Relationships>`);
                zip.file(`ppt/slides/_rels/${sName}.rels`, sRels);

                // Build Notes
                if (templateNotesXml) {
                    const style = templateNotesXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr sz="1600"/>';
                    const nLines = sections[i].trim().split('\n').map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formatted = this.escXml(nLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    zip.file(`ppt/notesSlides/${nName}`, templateNotesXml.replace(new RegExp(this.getPlaceholderRegexStr('[Presenter Note]'), 'gi'), formatted));
                    zip.file(`ppt/notesSlides/_rels/${nName}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${sName}"/>${notesMasterRel}</Relationships>`);
                }
                genTracker.push({ id: 256 + i, rid: `rIdG${i + 1}`, sName, nName });
            }

            // --- STEP 2: EXTENSIVE REPAIR ---
            await this.finalizeAndRepair(zip, genTracker);

            this.showLoading('Downloading...');
            const blob = await zip.generateAsync({ type: 'blob' });
            saveAs(blob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { console.error(err); alert("Error: " + err.message); this.hideLoading(); }
    },

    async finalizeAndRepair(zip, slides) {
        this.showLoading('Repairing Structure...');

        // 1. Repair [Content_Types].xml (Wipe old slides, register new ones)
        let ctXml = await zip.file('[Content_Types].xml').async('string');
        ctXml = ctXml.replace(/<Override [^>]+PartName="\/ppt\/slides\/[^>]+>/g, '')
                     .replace(/<Override [^>]+PartName="\/ppt\/notesSlides\/[^>]+>/g, '');
        
        const overrides = slides.map(s => 
            `<Override PartName="/ppt/slides/${s.sName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
            `<Override PartName="/ppt/notesSlides/${s.nName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`
        ).join('');
        zip.file('[Content_Types].xml', ctXml.replace('</Types>', overrides + '</Types>'));

        // 2. Repair presentation.xml (Strict sldIdLst update)
        let presXml = await zip.file('ppt/presentation.xml').async('string');
        const sldIdLst = '<p:sldIdLst>' + slides.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));

        // 3. Repair presentation.xml.rels (Preserving Namespaces & Non-Slide Links)
        let relsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
        const relHeader = relsXml.match(/<Relationships[^>]*>/)[0];
        const nonSlideRels = (relsXml.match(/<Relationship [^>]+>/g) || [])
            .filter(r => !r.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"'));
        
        const newRels = slides.map(s => 
            `<Relationship Id="${s.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${s.sName}"/>`
        );
        zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${relHeader}${nonSlideRels.join('')}${newRels.join('')}</Relationships>`);

        // 4. Repair docProps/app.xml (Update slide count)
        let appXml = await zip.file('docProps/app.xml').async('string');
        appXml = appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${slides.length}</Slides>`);
        zip.file('docProps/app.xml', appXml);
    },

    // --- SHARED UTILS (UNCHANGED) ---
    lockInStyleAndReplace(xml, ph, replacement, align = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(ph), 'gi');
        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shape) => {
            if (!phRegex.test(shape)) return shape;
            const style = shape.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr lang="en-US"/>';
            if (ph !== '[Lyrics and Chords]') {
                const escaped = replacement.split('\n').map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                return shape.replace(phRegex, escaped);
            }
            let injected = `</a:t></a:r></a:p>`;
            const lines = replacement.split('\n');
            for (let i = 0; i < lines.length; i++) {
                let cur = lines[i], next = lines[i+1];
                if (this.isChordLine(cur) && next && !this.isChordLine(next) && !next.trim().startsWith('[')) {
                    const max = Math.max(cur.length, next.length);
                    injected += (align === 'ctr') ? this.makeGhostAlignmentLine(cur.padEnd(max,' '), next.padEnd(max,' '), style, 'ctr') + this.makePptLine(next.padEnd(max,' '), style, 'ctr') : this.makePptLine(cur, this.getChordStyle(style), 'l') + this.makePptLine(next, style, 'l');
                    i++;
                } else {
                    const txt = cur.trim();
                    let s = (txt.startsWith('[') && txt.endsWith(']')) ? style.replace(/sz="\d+"/, 'sz="2000"') : style;
                    injected += txt ? this.makePptLine(txt, s, align === 'ctr' ? 'ctr' : 'l') : `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                }
            }
            return shape.replace(phRegex, injected + `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t xml:space="preserve">`).replace('</a:bodyPr>', '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/></a:bodyPr>');
        });
    },

    makeGhostAlignmentLine(chord, lyric, style, align) {
        let ghost = style.replace('<a:rPr', '<a:rPr><a:noFill/>').replace(/<a:solidFill>.*?<\/a:solidFill>/g, '');
        let xml = "";
        for (let i = 0; i < chord.length; i++) xml += (chord[i] === ' ') ? `<a:r>${ghost}<a:t xml:space="preserve">${this.escXml(lyric[i] || ' ')}</a:t></a:r>` : `<a:r>${this.getChordStyle(style)}<a:t xml:space="preserve">${this.escXml(chord[i])}</a:t></a:r>`;
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr>${xml}</a:p>`;
    },

    makePptLine(text, style, align) { return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr><a:r>${style}<a:t xml:space="preserve">${this.escXml(text)}</a:t></a:r></a:p>`; },
    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(slideRelsXml) { const m = slideRelsXml?.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    isChordLine(lineStr) {
        if (!lineStr || typeof lineStr !== 'string') return false;
        const t = lineStr.trim();
        if (t === '' || /^(A|I|The|And|Then|They|We|He|She)\s+[a-zA-Z]{2,}/i.test(t)) return false;
        const chords = t.match(this.chordRegex) || [];
        const words = t.split(/\s+/).filter(w => w.length > 0);
        return chords.length >= words.length * 0.5 || (chords.length > 0 && words.length <= 2);
    },
    getChordStyle(lyricStyle) {
        let s = lyricStyle.includes('sz=') ? lyricStyle.replace(/sz="\d+"/, 'sz="1800"') : lyricStyle.replace('<a:rPr', '<a:rPr sz="1800"');
        const grey = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>';
        return s.includes('<a:solidFill>') ? s.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, grey) : s.replace('</a:rPr>', grey + '</a:rPr>');
    },

    // --- PREVIEW & TRANSPOSE (UNCHANGED) ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            const entries = names.map(name => ({
                name, getFile: async () => {
                    const r = await fetch(`./${encodeURIComponent(name)}`);
                    const blob = await r.blob();
                    return new File([blob], name, { type: blob.type });
                }
            }));
            this.renderTemplateGallery(entries);
        } catch (e) { gallery.innerHTML = `<div class="text-center py-8 text-slate-400 italic">Template library unavailable.</div>`; }
    },
    renderTemplateGallery(entries) {
        const gallery = document.getElementById('templateGallery');
        gallery.innerHTML = '';
        const grid = document.createElement('div'); grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div'); card.className = 'template-card';
            const img = document.createElement('img'); img.className = 'template-thumb';
            img.src = entry.name.replace(/\.pptx$/i, '.png');
            img.addEventListener('error', () => {
                const ph = document.createElement('div'); ph.className = 'template-thumb-placeholder';
                ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>'; img.replaceWith(ph);
            });
            const nameDiv = document.createElement('div'); nameDiv.className = 'template-card-name'; nameDiv.textContent = entry.name.replace(/\.pptx$/i, '');
            card.appendChild(img); card.appendChild(nameDiv);
            card.addEventListener('click', async () => { const file = await entry.getFile(); this.selectTemplate({ name: entry.name, file }, card); });
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },
    async loadForPreview(file) {
        try {
            this.showLoading('Extracting notes...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml')).sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            this.originalSlides = [];
            for (const path of slideFiles) {
                const sFn = path.split('/').pop();
                const relsXml = zip.file(`ppt/slides/_rels/${sFn}.rels`) ? await zip.file(`ppt/slides/_rels/${sFn}.rels`).async('string') : null;
                const nPath = this.getNotesRelPath(relsXml);
                let nText = ""; 
                if (nPath && zip.file(nPath)) {
                    const nXml = await zip.file(nPath).async('string');
                    const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g; let pM;
                    while ((pM = pRegex.exec(nXml)) !== null) {
                        const tagRegex = /<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g; let match;
                        while ((match = tagRegex.exec(pM[1])) !== null) { if (match[0].startsWith('<a:br')) nText += '\n'; else nText += this.unescXml(match[2] || '').replace(/\u00A0/g, ' '); }
                        nText += '\n';
                    }
                }
                this.originalSlides.push({ path, notesPath: nPath, notes: nText.trim() });
            }
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides`;
            this.updatePreview(0); this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },
    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';
        if (this.originalSlides.length === 0) { container.innerHTML = '<div class="text-center py-20 text-slate-500 italic">No slides loaded.</div>'; return; }
        this.originalSlides.forEach((slide, idx) => {
            const wrapper = document.createElement('div'); wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div'); card.className = 'preview-card';
            const transposedText = this.transposeLine(slide.notes, semitones);
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left">Slide ${idx+1}</div><div class="whitespace-pre font-mono text-[11px] leading-snug text-left">${this.renderChordHTML(transposedText)}</div>`;
            wrapper.appendChild(card); container.appendChild(wrapper);
        });
        this.updateZoom();
    },
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
        let idx = this.musical.keys.indexOf(note); if (idx === -1) idx = this.musical.flats.indexOf(note);
        if (idx === -1) return note;
        let newIdx = (idx + semitones) % 12; if (newIdx < 0) newIdx += 12;
        return this.musical.preferred[newIdx];
    },
    renderChordHTML(text) { return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(this.chordRegex, '<span class="chord">$&</span>'); },
    showLoading(text) { this.elements.loadingText.textContent = text; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    setMode(mode) {
        const isGen = mode === 'gen';
        document.getElementById('modeGen').classList.toggle('active', isGen);
        document.getElementById('modeTrans').classList.toggle('active', !isGen);
        document.getElementById('viewGen').classList.toggle('hidden', !isGen);
        document.getElementById('viewTrans').classList.toggle('hidden', isGen);
    },
    updateZoom(val) {
        if (val === undefined) val = document.getElementById('zoomSlider').value;
        const zoomValEl = document.getElementById('zoomVal');
        if (zoomValEl) zoomValEl.textContent = val + '%';
        const scale = val / 100;
        const contents = document.getElementsByClassName('preview-card');
        for(let content of contents) { content.style.transform = `scale(${scale})`; content.style.transformOrigin = 'top center'; }
    },
    async changeSemitones(delta) {
        const current = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const next = Math.max(-11, Math.min(11, current + delta));
        this.elements.semitoneDisplay.textContent = (next > 0 ? '+' : '') + next;
        if (this.originalSlides.length > 0) this.updatePreview(next);
    }
};

App.init();
/* LyricSlide Pro - Version 4.0.0 (Absolute Full Production Build) */

const App = {
    version: "Version 4.0.0",
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
        if (this.elements.transFileInput) {
            this.elements.transFileInput.addEventListener('change', (e) => {
                if (e.target.files[0]) this.loadForPreview(e.target.files[0]);
            });
        }
        const alignSelect = document.getElementById('alignmentSelect');
        if (alignSelect) {
            alignSelect.addEventListener('change', () => {
                if (this.originalSlides.length > 0) this.updatePreview(0);
            });
        }
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        const vEl = document.getElementById('appVersion');
        if (vEl) vEl.textContent = this.version;
    },

    // --- THEME ENGINE ---
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
            document.querySelectorAll('.color-picker-input').forEach(p => p.addEventListener('input', (e) => {
                const idMap = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' };
                const varName = idMap[e.target.id];
                document.documentElement.style.setProperty(varName, e.target.value);
                const current = {};
                Object.keys(this.defaults).forEach(k => current[k] = getComputedStyle(document.documentElement).getPropertyValue(k).trim());
                localStorage.setItem('lyric_theme', JSON.stringify(current));
            }));
        },
        setVariable(name, val) { document.documentElement.style.setProperty(name, val); }
    },

    // --- CORE GENERATION (CLEAN MANIFEST REBUILD) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Input lyrics and select template.');

        try {
            this.showLoading('Wiping Clutter & Generating Slides...');
            const zip = await JSZip.loadAsync(file);

            // 1. DATA EXTRACTION
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const templatePath = slideRels[this.getSlideIds(presXml)[0].rid];
            const templateXml = await zip.file(`ppt/${templatePath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templatePath.split('/').pop()}.rels`).async('string');
            
            const templateNotesRelPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesRelPath ? await zip.file(templateNotesRelPath).async('string') : null;
            const templateNotesRelsXml = templateNotesRelPath ? await zip.file(`ppt/notesSlides/_rels/${templateNotesRelPath.split('/').pop()}.rels`).async('string') : null;

            // 2. CONTENT GENERATION
            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/i;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            const generated = [];

            // PHYSICALLY CLEAN THE ZIP
            Object.keys(zip.files).forEach(k => {
                if (k.includes('ppt/slides/slide') || k.includes('ppt/notesSlides/notesSlide')) zip.remove(k);
            });

            for (let i = 0; i < sections.length; i++) {
                const sName = `slide_gen_${i + 1}.xml`;
                const nName = `notes_gen_${i + 1}.xml`;
                
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', this.elements.songTitle.value || '');
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', this.elements.copyrightInfo.value || '');
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sections[i].trim(), document.getElementById('alignmentSelect').value);
                
                zip.file(`ppt/slides/${sName}`, slideXml);
                zip.file(`ppt/slides/_rels/${sName}.rels`, templateRelsXml.replace(/notesSlide\d+\.xml/g, nName));

                if (templateNotesXml) {
                    const noteLines = sections[i].trim().split(/\n/).map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formatted = this.escXml(noteLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">`);
                    zip.file(`ppt/notesSlides/${nName}`, templateNotesXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">${formatted}</a:t></a:r></a:p>`));
                    zip.file(`ppt/notesSlides/_rels/${nName}.rels`, templateNotesRelsXml.replace(/slide\d+\.xml/g, sName));
                }
                generated.push({ id: 256 + i, rid: `rId${100 + i}`, name: sName, nName: nName });
            }

            // 3. FULL REBUILD OF REGISTRIES (Stops Repair Prompt)
            
            // Build [Content_Types].xml
            let ctXml = await zip.file('[Content_Types].xml').async('string');
            ctXml = ctXml.replace(/<Override[^>]+PartName="\/ppt\/slides\/[^"]+"[^>]*\/>/g, '').replace(/<Override[^>]+PartName="\/ppt\/notesSlides\/[^"]+"[^>]*\/>/g, '');
            const ctEntries = generated.map(s => `<Override PartName="/ppt/slides/${s.name}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/notesSlides/${s.nName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`).join('');
            zip.file('[Content_Types].xml', ctXml.replace('</Types>', ctEntries + '</Types>'));

            // Build presentation.xml (Slide List)
            zip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst[^>]*>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('')}</p:sldIdLst>`));

            // Build presentation.xml.rels (rId Map)
            let relsBase = presRelsXml.replace(/<Relationship[^>]+Type="[^"]+slide"[^>]*\/>/g, '');
            const relsEntries = generated.map(s => `<Relationship Id="${s.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${s.name}"/>`).join('');
            zip.file('ppt/_rels/presentation.xml.rels', relsBase.replace('</Relationships>', relsEntries + '</Relationships>'));

            // Sync docProps Slide Count
            if (zip.file('docProps/app.xml')) {
                let appXml = await zip.file('docProps/app.xml').async('string');
                zip.file('docProps/app.xml', appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${generated.length}</Slides>`).replace(/<I4>\d+<\/I4>/, `<I4>${generated.length}</I4>`));
            }

            const blob = await zip.generateAsync({ type: 'blob', mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
            saveAs(blob, `${(this.elements.songTitle.value || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    // --- TRANSPOSITION ENGINE ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file || this.originalSlides.length === 0) return alert('Load a PPTX file first.');
        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file);
            for (const slide of this.originalSlides) {
                let xml = await zip.file(slide.path).async('string');
                zip.file(slide.path, this.transposeParagraphs(xml, semitones));
                if (slide.notesPath) {
                    const transposed = this.transposeLine(slide.notes, semitones);
                    let nXml = await zip.file(slide.notesPath).async('string');
                    const fmt = this.escXml(transposed).replace(/\n/g, `</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">`);
                    zip.file(slide.notesPath, nXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">${fmt}</a:t></a:r></a:p>`));
                }
            }
            saveAs(await zip.generateAsync({ type: 'blob' }), file.name.replace('.pptx', '_transposed.pptx'));
            this.hideLoading();
        } catch (e) { alert(e.message); this.hideLoading(); }
    },

    transposeParagraphs(xml, semitones) {
        return xml.replace(/<a:p[^>]*>([\s\S]*?)<\/a:p>/g, (match, pXml) => {
            let logicLine = "", charMeta = [], runRegex = /<a:r>([\s\S]*?)<\/a:r>|<a:br\/>/g, m;
            while ((m = runRegex.exec(pXml)) !== null) {
                if (m[0] === '<a:br/>') { logicLine += "\n"; charMeta.push({ isBr: true }); continue; }
                const rStyle = m[1].match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr/>';
                const text = this.unescXml(m[1].match(/<a:t[^>]*>(.*?)<\/a:t>/)?.[1] || "");
                for (let char of text) { charMeta.push({ isGhost: rStyle.includes('<a:noFill/>'), originalChar: char, style: rStyle }); logicLine += rStyle.includes('<a:noFill/>') ? " " : char; }
            }
            if (!logicLine.trim()) return match;
            const transposed = this.transposeLine(logicLine, semitones);
            let newRuns = "", metaIdx = 0;
            for (let i = 0; i < transposed.length; i++) {
                if (transposed[i] === "\n") { newRuns += "<a:br/>"; while(metaIdx < charMeta.length && !charMeta[metaIdx].isBr) metaIdx++; metaIdx++; continue; }
                const meta = charMeta[metaIdx] || { isGhost: false, style: '<a:rPr sz="1800"/>' };
                const finalStyle = (meta.isGhost && transposed[i] === " ") ? meta.style : this.getChordStyle(meta.style).replace('<a:noFill/>', '');
                newRuns += `<a:r>${finalStyle}<a:t xml:space="preserve">${this.escXml(meta.isGhost && transposed[i] === " " ? meta.originalChar : transposed[i]).replace(/ /g, '\u00A0')}</a:t></a:r>`;
                metaIdx++;
            }
            return match.replace(pXml, (pXml.match(/<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/)?.[0] || '') + newRuns);
        });
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        return text.split('\n').map(line => {
            if (!this.isChordLine(line)) return line;
            let res = line, off = 0;
            const matches = [...line.matchAll(this.chordRegex)];
            for (const m of matches) {
                const nr = this.shiftNote(m[1], semitones), nb = m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '';
                const nf = nr + (m[2] || '') + nb, p = m.index + off, d = nf.length - m[0].length;
                let pre = res.substring(0, p), suf = res.substring(p + m[0].length);
                if (d > 0 && suf.startsWith(' ')) { suf = suf.substring(1); off--; }
                else if (d < 0) { suf = ' '.repeat(Math.abs(d)) + suf; off += Math.abs(d); }
                res = pre + nf + suf; off += d;
            }
            return res;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let idx = this.musical.keys.indexOf(note); if (idx === -1) idx = this.musical.flats.indexOf(note); if (idx === -1) return note;
        let newIdx = (idx + semitones) % 12; if (newIdx < 0) newIdx += 12; return this.musical.preferred[newIdx];
    },

    // --- DATA LOADING & PREVIEW ---
    async loadForPreview(file) {
        try {
            this.showLoading('Reading Slides...');
            const zip = await JSZip.loadAsync(file);
            const slides = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide')).sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            this.originalSlides = [];
            for (const path of slides) {
                const rels = await zip.file(`ppt/slides/_rels/${path.split('/').pop()}.rels`)?.async('string');
                const nPath = this.getNotesRelPath(rels);
                let nTxt = "";
                if (nPath) {
                    const nXml = await zip.file(nPath).async('string');
                    const tMatches = nXml.match(/<a:t>(.*?)<\/a:t>/g);
                    if (tMatches) nTxt = tMatches.map(m => this.unescXml(m.replace(/<\/?a:t>/g, ''))).join(' ');
                }
                this.originalSlides.push({ path, notesPath: nPath, notes: nTxt.trim() });
            }
            this.updatePreview(0); this.hideLoading();
        } catch (e) { alert(e.message); this.hideLoading(); }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer'); container.innerHTML = '';
        this.originalSlides.forEach((s, i) => {
            const card = document.createElement('div'); card.className = 'preview-card-wrapper';
            const transposed = this.transposeLine(s.notes, semitones);
            card.innerHTML = `<div class="preview-card"><div class="text-[10px] text-slate-400 mb-1">Slide ${i+1}</div><div class="whitespace-pre font-mono text-[10px] text-left">${this.renderChordHTML(transposed)}</div></div>`;
            container.appendChild(card);
        });
    },

    renderChordHTML(text) {
        let html = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return html.replace(this.chordRegex, '<span class="chord">$&</span>');
    },

    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery'); if (!gallery) return;
        try {
            const res = await fetch('./templates.json'); const names = await res.json();
            const entries = names.map(name => ({ name, getFile: async () => { const r = await fetch(`./${encodeURIComponent(name)}`); const b = await r.blob(); return new File([b], name, { type: b.type }); } }));
            gallery.innerHTML = ''; const grid = document.createElement('div'); grid.className = 'template-grid';
            entries.forEach(entry => {
                const card = document.createElement('div'); card.className = 'template-card';
                const img = document.createElement('img'); img.className = 'template-thumb'; img.src = entry.name.replace(/\.pptx$/i, '.png');
                img.onerror = () => { const ph = document.createElement('div'); ph.className = 'template-thumb-placeholder'; ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>'; img.replaceWith(ph); };
                const nDiv = document.createElement('div'); nDiv.className = 'template-card-name'; nDiv.textContent = entry.name.replace(/\.pptx$/i, '');
                card.appendChild(img); card.appendChild(nDiv);
                card.onclick = async () => { 
                    this.selectedTemplateFile = await entry.getFile();
                    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                    card.classList.add('selected');
                    document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                    document.getElementById('selectedTemplateName').textContent = entry.name;
                };
                grid.appendChild(card);
            });
            gallery.appendChild(grid);
        } catch (e) { gallery.innerHTML = `<div class="text-center py-8 text-slate-400 italic">No templates available.</div>`; }
    },

    // --- XML UTILITIES ---
    lockInStyleAndReplace(xml, ph, val, align = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(ph), 'gi');
        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shape) => {
            if (!phRegex.test(shape)) return shape;
            const style = shape.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr lang="en-US"/>';
            if (!ph.toLowerCase().includes('lyrics')) return shape.replace(phRegex, val.split('\n').map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`));
            let injected = `</a:t></a:r></a:p>`;
            const lines = val.split('\n');
            for (let i = 0; i < lines.length; i++) {
                if (this.isChordLine(lines[i]) && lines[i+1] && !this.isChordLine(lines[i+1]) && !lines[i+1].trim().startsWith('[')) {
                    const max = Math.max(lines[i].length, lines[i+1].length);
                    injected += (align === 'ctr') ? this.makeGhostAlignmentLine(lines[i].padEnd(max,' '), lines[i+1].padEnd(max,' '), style, 'ctr') + this.makePptLine(lines[i+1].padEnd(max,' '), style, 'ctr') : this.makePptLine(lines[i], this.getChordStyle(style), 'l') + this.makePptLine(lines[i+1], style, 'l');
                    i++;
                } else {
                    const txt = lines[i].trim();
                    injected += txt ? this.makePptLine(txt, txt.startsWith('[') ? style.replace(/sz="\d+"/, 'sz="2000"') : style, align === 'ctr' ? 'ctr' : 'l') : `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                }
            }
            return shape.replace(phRegex, injected + `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t xml:space="preserve">`).replace('</a:bodyPr>', '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/></a:bodyPr>');
        });
    },

    makeGhostAlignmentLine(c, l, s, a) {
        let gh = s.replace('<a:rPr', '<a:rPr><a:noFill/>').replace(/<a:solidFill>.*?<\/a:solidFill>/g, ''), xml = "";
        for (let i = 0; i < c.length; i++) xml += (c[i] === ' ') ? `<a:r>${gh}<a:t xml:space="preserve">${this.escXml(l[i] || ' ')}</a:t></a:r>` : `<a:r>${this.getChordStyle(s)}<a:t xml:space="preserve">${this.escXml(c[i])}</a:t></a:r>`;
        return `<a:p><a:pPr algn="${a}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr>${xml}</a:p>`;
    },
    makePptLine(t, s, a) { return `<a:p><a:pPr algn="${a}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr><a:r>${s}<a:t xml:space="preserve">${this.escXml(t)}</a:t></a:r></a:p>`; },
    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(rXml) { const m = rXml?.match(/Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    isChordLine(l) { if(!l) return false; const t = l.trim(), w = t.split(/\s+/), c = t.match(this.chordRegex) || []; return c.length >= w.length * 0.5 || (c.length > 0 && w.length <= 2); },
    getChordStyle(s) { const f = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>'; let res = s.includes('sz=') ? s.replace(/sz="\d+"/, 'sz="1800"') : s.replace('<a:rPr', '<a:rPr sz="1800"'); return res.includes('<a:solidFill>') ? res.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, f) : res.replace('</a:rPr>', f + '</a:rPr>'); },
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; }
};

App.init();
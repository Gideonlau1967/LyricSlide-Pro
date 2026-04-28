/* LyricSlide Pro - Version 3.8.1 (Final Full Integrity Build) */

const App = {
    version: "Version 3.8.1",
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
    originalSlides: [], selectedTemplateFile: null, 
    
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
        if (document.getElementById('appVersion')) document.getElementById('appVersion').textContent = this.version;
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
                const map = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' };
                this.setVariable(map[e.target.id], e.target.value);
                localStorage.setItem('lyric_theme', JSON.stringify(Object.keys(this.defaults).reduce((a, k) => ({ ...a, [k]: getComputedStyle(document.documentElement).getPropertyValue(k).trim() }), {})));
            }));
        },
        setVariable(name, val) { document.documentElement.style.setProperty(name, val); }
    },

    // --- CORE GENERATION ---
    async generate() {
        const file = this.selectedTemplateFile;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Input lyrics and select template.');

        try {
            this.showLoading('Generating Multi-Slide PPTX...');
            const zip = await JSZip.loadAsync(file);

            // 1. ANALYZE TEMPLATE
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const templatePath = slideRels[this.getSlideIds(presXml)[0].rid];
            const templateXml = await zip.file(`ppt/${templatePath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templatePath.split('/').pop()}.rels`).async('string');
            
            const templateNotesRelPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesRelPath ? await zip.file(templateNotesRelPath).async('string') : null;
            const templateNotesRelsXml = templateNotesRelPath ? await zip.file(`ppt/notesSlides/_rels/${templateNotesRelPath.split('/').pop()}.rels`).async('string') : null;

            // 2. PROCESS LYRICS
            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/i;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            const generated = [];

            // Wipe existing slide files
            Object.keys(zip.files).forEach(k => { if (k.includes('ppt/slides/slide') || k.includes('ppt/notesSlides/notesSlide')) zip.remove(k); });

            for (let i = 0; i < sections.length; i++) {
                const sName = `slide_gen_${i + 1}.xml`;
                const nName = `notes_gen_${i + 1}.xml`;
                
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', this.elements.songTitle.value || '');
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', this.elements.copyrightInfo.value || '');
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sections[i].trim(), document.getElementById('alignmentSelect').value);
                
                zip.file(`ppt/slides/${sName}`, slideXml);
                zip.file(`ppt/slides/_rels/${sName}.rels`, templateRelsXml.replace(/Target="..\/notesSlides\/[^"]+"/, `Target="../notesSlides/${nName}"`));

                if (templateNotesXml) {
                    const noteLines = sections[i].trim().split(/\n/).map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formatted = this.escXml(noteLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">`);
                    zip.file(`ppt/notesSlides/${nName}`, templateNotesXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">${formatted}</a:t></a:r></a:p>`));
                    zip.file(`ppt/notesSlides/_rels/${nName}.rels`, templateNotesRelsXml.replace(/Target="..\/slides\/[^"]+"/, `Target="../slides/${sName}"`));
                }
                generated.push({ id: 256 + i, rid: `rId${100 + i}`, name: sName, nName: nName });
            }

            // 3. REBUILD MANIFESTS (String-based to prevent Namespace corruption)
            let ctXml = await zip.file('[Content_Types].xml').async('string');
            ctXml = ctXml.replace(/<Override[^>]+PartName="\/ppt\/slides\/[^"]+"[^>]*\/>/g, '').replace(/<Override[^>]+PartName="\/ppt\/notesSlides\/[^"]+"[^>]*\/>/g, '');
            const ctEntries = generated.map(s => `<Override PartName="/ppt/slides/${s.name}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/notesSlides/${s.nName}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`).join('');
            zip.file('[Content_Types].xml', ctXml.replace('</Types>', ctEntries + '</Types>'));

            zip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst[^>]*>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('')}</p:sldIdLst>`));

            let relsBase = presRelsXml.replace(/<Relationship[^>]+Type="[^"]+slide"[^>]*\/>/g, '');
            const relsEntries = generated.map(s => `<Relationship Id="${s.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${s.name}"/>`).join('');
            zip.file('ppt/_rels/presentation.xml.rels', relsBase.replace('</Relationships>', relsEntries + '</Relationships>'));

            if (zip.file('docProps/app.xml')) {
                let appXml = await zip.file('docProps/app.xml').async('string');
                zip.file('docProps/app.xml', appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${generated.length}</Slides>`).replace(/<I4>\d+<\/I4>/, `<I4>${generated.length}</I4>`));
            }

            saveAs(await zip.generateAsync({ type: 'blob', mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }), `${(this.elements.songTitle.value || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    // --- TRANSPOSITION ENGINE ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file || this.originalSlides.length === 0) return alert('Load file first.');
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
                const finalStyle = meta.isGhost && transposed[i] === " " ? meta.style : this.getChordStyle(meta.style).replace('<a:noFill/>', '');
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

    // --- UTILITIES ---
    async loadForPreview(file) {
        try {
            this.showLoading('Extracting...');
            const zip = await JSZip.loadAsync(file);
            const slides = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide')).sort();
            this.originalSlides = [];
            for (const path of slides) {
                const rels = await zip.file(`ppt/slides/_rels/${path.split('/').pop()}.rels`)?.async('string');
                const notesPath = this.getNotesRelPath(rels);
                let notesTxt = "";
                if (notesPath) {
                    const nXml = await zip.file(notesPath).async('string');
                    const tMatches = nXml.match(/<a:t>(.*?)<\/a:t>/g);
                    if (tMatches) notesTxt = tMatches.map(m => this.unescXml(m.replace(/<\/?a:t>/g, ''))).join(' ');
                }
                this.originalSlides.push({ path, notesPath, notes: notesTxt.trim() });
            }
            this.updatePreview(0); this.hideLoading();
        } catch (e) { alert(e.message); this.hideLoading(); }
    },
    updatePreview(semitones) {
        const container = document.getElementById('previewContainer'); container.innerHTML = '';
        this.originalSlides.forEach(
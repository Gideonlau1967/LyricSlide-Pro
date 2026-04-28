/* 
   LyricSlide Pro - Version 3.1 (COMPLETE MERGED VERSION)
   Includes all UI logic from 3.0 + Structural Fixes for MS Office Repair Error.
*/

const App = {
    version: "Version 3.1 Stable",
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
        if (alignSelect) {
            alignSelect.addEventListener('change', () => { if (this.originalSlides.length > 0) this.updatePreview(0); });
        }

        // Initialize UI components from 3.0
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;

        const v = document.getElementById('appVersion'); 
        if (v) v.textContent = this.version;
    },

    // --- UI HELPERS (From 3.0) ---
    setMode(mode) {
        const isGen = mode === 'gen';
        document.getElementById('modeGen').classList.toggle('active', isGen);
        document.getElementById('modeTrans').classList.toggle('active', !isGen);
        document.getElementById('viewGen').classList.toggle('hidden', !isGen);
        document.getElementById('viewTrans').classList.toggle('hidden', isGen);
    },

    async changeSemitones(delta) {
        const current = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const next = Math.max(-11, Math.min(11, current + delta));
        this.elements.semitoneDisplay.textContent = (next > 0 ? '+' : '') + next;
        if (this.originalSlides.length > 0) this.updatePreview(next);
    },

    // --- THEME MANAGEMENT ---
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
            const map = {
                'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start',
                'picker-bg-end': '--bg-end', 'picker-text': '--text-main',
                'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg',
                'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color'
            };
            return map[id];
        },
        setVariable(name, val) {
            document.documentElement.style.setProperty(name, val);
            if (name === '--primary-color') document.documentElement.style.setProperty('--primary-gradient', val);
        },
        save() {
            const current = {};
            Object.keys(this.defaults).forEach(key => {
                current[key] = getComputedStyle(document.documentElement).getPropertyValue(key).trim();
            });
            localStorage.setItem('lyric_theme', JSON.stringify(current));
        }
    },

    // --- CORE GENERATION ENGINE (FIXED) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Select a template and input lyrics.');

        try {
            this.showLoading('Generating PPTX...');
            const zip = await JSZip.loadAsync(file);
            
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const contentTypesXml = await zip.file('[Content_Types].xml').async('string');

            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templateRelPath.split('/').pop()}.rels`).async('string');
            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionText = sections[i].trim();
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);

                const name = `song_gen_${i + 1}.xml`;
                zip.file(`ppt/slides/${name}`, slideXml);
                
                let notesName = null;
                if (templateNotesXml) {
                    notesName = `notes_gen_${i + 1}.xml`;
                    const style = '<a:rPr lang="en-US" sz="1600"/>';
                    const noteLines = sectionText.split(/\n/).map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formattedNotes = this.escXml(noteLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    
                    zip.file(`ppt/notesSlides/${notesName}`, templateNotesXml.replace(new RegExp(this.getPlaceholderRegexStr('[Presenter Note]'), 'gi'), formattedNotes));
                    zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml.replace(/Target="..\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/${notesName}"`));
                    zip.file(`ppt/notesSlides/_rels/${notesName}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${name}"/></Relationships>`);
                }
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, notesName });
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            this.updateContentTypes(zip, contentTypesXml, generated);

            const blob = await zip.generateAsync({ type: 'blob' });
            saveAs(blob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { console.error(err); alert(err.message); this.hideLoading(); }
    },

    // --- STRUCTURAL PPT FIXES (3.1) ---
    syncPresentationRegistry(zip, xml, rels, gen) {
        const sldIdList = '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', xml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdList));

        let doc = new DOMParser().parseFromString(rels, 'application/xml');
        let rNode = doc.documentElement;
        const ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        [...rNode.getElementsByTagName('Relationship')].forEach(n => {
            if (n.getAttribute('Type').endsWith('slide')) n.remove();
        });

        gen.forEach(s => {
            let e = doc.createElementNS(ns, 'Relationship');
            e.setAttribute('Id', s.rid);
            e.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            e.setAttribute('Target', `slides/${s.name}`);
            rNode.appendChild(e);
        });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(doc));
    },

    updateContentTypes(zip, xml, gen) {
        let doc = new DOMParser().parseFromString(xml, 'application/xml');
        let root = doc.documentElement;
        const ns = "http://schemas.openxmlformats.org/package/2006/content-types";

        [...root.getElementsByTagName('Override')].forEach(n => {
            const pn = n.getAttribute('PartName');
            if (pn.includes('/ppt/slides/') || pn.includes('/ppt/notesSlides/')) n.remove();
        });

        gen.forEach(s => {
            let slideNode = doc.createElementNS(ns, 'Override');
            slideNode.setAttribute('PartName', `/ppt/slides/${s.name}`);
            slideNode.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            root.appendChild(slideNode);

            if (s.notesName) {
                let notesNode = doc.createElementNS(ns, 'Override');
                notesNode.setAttribute('PartName', `/ppt/notesSlides/${s.notesName}`);
                notesNode.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
                root.appendChild(notesNode);
            }
        });
        zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(doc));
    },

    // --- TRANSPOSITION ENGINE ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file || this.originalSlides.length === 0) return alert('Select file and wait for load.');
        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file);
            for (const slide of this.originalSlides) {
                let slideXml = await zip.file(slide.path).async('string');
                slideXml = this.transposeParagraphs(slideXml, semitones);
                zip.file(slide.path, slideXml);
                if (slide.notesPath) {
                    const transposedNotes = this.transposeLine(slide.notes, semitones);
                    let notesXml = await zip.file(slide.notesPath).async('string');
                    const style = '<a:rPr lang="en-US" sz="1200"/>';
                    const formatted = this.escXml(transposedNotes).replace(/\n/g, `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    notesXml = notesXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r>${style}<a:t xml:space="preserve">${formatted}</a:t></a:r></a:p>`);
                    zip.file(slide.notesPath, notesXml);
                }
            }
            saveAs(await zip.generateAsync({ type: 'blob' }), file.name.replace('.pptx', `_transposed.pptx`));
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    transposeParagraphs(xml, semitones) {
        return xml.replace(/<a:p[^>]*>([\s\S]*?)<\/a:p>/g, (matchFull, pXml) => {
            let logicLine = "", charMeta = []; 
            const runRegex = /<a:r>([\s\S]*?)<\/a:r>|<a:br\/>/g;
            let m;
            while ((m = runRegex.exec(pXml)) !== null) {
                if (m[0] === '<a:br/>') { logicLine += "\n"; charMeta.push({ isBr: true }); continue; }
                const rStyle = m[1].match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr/>';
                const isGhost = rStyle.includes('<a:noFill/>');
                const text = this.unescXml(m[1].match(/<a:t[^>]*>(.*?)<\/a:t>/)?.[1] || "");
                for (let char of text) { charMeta.push({ isGhost, originalChar: char, style: rStyle, isBr: false }); logicLine += isGhost ? " " : char; }
            }
            if (!logicLine.trim()) return matchFull;
            const transposedLogic = this.transposeLine(logicLine, semitones);
            if (transposedLogic === logicLine) return matchFull;
            const pPr = pXml.match(/<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/)?.[0] || '';
            const pTagOpen = matchFull.match(/^<a:p[^>]*>/)?.[0] || '<a:p>';
            let newRuns = "", metaIdx = 0;
            for (let i = 0; i < transposedLogic.length; i++) {
                const newChar = transposedLogic[i];
                if (newChar === "\n") { newRuns += "<a:br/>"; while(metaIdx < charMeta.length && !charMeta[metaIdx].isBr) metaIdx++; metaIdx++; continue; }
                const meta = charMeta[metaIdx] || { isGhost: false, style: '<a:rPr sz="1800"/>' };
                if (meta.isGhost && newChar === " ") {
                    newRuns += `<a:r>${meta.style}<a:t xml:space="preserve">${this.escXml(meta.originalChar).replace(/ /g, '\u00A0')}</a:t></a:r>`;
                } else {
                    const chordStyle = this.getChordStyle(meta.style).replace('<a:noFill/>', ''); 
                    newRuns += `<a:r>${chordStyle}<a:t xml:space="preserve">${this.escXml(newChar).replace(/ /g, '\u00A0')}</a:t></a:r>`;
                }
                metaIdx++;
            }
            return `${pTagOpen}${pPr}${newRuns}</a:p>`;
        });
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
        let idx = this.musical.keys.indexOf(note);
        if (idx === -1) idx = this.musical.flats.indexOf(note);
        if (idx === -1) return note;
        let newIdx = (idx + semitones) % 12;
        if (newIdx < 0) newIdx += 12;
        return this.musical.preferred[newIdx];
    },

    // --- LOGIC HELPERS ---
    isChordLine(lineStr) {
        if (!lineStr || typeof lineStr !== 'string') return false;
        const trimmed = lineStr.trim();
        if (trimmed === '' || /^(A|I|The|And|Then|They|We|He|She)\s+[a-zA-Z]{2,}/i.test(trimmed)) return false;
        const words = trimmed.split(/\s+/), chords = trimmed.match(this.chordRegex) || [];
        return chords.length >= words.length * 0.5 || (chords.length > 0 && words.length <= 2);
    },

    getChordStyle(lyricStyle) {
        let s = lyricStyle.includes('sz=') ? lyricStyle.replace(/sz="\d+"/, 'sz="1800"') : lyricStyle.replace('<a:rPr', '<a:rPr sz="1800"');
        const greyFill = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>';
        return s.includes('<a:solidFill>') ? s.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, greyFill) : s.replace('</a:rPr>', greyFill + '</a:rPr>');
    },

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

    // --- UTILS ---
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(slideRelsXml) { const m = slideRelsXml?.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },

    async loadForPreview(file) {
        try {
            this.showLoading('Extracting notes...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml')).sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            this.originalSlides = [];
            for (const path of slideFiles) {
                const sfn = path.split('/').pop();
                const relsXml = zip.file(`ppt/slides/_rels/${sfn}.rels`) ? await zip.file(`ppt/slides/_rels/${sfn}.rels`).async('string') : null;
                const notesPath = this.getNotesRelPath(relsXml);
                let notesText = ""; 
                if (notesPath && zip.file(notesPath)) {
                    const nx = await zip.file(notesPath).async('string');
                    const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g; let pm;
                    while ((pm = pRegex.exec(nx)) !== null) {
                        const pc = pm[1]; const tr = /<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g;
                        let pt = ''; let m;
                        while ((m = tr.exec(pc)) !== null) {
                            if (m[0].startsWith('<a:br')) pt += '\n';
                            else pt += this.unescXml(m[2] || '').replace(/\u00A0/g, ' ');
                        }
                        notesText += pt + '\n';
                    }
                }
                this.originalSlides.push({ path, notesPath, notes: notesText.trim() });
            }
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides`;
            this.updatePreview(0);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';
        if (this.originalSlides.length === 0) return;
        this.originalSlides.forEach((slide, idx) => {
            const wrapper = document.createElement('div'); wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div'); card.className = 'preview-card';
            const transText = this.transposeLine(slide.notes, semitones);
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left">Slide ${idx + 1}</div><div class="whitespace-pre font-mono text-[11px] leading-snug text-left">${this.renderChordHTML(transText)}</div>`;
            wrapper.appendChild(card); container.appendChild(wrapper);
        });
        this.updateZoom();
    },

    renderChordHTML(text) {
        let html = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return html.replace(this.chordRegex, '<span class="chord">$&</span>');
    },

    updateZoom(val) {
        if (val === undefined) val = document.getElementById('zoomSlider').value;
        const zv = document.getElementById('zoomVal'); if (zv) zv.textContent = val + '%';
        const sc = val / 100;
        const contents = document.getElementsByClassName('preview-card');
        for(let c of contents) { c.style.transform = `scale(${sc})`; c.style.transformOrigin = 'top center'; }
    },

    showLoading(text) { this.elements.loadingText.textContent = text; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },

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
        const gallery = document.getElementById('templateGallery'); gallery.innerHTML = '';
        const grid = document.createElement('div'); grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div'); card.className = 'template-card';
            const img = document.createElement('img'); img.className = 'template-thumb';
            img.src = entry.name.replace(/\.pptx$/i, '.png');
            img.addEventListener('error', () => {
                const ph = document.createElement('div'); ph.className = 'template-thumb-placeholder';
                ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>'; img.replaceWith(ph);
            });
            const nd = document.createElement('div'); nd.className = 'template-card-name'; nd.textContent = entry.name.replace(/\.pptx$/i, '');
            card.appendChild(img); card.appendChild(nd);
            card.addEventListener('click', async () => {
                const file = await entry.getFile(); this.selectTemplate({ name: entry.name, file }, card);
            });
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },

    selectTemplate(item, cardEl) {
        this.selectedTemplateFile = item.file;
        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
        cardEl.classList.add('selected');
        document.getElementById('selectedTemplateInfo').classList.remove('hidden');
        document.getElementById('selectedTemplateName').textContent = item.name;
    }
};

App.init();
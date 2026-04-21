/* LyricSlide Pro - Version 2.7 */

const App = {
    version: "2.7",
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
        flats: ['C', 'Db', 'D', 'Eb', 'E', 'F', 'Gb', 'G', 'Ab', 'A', 'Bb', 'B']
    },

    originalSlides: [],   
    selectedTemplateFile: null, 
    songTitle: "",

    init() {
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        
        // Add listener to update preview when alignment changes
        const alignSelect = document.getElementById('alignmentSelect');
        if (alignSelect) {
            alignSelect.addEventListener('change', () => {
                if (this.originalSlides.length > 0) this.updatePreview(0);
            });
        }

        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;

        const versionEl = document.getElementById('appVersion');
        if (versionEl) {
            versionEl.textContent = this.version;
        }
        
        console.log(`App Initialized. Version ${this.version}`);
    },

    // --- [PROTECTED] THEME MANAGEMENT (UNTOUCHED) ---
    theme: {
        defaults: {
            '--primary-color': '#334155',
            '--bg-start': '#f8fafc',
            '--bg-end': '#f8fafc',
            '--text-main': '#1e293b',
            '--card-accent': '#e2e8f0',
            '--preview-card-bg': '#ffffff',
            '--preview-chord-color': '#334155',
            '--preview-lyrics-color': '#1e293b'
        },
        init() {
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                this.setVariable(key, val);
                const pickerId = 'picker-' + key.replace('--', '').replace('-color', '');
                const picker = document.getElementById(pickerId);
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
        setVariable(name, val) {
            document.documentElement.style.setProperty(name, val);
            if (name === '--primary-color') document.documentElement.style.setProperty('--primary-gradient', val);
        },
        save() {
            const current = {};
            Object.keys(this.defaults).forEach(key => { current[key] = getComputedStyle(document.documentElement).getPropertyValue(key).trim(); });
            localStorage.setItem('lyric_theme', JSON.stringify(current));
        },
        reset() {
            if (confirm('Reset theme to default minimal colors?')) {
                Object.keys(this.defaults).forEach(key => {
                    this.setVariable(key, this.defaults[key]);
                    const pickerId = 'picker-' + key.replace('--', '').replace('-color', '');
                    const picker = document.getElementById(pickerId);
                    if (picker) picker.value = this.defaults[key];
                });
                this.save();
            }
        }
    },

    // --- [PROTECTED] UI HELPERS (UNTOUCHED) ---
    setMode(mode) {
        const isGen = mode === 'gen';
        document.getElementById('modeGen').classList.toggle('active', isGen);
        document.getElementById('modeTrans').classList.toggle('active', !isGen);
        document.getElementById('viewGen').classList.toggle('hidden', !isGen);
        document.getElementById('viewTrans').classList.toggle('hidden', isGen);
    },
    updateZoom(val) {
        if (val === undefined) val = document.getElementById('zoomSlider').value;
        document.getElementById('zoomVal').textContent = val + '%';
        const scale = val / 100;
        const contents = document.getElementsByClassName('slide-content');
        for(let content of contents) { content.style.transform = `scale(${scale})`; }
    },
    async changeSemitones(delta) {
        const current = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const next = Math.max(-11, Math.min(11, current + delta));
        this.elements.semitoneDisplay.textContent = (next > 0 ? '+' : '') + next;
        if (this.originalSlides.length > 0) this.updatePreview(next);
    },
    toggleThemeSidebar() {
        document.getElementById('themeSidebar').classList.toggle('open');
        document.getElementById('sidebarBackdrop').classList.toggle('open');
    },

    // --- [PROTECTED] PARSING ENGINE (RESTORED FROM WORKING COPY) ---
    async loadForPreview(file) {
        try {
            this.showLoading('Extracting slide text...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files)
                .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
                .sort((a, b) => {
                    const numA = parseInt(a.match(/\d+/)[0]);
                    const numB = parseInt(b.match(/\d+/)[0]);
                    return numA - numB;
                });
            this.originalSlides = [];
            let globalSongTitle = "";
            for (const path of slideFiles) {
                const xml = await zip.file(path).async('string');
                const slideData = [];
                const spRegex = /<p:sp>([\s\S]*?)<\/p:sp>/g;
                let spMatch;
                while ((spMatch = spRegex.exec(xml)) !== null) {
                    const spContent = spMatch[1];
                    const phMatch = spContent.match(/<p:ph[^>]*type="(?:title|ctrTitle|ftr|dt|sldNum)"/);
                    const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g;
                    let pMatch;
                    while ((pMatch = pRegex.exec(spContent)) !== null) {
                        const pContent = pMatch[1];
                        const tagRegex = /<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g;
                        let pText = '';
                        let match;
                        while ((match = tagRegex.exec(pContent)) !== null) {
                            if (match[0].startsWith('<a:br')) pText += '\n';
                            else pText += this.unescXml(match[2] || '');
                        }
                        if (phMatch && (phMatch[0].includes('title') || phMatch[0].includes('ctrTitle')) && pText.trim() && !globalSongTitle) {
                            globalSongTitle = pText.trim();
                        }
                        slideData.push({ text: pText, isTitle: !!phMatch });
                    }
                }
                this.originalSlides.push(slideData);
            }
            this.songTitle = globalSongTitle;
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides Loaded`;
            this.updatePreview(0);
            this.hideLoading();
        } catch (err) { console.error(err); alert("Error loading preview: " + err.message); this.hideLoading(); }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        const userAlign = document.getElementById('alignmentSelect').value;
        container.innerHTML = '';
        if (this.originalSlides.length === 0) return;
        this.originalSlides.forEach((slideData, idx) => {
            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left">Slide ${idx + 1}</div>`;
            const contentDiv = document.createElement('div');
            contentDiv.className = 'slide-content'; 
            slideData.forEach((para) => {
                const text = para.text;
                const isTitle = para.isTitle || (this.songTitle && text.trim().toLowerCase() === this.songTitle.toLowerCase());
                const isMetadata = /©|Copyright|Words:|Music:|Lyrics:|Chris Tomlin|CCLI|DAYEG AMBASSADOR/i.test(text);
                if (text.trim() && !isMetadata && !isTitle) {
                    const lineDiv = document.createElement('div');
                    lineDiv.style.textAlign = (userAlign === 'l' ? 'left' : 'center'); 
                    lineDiv.innerHTML = this.renderChordHTML(this.transposeLine(text, semitones));
                    contentDiv.appendChild(lineDiv);
                }
            });
            card.appendChild(contentDiv);
            container.appendChild(card);
        });
        this.updateZoom();
    },

    // --- [LATEST] PRO-ALIGNMENT LOGIC (MERGED) ---
    lockInStyleAndReplace(xml, placeholder, replacement, userAlign = 'ctr') {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');
        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));
                const rawLines = (replacement || '').split(/\r?\n/);
                if (placeholder !== '[Lyrics and Chords]') {
                    return shapeXml.replace(phRegex, rawLines.map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`));
                }
                let maxLen = 0;
                rawLines.forEach(l => { if(l.length > maxLen) maxLen = l.length; });
                let injectedXml = `</a:t></a:r></a:p>`;
                for (let i = 0; i < rawLines.length; i++) {
                    let line = rawLines[i]; let nextLine = rawLines[i + 1];
                    if (this.isChordLine(line) && nextLine !== undefined && !this.isChordLine(nextLine) && !nextLine.trim().startsWith('[')) {
                        let chordText = line; let lyricText = nextLine;
                        let finalChordAlign = 'l'; 
                        let finalLyricAlign = (userAlign === 'l') ? 'l' : 'ctr';
                        if (userAlign === 'smart' || userAlign === 'ctr') {
                            const push = Math.floor((maxLen - lyricText.length) / 2);
                            chordText = " ".repeat(push) + chordText;
                        }
                        injectedXml += this.makeMixedStyleLine(chordText, style, finalChordAlign);
                        injectedXml += this.makePptLine(lyricText.trim(), style, finalLyricAlign);
                        i++; 
                    } else if (line.trim() !== "") {
                        let sAlign = (userAlign === 'l' ? 'l' : 'ctr');
                        injectedXml += this.makePptLine(line.trim(), style, sAlign);
                    } else {
                        injectedXml += `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                    }
                }
                injectedXml += `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r>${style}<a:t xml:space="preserve">`;
                let result = shapeXml.replace(phRegex, () => injectedXml).replace(/<a:p><a:pPr[^>]*><a:buNone\/><\/a:pPr><a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r><\/a:p>/g, '');
                if (!result.includes('Autofit')) result = result.replace('</a:bodyPr>', '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/></a:bodyPr>');
                return result;
            }
            return shapeXml;
        });
    },

    makeMixedStyleLine(text, lyricStyle, align) {
        const chordStyle = lyricStyle.includes('sz=') ? lyricStyle.replace(/sz="\d+"/, 'sz="1800"') : lyricStyle.replace('<a:rPr', '<a:rPr sz="1800"');
        let runsXml = ""; const segments = text.split(/(\s+)/);
        segments.forEach(seg => {
            if (!seg) return;
            const style = /^\s+$/.test(seg) ? lyricStyle : chordStyle;
            runsXml += `<a:r>${style}<a:t xml:space="preserve">${this.escXml(seg).replace(/ /g, '\u00A0')}</a:t></a:r>`;
        });
        return `<a:p><a:pPr algn="${align}"><a:buNone/></a:pPr>${runsXml}</a:p>`;
    },

    makePptLine(text, style, align) {
        return `<a:p>
            <a:pPr algn="${align}"><a:buNone/></a:pPr>
            <a:r>${style}<a:t xml:space="preserve">${this.escXml(text).replace(/ /g, '\u00A0')}</a:t></a:r>
        </a:p>`;
    },

    isChordLine(text) {
        if (!text || text.trim() === '') return false;
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        const matches = text.match(chordRegex) || [];
        const words = text.trim().split(/\s+/).filter(w => w.length > 0);
        return matches.length > 0 && (matches.length >= words.length * 0.4 || words.length < 3);
    },

    // --- [PROTECTED] TRANSPOSE ENGINE (RESTORED FROM WORKING COPY) ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const getSz = id => { const v = parseFloat(document.getElementById(id).value); return (!isNaN(v) && v > 0) ? Math.round(v * 100) : null; };
        const getFont = id => document.getElementById(id).value.trim();
        const tFont = getFont('fontTitle'); const tSz = getSz('fontSizeTitle');
        const lFont = getFont('fontLyrics'); const lSz = getSz('fontSizeLyrics');
        const cFont = getFont('fontCopyright'); const cSz = getSz('fontSizeCopyright');
        if (!file || (semitones === 0 && !tFont && !lFont && !cFont)) return alert('Select file and changes.');
        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');
                if (semitones !== 0) content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, t) => `<a:t>${this.transposeLine(t, semitones)}</a:t>`);
                if (tFont || tSz || lFont || lSz || cFont || cSz) {
                    content = content.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (m, c) => {
                        const isT = /<p:ph[^>]*type="(?:title|ctrTitle)"/.test(c);
                        const isC = /type="ftr"|©|copyright|ccli/i.test(c.replace(/<[^>]+>/g, ''));
                        let f, s; if (isT) { f = tFont; s = tSz; } else if (isC) { f = cFont; s = cSz; } else { f = lFont; s = lSz; }
                        return (f || s) ? `<p:sp>${this.applyFontToShapeXml(c, f, s)}</p:sp>` : m;
                    });
                }
                zip.file(path, content);
            }
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, file.name.replace('.pptx', '_transposed.pptx'));
            this.hideLoading();
        } catch (err) { alert("Error: " + err.message); this.hideLoading(); }
    },
    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        const lines = text.split('\n');
        return lines.map(line => {
            const matches = [...line.matchAll(chordRegex)];
            if (matches.length === 0 || matches.length < line.split(/\s+/).filter(w=>w.length>0).length * 0.4) return line;
            let res = line; let off = 0;
            for (const m of matches) {
                const orig = m[0]; const pos = m.index + off;
                const root = this.shiftNote(m[1], semitones);
                const bass = m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '';
                const nc = root + (m[2] || '') + bass;
                const diff = nc.length - orig.length;
                res = res.substring(0, pos) + nc + res.substring(pos + orig.length);
                if (diff > 0) {
                    let sm = res.substring(pos + nc.length).match(/^ +/);
                    if (sm && sm[0].length >= diff) res = res.substring(0, pos + nc.length) + res.substring(pos + nc.length + diff);
                    else off += diff;
                } else if (diff < 0) {
                    res = res.substring(0, pos + nc.length) + " ".repeat(Math.abs(diff)) + res.substring(pos + nc.length);
                }
            }
            return res;
        }).join('\n');
    },
    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(note);
        if (idx === -1) { list = (list === this.musical.keys ? this.musical.flats : this.musical.keys); idx = list.indexOf(note); }
        if (idx === -1) return note;
        return (semitones >= 0 ? this.musical.keys : this.musical.flats)[(idx + semitones + 12) % 12];
    },
    applyFontToShapeXml(shapeXml, fontFamily, fontSizeHundredths) {
        const build = (f) => f ? `<a:latin typeface="${this.escXml(f)}"/><a:ea typeface="${f}"/><a:cs typeface="${f}"/>` : '';
        const szAttr = (a, s) => s ? (/\bsz="[^"]+"/.test(a) ? a.replace(/\bsz="[^"]+"/, `sz="${s}"`) : a + ` sz="${s}"`) : a;
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)\/>/g, (_, a) => `<a:rPr${szAttr(a, fontSizeHundredths)}>${build(fontFamily)}</a:rPr>`);
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/g, (_, a, i) => {
            if (fontFamily) i = i.replace(/<a:(latin|ea|cs)[^>]*(\/>|>[\s\S]*?<\/a:\1>)/g, '');
            return `<a:rPr${szAttr(a, fontSizeHundredths)}>${build(fontFamily)}${i}</a:rPr>`;
        });
        return shapeXml;
    },

    // --- [PROTECTED] XML UTILS (UNTOUCHED) ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
            if (!res.ok) throw new Error();
            const names = await res.json();
            const entries = names.map(name => ({ name, getFile: async () => { const r = await fetch(`./${encodeURIComponent(name)}`); const b = await r.blob(); return new File([b], name, { type: b.type }); } }));
            this.renderTemplateGallery(entries);
        } catch (e) { if(gallery) gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Could not read templates.json.</div>`; }
    },
    renderTemplateGallery(entries) {
        const gallery = document.getElementById('templateGallery'); if(!gallery) return;
        gallery.innerHTML = ''; const grid = document.createElement('div'); grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div'); card.className = 'template-card';
            const img = document.createElement('img'); img.className = 'template-thumb'; img.src = entry.name.replace(/\.pptx$/i, '.png');
            img.onerror = () => { const ph = document.createElement('div'); ph.className = 'template-thumb-placeholder'; ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>'; img.replaceWith(ph); };
            const nameDiv = document.createElement('div'); nameDiv.className = 'template-card-name'; nameDiv.textContent = entry.name.replace(/\.pptx$/i, '');
            card.appendChild(img); card.appendChild(nameDiv);
            card.addEventListener('click', async () => { card.style.opacity = '0.6'; const file = await entry.getFile(); card.style.opacity = '1'; this.selectTemplate({ name: entry.name, file }, card); });
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },
    selectTemplate(item, cardEl) { this.selectedTemplateFile = item.file; document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected')); cardEl.classList.add('selected'); document.getElementById('selectedTemplateInfo').classList.remove('hidden'); document.getElementById('selectedTemplateName').textContent = item.name; },
    clearTemplate() { this.selectedTemplateFile = null; document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected')); document.getElementById('selectedTemplateInfo').classList.add('hidden'); },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(slideRelsXml) { if (!slideRelsXml) return null; const m = slideRelsXml.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    getPlaceholderRegexStr(ph) { const inner = ph.replace(/[\[\]]/g, '').trim(); const pts = inner.split(''); return '\\[' + '(?:<[^>]+>|\\s)*' + pts.map((p, i) => (p === ' ' ? '\\s+' : p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')) + (i < pts.length - 1 ? '(?:<[^>]+>|\\s)*' : '')).join('') + '(?:<[^>]+>|\\s)*' + '\\]'; },
    syncPresentationRegistry(newZip, presXml, presRelsXml, generated) {
        const sldIdLst = '<p:sldIdLst>' + generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        newZip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));
        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let relationships = relsDoc.getElementsByTagName('Relationship');
        for (let j = relationships.length - 1; j >= 0; j--) { if (relationships[j].getAttribute('Type').endsWith('slide')) relationships[j].parentNode.removeChild(relationships[j]); }
        generated.forEach(s => { let el = relsDoc.createElement('Relationship'); el.setAttribute('Id', s.rid); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); el.setAttribute('Target', `slides/${s.name}`); relsDoc.documentElement.appendChild(el); });
        newZip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));
        const ctXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>';
        let ctEntries = generated.map(s => `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
        newZip.file('[Content_Types].xml', (ctXml + ctEntries + '</Types>').replace('><Override', '><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'));
    },
    showLoading(text) { this.elements.loadingText.textContent = text; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; }
};

App.init();
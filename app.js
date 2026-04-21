/* LyricSlide Pro - Stable Version 2.1.6 */
const App = {
    version: "2.1.6",
    config: {
        VIRTUAL_WIDTH: 60, // Character width for Smart Align centering
        CHORD_SIZE: "1800", // 18pt
        AUTOFIT_TAG: '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/>'
    },
    elements: {
        songTitle: document.getElementById('songTitle'),
        lyricsInput: document.getElementById('lyricsInput'),
        copyrightInfo: document.getElementById('copyrightInfo'),
        generateBtn: document.getElementById('generateBtn'),
        alignmentSelect: document.getElementById('alignmentSelect'),
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

    init() {
        if (!this.elements.generateBtn) return; // Prevent init if DOM not ready

        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        
        this.elements.alignmentSelect.addEventListener('change', () => { 
            if (this.originalSlides.length > 0) this.updatePreview(0); 
        });

        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;

        const versionEl = document.getElementById('appVersion');
        if (versionEl) versionEl.textContent = this.version;
        console.log(`App Initialized. Version ${this.version}`);
    },

    // --- THEME MANAGEMENT ---
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
                    const map = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' };
                    this.setVariable(map[e.target.id], e.target.value);
                    this.save();
                });
            });
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
            if (confirm('Reset theme?')) { 
                Object.keys(this.defaults).forEach(key => { this.setVariable(key, this.defaults[key]); }); 
                this.save(); 
                location.reload();
            } 
        }
    },

    // --- CORE LOGIC ---
    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        const userAlign = this.elements.alignmentSelect.value;
        container.innerHTML = '';
        if (this.originalSlides.length === 0) return;

        this.originalSlides.forEach((slideData, idx) => {
            const wrapper = document.createElement('div'); wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div'); card.className = 'preview-card';
            const contentDiv = document.createElement('div'); contentDiv.className = 'slide-content'; 
            
            slideData.forEach(para => {
                if (para.text.trim() && !/©|Copyright|Words:|Music:|Lyrics:|Chris Tomlin|CCLI|DAYEG AMBASSADOR/i.test(para.text) && !para.isTitle) {
                    const lineDiv = document.createElement('div');
                    let displayText = this.transposeLine(para.text, semitones);
                    if (userAlign === 'smart') {
                        const offset = Math.max(0, Math.floor((this.config.VIRTUAL_WIDTH - para.text.trim().length) / 2));
                        displayText = "\u00A0".repeat(offset) + displayText.trimEnd();
                    } else { 
                        displayText = displayText.trimEnd(); 
                    }
                    lineDiv.style.minHeight = '1.2em';
                    lineDiv.innerHTML = this.renderChordHTML(displayText);
                    contentDiv.appendChild(lineDiv);
                }
            });
            if (contentDiv.children.length > 0) {
                card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black">Slide ${idx + 1}</div>`;
                card.appendChild(contentDiv); wrapper.appendChild(card); container.appendChild(wrapper);
            }
        });
    },

    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = this.elements.alignmentSelect.value;

        if (!file || !lyrics) return alert('Select template and enter lyrics.');

        try {
            this.showLoading('Generating PPTX...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const slideFileName = templateRelPath.split('/').pop();
            const relsPath = `ppt/slides/_rels/${slideFileName}.rels`;
            const templateRelsXml = zip.file(relsPath) ? await zip.file(relsPath).async('string') : null;
            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const sections = ("\n" + lyrics).split(/\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/).filter(s => s.trim() !== '');
            const newZip = zip;
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionText = sections[i].trim();
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);

                const name = `song_gen_${i + 1}.xml`;
                const path = `ppt/slides/${name}`;
                newZip.file(path, slideXml);
                
                if (templateNotesXml) {
                    const notesName = `notes_gen_${i + 1}.xml`;
                    const notesPath = `ppt/notesSlides/${notesName}`;
                    newZip.file(notesPath, templateNotesXml.replace(/\[Presenter Note\]/g, this.escXml(sectionText).replace(/\r?\n/g, '</a:t></a:r><a:br/><a:r><a:t xml:space="preserve">')));
                    newZip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml.replace(/Target="..\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/${notesName}"`));
                    newZip.file(`ppt/notesSlides/_rels/${notesName}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${name}"/></Relationships>`);
                } else if (templateRelsXml) {
                    newZip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                }
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path });
            }

            this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);
            const finalBlob = await newZip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${title.replace(/[^a-z0-9]/gi, '_') || 'Song'}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    lockInStyleAndReplace(xml, placeholder, replacement, userAlign = 'smart') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(placeholder), 'gi');
        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));
                const rawLines = (replacement || '').split(/\r?\n/);

                if (placeholder !== '[Lyrics and Chords]') {
                    return shapeXml.replace(phRegex, rawLines.map(l => this.escXml(l.trim())).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`));
                }

                let injectedXml = `</a:t></a:r></a:p>`;
                for (let i = 0; i < rawLines.length; i++) {
                    let line = rawLines[i], nextLine = rawLines[i + 1];
                    // Pair Detection (Chord + Lyric)
                    if (this.isChordLine(line) && nextLine !== undefined && !this.isChordLine(nextLine) && !nextLine.trim().startsWith('[')) {
                        let cText = line, lText = nextLine;
                        if (userAlign === 'smart') {
                            const offset = Math.max(0, Math.floor((this.config.VIRTUAL_WIDTH - lText.trim().length) / 2));
                            const pad = " ".repeat(offset);
                            cText = pad + cText.trimEnd(); lText = pad + lText.trimEnd();
                        } else { 
                            cText = cText.trimEnd(); lText = lText.trimEnd(); 
                        }
                        injectedXml += this.makeMixedStyleLine(cText, style, 'l') + this.makePptLine(lText, style, 'l'); 
                        i++;
                    } else if (line.trim() !== "") {
                        let text = line.trim();
                        if (userAlign === 'smart') text = " ".repeat(Math.max(0, Math.floor((this.config.VIRTUAL_WIDTH - text.length) / 2))) + text;
                        injectedXml += this.makePptLine(text, style, 'l');
                    } else {
                        injectedXml += `<a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                    }
                }
                injectedXml += `<a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r>${style}<a:t xml:space="preserve">`;
                let res = shapeXml.replace(phRegex, () => injectedXml).replace(/<a:p><a:pPr[^>]*><a:buNone\/><\/a:pPr><a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r><\/a:p>/g, '');
                return res.includes('Autofit') ? res : res.replace('</a:bodyPr>', this.config.AUTOFIT_TAG + '</a:bodyPr>');
            }
            return shapeXml;
        });
    },

    makeMixedStyleLine(text, lyricStyle, align) {
        const chordStyle = lyricStyle.includes('sz=') 
            ? lyricStyle.replace(/sz="\d+"/, `sz="${this.config.CHORD_SIZE}"`) 
            : lyricStyle.replace('<a:rPr', `<a:rPr sz="${this.config.CHORD_SIZE}"`);
        
        let runsXml = ""; 
        const segments = text.split(/(\s+)/);
        segments.forEach(seg => { 
            if (seg === "") return; 
            const activeStyle = /^\s+$/.test(seg) ? lyricStyle : chordStyle; 
            runsXml += `<a:r>${activeStyle}<a:t xml:space="preserve">${this.escXml(seg).replace(/ /g, '\u00A0')}</a:t></a:r>`; 
        });
        return `<a:p><a:pPr algn="${align}"><a:buNone/></a:pPr>${runsXml}</a:p>`;
    },

    makePptLine(text, style, align) { 
        return `<a:p><a:pPr algn="${align}"><a:buNone/></a:pPr><a:r>${style}<a:t xml:space="preserve">${this.escXml(text).replace(/ /g, '\u00A0')}</a:t></a:r></a:p>`; 
    },

    isChordLine(text) { 
        if (!text || text.trim() === '') return false; 
        const matches = text.match(/\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g) || []; 
        const words = text.trim().split(/\s+/).filter(w => w.length > 0); 
        return matches.length > 0 && (matches.length >= words.length * 0.4 || words.length < 3); 
    },

    // --- TRANSPOSE / UTILS ---
    async transpose() {
        const file = this.elements.transFileInput.files[0], semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file) return alert('Select file.');
        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file), slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');
                if (semitones !== 0) content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                zip.file(path, content);
            }
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, file.name.replace('.pptx', `_transposed.pptx`)); 
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        return text.split('\n').map(line => {
            const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
            if ((line.match(chordRegex) || []).length === 0) return line;
            let res = line, off = 0; const matches = [...line.matchAll(chordRegex)];
            for (const m of matches) {
                const pos = m.index + off, newC = this.shiftNote(m[1], semitones) + (m[2] || '') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '');
                const diff = newC.length - m[0].length; res = res.substring(0, pos) + newC + res.substring(pos + m[0].length);
                if (diff > 0) { 
                    let sm = res.substring(pos + newC.length).match(/^ +/); 
                    if (sm && sm[0].length >= diff) res = res.substring(0, pos + newC.length) + res.substring(pos + newC.length + diff); 
                    else off += diff; 
                } else if (diff < 0) res = res.substring(0, pos + newC.length) + " ".repeat(Math.abs(diff)) + res.substring(pos + newC.length);
            }
            return res;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys, idx = list.indexOf(note);
        if (idx === -1) { list = (list === this.musical.keys ? this.musical.flats : this.musical.keys); idx = list.indexOf(note); }
        return idx === -1 ? note : (semitones >= 0 ? this.musical.keys : this.musical.flats)[(idx + semitones + 12) % 12];
    },

    // --- XML HELPERS ---
    syncPresentationRegistry(newZip, presXml, presRelsXml, generated) {
        const sldIdLst = '<p:sldIdLst>' + generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        newZip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));
        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let relationships = relsDoc.getElementsByTagName('Relationship');
        for (let j = relationships.length - 1; j >= 0; j--) if (relationships[j].getAttribute('Type').endsWith('slide')) relationships[j].parentNode.removeChild(relationships[j]);
        generated.forEach(s => { 
            let el = relsDoc.createElement('Relationship'); el.setAttribute('Id', s.rid); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); el.setAttribute('Target', `slides/${s.name}`); relsDoc.documentElement.appendChild(el); 
        });
        newZip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));
        const ctEntries = generated.map(s => `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
        newZip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>${ctEntries}</Types>`);
    },

    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            const entries = names.map(name => ({ name, getFile: async () => { const r = await fetch(`./${encodeURIComponent(name)}`); const blob = await r.blob(); return new File([blob], name, { type: blob.type }); } }));
            this.renderTemplateGallery(entries);
            document.getElementById('dirName').textContent = `${names.length} templates available`;
        } catch (e) { gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Library Offline</div>`; }
    },

    renderTemplateGallery(entries) {
        const gallery = document.getElementById('templateGallery'); gallery.innerHTML = '';
        const grid = document.createElement('div'); grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div'); card.className = 'template-card';
            const img = document.createElement('img'); img.className = 'template-thumb'; img.src = entry.name.replace(/\.pptx$/i, '.png');
            img.addEventListener('error', () => { const ph = document.createElement('div'); ph.className = 'template-thumb-placeholder'; ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>'; img.replaceWith(ph); });
            const nameDiv = document.createElement('div'); nameDiv.className = 'template-card-name'; nameDiv.textContent = entry.name.replace(/\.pptx$/i, '');
            card.appendChild(img); card.appendChild(nameDiv);
            card.addEventListener('click', async () => { try { card.style.opacity = '0.6'; const file = await entry.getFile(); card.style.opacity = '1'; this.selectTemplate({ name: entry.name, file }, card); } catch (e) { alert(e.message); } });
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },

    selectTemplate(item, cardEl) { this.selectedTemplateFile = item.file; document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected')); cardEl.classList.add('selected'); document.getElementById('selectedTemplateInfo').classList.remove('hidden'); document.getElementById('selectedTemplateName').textContent = item.name; },
    clearTemplate() { this.selectedTemplateFile = null; document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected')); document.getElementById('selectedTemplateInfo').classList.add('hidden'); },
    showLoading(text) { this.elements.loadingText.textContent = text; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    getPlaceholderRegexStr(ph) { const pts = ph.replace(/[\[\]]/g, '').trim().split(''); return '\\[' + '(?:<[^>]+>|\\s)*' + pts.map((p, i) => (p === ' ' ? '\\s+' : p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')) + (i < pts.length - 1 ? '(?:<[^>]+>|\\s)*' : '')).join('') + '(?:<[^>]+>|\\s)*' + '\\]'; },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(relXml) { if (!relXml) return null; const m = relXml.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; }
};

App.init();
/* LyricSlide Pro - v22.0 (Ported Stable Build) */

const App = {
    version: "v22.0 (Stable)",

    elements: {
        songTitle: document.getElementById('songTitle'),
        lyricsInput: document.getElementById('lyricsInput'),
        copyrightInfo: document.getElementById('copyrightInfo'),
        generateBtn: document.getElementById('generateBtn'),
        transFileInput: document.getElementById('transFileInput'),
        transposeBtn: document.getElementById('transposeBtn'),
        semitoneDisplay: document.getElementById('semitoneDisplay'),
        loadingOverlay: document.getElementById('loadingOverlay'),
        loadingText: document.getElementById('loadingText'),
        navGen: document.getElementById('modeGen'),
        navTrans: document.getElementById('modeTrans'),
        viewGen: document.getElementById('viewGen'),
        viewTrans: document.getElementById('viewTrans')
    },

    musical: {
        keys: ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'],
        flats: ['C', 'Db', 'D', 'Eb', 'E', 'F', 'Gb', 'G', 'Ab', 'A', 'Bb', 'B']
    },

    selectedTemplateFile: null, 

    init() {
        const verEl = document.getElementById('appVersion');
        if (verEl) verEl.textContent = this.version;

        // Navigation
        if (this.elements.navGen) this.elements.navGen.onclick = () => this.setMode('gen');
        if (this.elements.navTrans) this.elements.navTrans.onclick = () => this.setMode('trans');

        // Actions
        if (this.elements.generateBtn) this.elements.generateBtn.onclick = () => this.generate();
        if (this.elements.transposeBtn) this.elements.transposeBtn.onclick = () => this.transpose();
        
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        console.log("App v22.0 Ported & Initialized.");
    },

    setMode(mode) {
        const isGen = mode === 'gen';
        if (this.elements.navGen) this.elements.navGen.classList.toggle('active', isGen);
        if (this.elements.navTrans) this.elements.navTrans.classList.toggle('active', !isGen);
        if (this.elements.viewGen) this.elements.viewGen.classList.toggle('hidden', !isGen);
        if (this.elements.viewTrans) this.elements.viewTrans.classList.toggle('hidden', isGen);
    },

    // --- THEME MANAGEMENT (v21 Style) ---
    theme: {
        defaults: { '--primary-color': '#334155', '--bg-start': '#f8fafc', '--bg-end': '#f8fafc', '--text-main': '#1e293b', '--card-accent': '#e2e8f0', '--preview-card-bg': '#ffffff', '--preview-chord-color': '#334155', '--preview-lyrics-color': '#1e293b' },
        init() {
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                document.documentElement.style.setProperty(key, val);
                const picker = document.getElementById('picker-' + key.replace('--', '').replace('-color', ''));
                if (picker) picker.value = val;
            });
            document.querySelectorAll('.color-picker-input').forEach(p => {
                p.oninput = (e) => {
                    const map = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' };
                    if(map[e.target.id]) document.documentElement.style.setProperty(map[e.target.id], e.target.value);
                    this.save();
                };
            });
        },
        save() {
            const current = {};
            Object.keys(this.defaults).forEach(k => { current[k] = getComputedStyle(document.documentElement).getPropertyValue(k).trim(); });
            localStorage.setItem('lyric_theme', JSON.stringify(current));
        }
    },

    // --- GENERATION ENGINE (v12 Ported) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const lyrics = this.elements.lyricsInput.value;
        const title = this.elements.songTitle.value || 'Song';
        if (!file || !lyrics) return alert('Select a template and enter lyrics.');

        try {
            this.showLoading('Generating PPTX...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[this.getSlideIds(presXml)[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templateRelPath.split('/').pop()}.rels`).async('string');
            
            // Ported Notes Logic
            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const sections = ("\n" + lyrics).split(/\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/).filter(s => s.trim() !== '');
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                let sectionText = sections[i].trim();
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', this.elements.copyrightInfo.value);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText);

                const name = `song_gen_${i + 1}.xml`;
                const path = `ppt/slides/${name}`;
                zip.file(path, slideXml);
                
                if (templateNotesXml) {
                    const notesName = `notes_gen_${i + 1}.xml`;
                    const notesPath = `ppt/notesSlides/${notesName}`;
                    const formattedNotes = this.escXml(sectionText).replace(/\r?\n/g, '</a:t></a:r><a:br/><a:r><a:t xml:space="preserve">');
                    zip.file(notesPath, templateNotesXml.replace(/\[Presenter Note\]/g, formattedNotes));
                    zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml.replace(/Target="..\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/${notesName}"`));
                    zip.file(`ppt/notesSlides/_rels/${notesName}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${name}"/></Relationships>`);
                    generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path, notesPath });
                } else {
                    zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                    generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path });
                }
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${title.replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (e) { console.error(e); this.hideLoading(); alert("Error: " + e.message); }
    },

    // --- COORDINATE INHERITANCE (v12) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(placeholder), 'gi');
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                // Inherit Box Size/Position from Template
                const x = (shapeXml.match(/<a:off x="(\d+)"/) || [0, "0"])[1];
                const y = (shapeXml.match(/<a:off [^>]*y="(\d+)"/) || [0, "1000000"])[1];
                const cx = (shapeXml.match(/<a:ext cx="(\d+)"/) || [0, "9144000"])[1];
                const font = (shapeXml.match(/<a:latin typeface="([^"]+)"/) || [0, "Arial"])[1];
                const size = (shapeXml.match(/sz="(\d+)"/) || [0, "2400"])[1];

                if (placeholder !== '[Lyrics and Chords]') {
                    const rPr = (shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g) || ['<a:rPr lang="en-US"/>'])[0];
                    const esc = this.escXml(replacement || '').split(/\n/).join(`</a:t></a:r><a:br/><a:r>${rPr}<a:t xml:space="preserve">`);
                    return shapeXml.replace(phRegex, esc);
                }

                // Table-based formatting for Lyrics
                const lines = (replacement || '').split(/\r?\n/);
                let rows = '';
                lines.forEach(line => {
                    const isTag = line.trim().startsWith('[') && line.trim().endsWith(']');
                    const isChord = line.match(chordRegex);
                    const typeface = (isChord || isTag) ? "Courier New" : font;
                    const fontSize = isTag ? Math.round(size * 0.8) : size;
                    const escLine = this.escXml(line).replace(/ /g, '&#160;');
                    
                    rows += `<a:tr h="450000"><a:tc><a:txBody><a:bodyPr vert="ctr" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="${fontSize}" lang="en-US"><a:latin typeface="${typeface}"/><a:cs typeface="${typeface}"/></a:r><a:t xml:space="preserve">${escLine}</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB><a:noFill/></a:tcPr></a:tc></a:tr>`;
                });

                return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="1025" name="Lyrics"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="5000000"/></p:xfm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="0" bandRow="0"><a:tableStyleId>{5C22544A-7EE6-4342-B051-7303C2061113}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="${cx}"/></a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
            }
            return shapeXml;
        });
    },

    // --- TEMPLATE LIB (v12 Feedback) ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json?v=' + Date.now());
            const names = await res.json();
            document.getElementById('dirName').textContent = `${names.length} templates available`;
            gallery.innerHTML = '';
            const grid = document.createElement('div');
            grid.className = 'template-grid';
            
            names.forEach(name => {
                const card = document.createElement('div');
                card.className = 'template-card';
                card.innerHTML = `<img class="template-thumb" src="${name.replace('.pptx','.png')}" onerror="this.src='data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><rect width=%22100%22 height=%22100%22 fill=%22%23eee%22/><text x=%2250%%22 y=%2250%%22 text-anchor=%22middle%22 dy=%22.3em%22 font-family=%22sans-serif%22 fill=%22%23999%22>PPTX</text></svg>'"><div class="template-card-name">${name.replace('.pptx','')}</div>`;
                
                card.onclick = async () => {
                    try {
                        card.style.opacity = '0.5'; // Visual feedback
                        const r = await fetch(`./${encodeURIComponent(name)}`);
                        const blob = await r.blob();
                        this.selectedTemplateFile = new File([blob], name, { type: blob.type });
                        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                        card.classList.add('selected');
                        card.style.opacity = '1';
                        // Update UI
                        const info = document.getElementById('selectedTemplateInfo');
                        if(info) info.classList.remove('hidden');
                        const infoName = document.getElementById('selectedTemplateName');
                        if(infoName) infoName.textContent = name;
                    } catch (err) { alert("Failed to load template."); card.style.opacity = '1'; }
                };
                grid.appendChild(card);
            });
            gallery.appendChild(grid);
        } catch (e) { gallery.innerHTML = 'Library error.'; }
    },

    // --- TRANSPOSITION (v12 logic) ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file) return alert('Select a PPTX file.');
        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.includes('ppt/slides/slide') || k.includes('ppt/notesSlides/notesSlide'));
            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');
                content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(this.unescXml(text), semitones)}</a:t>`);
                zip.file(path, content);
            }
            saveAs(await zip.generateAsync({ type: 'blob' }), file.name.replace('.pptx', `_transposed.pptx`));
            this.hideLoading();
        } catch (e) { this.hideLoading(); alert("Transposition failed."); }
    },

    transposeLine(text, semitones) {
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        return text.split('\n').map(line => {
            if (!line.match(chordRegex)) return line;
            let result = line, offset = 0;
            const matches = [...line.matchAll(chordRegex)];
            for (const m of matches) {
                const nc = this.shiftNote(m[1], semitones) + (m[2] || '') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '');
                result = result.substring(0, m.index + offset) + nc + result.substring(m.index + offset + m[0].length);
                offset += (nc.length - m[0].length);
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(note);
        if (idx === -1) idx = (list === this.musical.keys ? this.musical.flats : this.musical.keys).indexOf(note);
        if (idx === -1) return note;
        return (semitones >= 0 ? this.musical.keys : this.musical.flats)[(idx + semitones + 12) % 12];
    },

    // --- HELPERS ---
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    unescXml(s) { return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&apos;/g,"'"); },
    escXml(s) { return (s||'').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getPlaceholderRegexStr(ph) {
        const inner = ph.replace(/[\[\]]/g, '').trim();
        const fuzzy = inner.split('').map(c => c === ' ' ? '\\s*' : this.escRegex(c) + '(?:<[^>]+>)*').join('(?:<[^>]+>)*');
        return '\\[' + '(?:<[^>]+>|\\s)*' + fuzzy + '(?:<[^>]+>|\\s)*' + '\\]';
    },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(xml) { if(!xml) return null; const m = xml.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    
    syncPresentationRegistry(zip, pres, rels, gen) {
        const list = '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, list));
        let doc = new DOMParser().parseFromString(rels, 'application/xml');
        let rs = doc.getElementsByTagName('Relationship');
        for (let j = rs.length - 1; j >= 0; j--) if (rs[j].getAttribute('Type').endsWith('slide')) rs[j].parentNode.removeChild(rs[j]);
        gen.forEach(s => { let el = doc.createElement('Relationship'); el.setAttribute('Id', s.rid); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); el.setAttribute('Target', `slides/${s.name}`); doc.documentElement.appendChild(el); });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(doc));
        
        // Finalize Content Types (MANDATORY FOR PPTX VALIDITY)
        const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>';
        let entries = gen.map(s => {
            let row = `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
            if(s.notesPath) row += `<Override PartName="/${s.notesPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
            return row;
        }).join('');
        zip.file('[Content_Types].xml', (head + entries + '</Types>').replace('><Override', '><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'));
    }
};

function changeSemitones(delta) {
    const d = document.getElementById('semitoneDisplay');
    let n = Math.max(-11, Math.min(11, (parseInt(d.textContent) || 0) + delta));
    d.textContent = (n > 0 ? '+' : '') + n;
}

App.init();
/* LyricSlide Pro - Core Logic v15.5 (Auto-Load Root Templates, Background & Notes Fix) */

const App = {
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

    originalSlides: [],   // Slide data for live preview
    selectedTemplateFile: null, // Currently selected template File object

    init() {
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        
        this.theme.init();
        this.loadDefaultTemplates(); // Auto-loads without user selection
        window.LyricApp = this;
        console.log("App Initialized. Version 15.5 (Auto-Load & Style Preserved)");
    },

    // --- THEME MANAGEMENT (Original v15.2) ---
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

    // --- UI HELPERS (Original v15.2) ---
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
    },

    showLoading(text) {
        this.elements.loadingText.textContent = text;
        this.elements.loadingOverlay.style.display = 'flex';
    },

    hideLoading() {
        this.elements.loadingOverlay.style.display = 'none';
    },

    // --- TEMPLATE LIBRARY (Zero-Click Auto-Load) ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        const dirName = document.getElementById('dirName');
        const selectedInfo = document.getElementById('selectedTemplateInfo');
        const selectedName = document.getElementById('selectedTemplateName');

        // Strategy 1: Read templates.json and auto-select the first file
        try {
            const res = await fetch('./templates.json');
            if (res.ok) {
                const names = await res.json();
                if (names.length > 0) {
                    if (dirName) dirName.textContent = `${names.length} template(s) available`;
                    
                    const firstTemplate = names[0];
                    const r = await fetch(`./${encodeURIComponent(firstTemplate)}`);
                    const blob = await r.blob();
                    
                    this.selectedTemplateFile = new File([blob], firstTemplate, { type: blob.type });
                    
                    if (selectedInfo) selectedInfo.classList.remove('hidden');
                    if (selectedName) selectedName.textContent = firstTemplate;
                    if (gallery) gallery.innerHTML = `<div class="text-xs text-slate-500 text-center py-4">Auto-loaded: ${firstTemplate}</div>`;
                    
                    console.log("Successfully auto-loaded from templates.json:", firstTemplate);
                    return; 
                }
            }
        } catch (e) {
            console.log("Auto-load via templates.json failed.");
        }

        // Strategy 2: Directly fetch a default file named "template.pptx"
        try {
            const res = await fetch('./template.pptx');
            if (res.ok) {
                const blob = await res.blob();
                this.selectedTemplateFile = new File([blob], 'template.pptx', { type: blob.type });
                
                if (selectedInfo) selectedInfo.classList.remove('hidden');
                if (selectedName) selectedName.textContent = "template.pptx (Default)";
                if (gallery) gallery.innerHTML = `<div class="text-xs text-slate-500 text-center py-4">Auto-loaded: template.pptx</div>`;
                
                console.log("Successfully auto-loaded template.pptx");
                return;
            }
        } catch (e) {
            console.log("Auto-load via direct file name failed.");
        }

        // If both failed
        if (dirName) dirName.textContent = 'Could not load templates';
        if (gallery) {
            gallery.innerHTML = `<div class="text-center py-4 text-red-500 text-xs italic">
                No templates found. Please ensure 'templates.json' or 'template.pptx' exists in the same directory and you are running a local server.
            </div>`;
        }
    },

    // --- GENERATION LOGIC (Preserves Background & Writes Presenter Notes) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';

        if (!file) return alert('No template loaded. Please add a template to the folder or check console.');
        if (!lyrics) return alert('Lyrics are required.');

        try {
            this.showLoading('Reading template...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateSlidePath = `ppt/${templateRelPath}`;
            const templateXml = await zip.file(templateSlidePath).async('string');

            // FIX: Read the original relationship file to preserve background images
            const templateRelFileName = templateRelPath.split('/').pop();
            const templateRelsPath = `ppt/slides/_rels/${templateRelFileName}.rels`;
            let originalRelsXml = await zip.file(templateRelsPath).async('string');

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            if (sections.length === 0 && lyrics.trim() !== '') sections = [lyrics.trim()];
            
            const newZip = zip;
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionContent = sections[i].trim();
                let slideXml = templateXml;
                slideXml = this.lockInStyleAndReplace(slideXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionContent);

                const name = `song_gen_${i + 1}.xml`;
                const path = `ppt/slides/${name}`;
                newZip.file(path, slideXml);

                // Create Presenter Notes
                const noteName = `notesSlideGen${i + 1}.xml`;
                newZip.file(`ppt/notesSlides/${noteName}`, this.createNotesSlideXml(sectionContent));

                // RELS: Clone the original slide rels (keeping backgrounds) and append link to notes
                let relsDoc = new DOMParser().parseFromString(originalRelsXml, 'application/xml');
                let noteRel = relsDoc.createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship');
                noteRel.setAttribute('Id', 'rIdNotesUpdate'); 
                noteRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide');
                noteRel.setAttribute('Target', `../notesSlides/${noteName}`);
                relsDoc.documentElement.appendChild(noteRel);
                
                newZip.file(`ppt/slides/_rels/${name}.rels`, new XMLSerializer().serializeToString(relsDoc));

                // Notes Rel
                const noteRelXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/relationships">
                    <Relationship Id="rIdNotesMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMaster/notesMaster1.xml"/>
                </Relationships>`;
                newZip.file(`ppt/notesSlides/_rels/${noteName}.rels`, noteRelXml);

                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path, noteName });
            }

            await this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);

            this.showLoading('Downloading...');
            const finalBlob = await newZip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error: " + err.message);
            this.hideLoading();
        }
    },

    async syncPresentationRegistry(newZip, presXml, presRelsXml, generated) {
        const sldIdLst = '<p:sldIdLst>' + generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        newZip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));

        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let relationships = relsDoc.getElementsByTagName('Relationship');
        for (let j = relationships.length - 1; j >= 0; j--) {
            if (relationships[j].getAttribute('Type').endsWith('slide')) relationships[j].parentNode.removeChild(relationships[j]);
        }
        generated.forEach(s => {
            let el = relsDoc.createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship');
            el.setAttribute('Id', s.rid);
            el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            el.setAttribute('Target', `slides/${s.name}`);
            relsDoc.documentElement.appendChild(el);
        });
        newZip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));

        const ctFile = await newZip.file('[Content_Types].xml').async('string');
        let ctDoc = new DOMParser().parseFromString(ctFile, 'application/xml');
        
        generated.forEach(s => {
            let sld = ctDoc.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Override');
            sld.setAttribute('PartName', `/${s.path}`);
            sld.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            ctDoc.documentElement.appendChild(sld);

            let nte = ctDoc.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Override');
            nte.setAttribute('PartName', `/ppt/notesSlides/${s.noteName}`);
            nte.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
            ctDoc.documentElement.appendChild(nte);
        });
        newZip.file('[Content_Types].xml', new XMLSerializer().serializeToString(ctDoc));
    },

    // --- TRANSPOSITION LOGIC (Original v15.2) ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file) return alert('Select a PPTX file to transpose.');
        try {
            this.showLoading('Applying changes...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');
                if (semitones !== 0) {
                    content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                }
                zip.file(path, content);
            }
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, file.name.replace('.pptx', `_modified.pptx`));
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        return text.split('\n').map(line => {
            const matches = [...line.matchAll(chordRegex)];
            if (matches.length === 0) return line;
            let result = line, offset = 0;
            for (const m of matches) {
                const pos = m.index + offset;
                const newRoot = this.shiftNote(m[1], semitones);
                const newBass = m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '';
                const newChord = newRoot + (m[2] || '') + newBass;
                result = result.substring(0, pos) + newChord + result.substring(pos + m[0].length);
                offset += (newChord.length - m[0].length);
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(note);
        if (idx === -1) idx = (list === this.musical.keys ? this.musical.flats : this.musical.keys).indexOf(note);
        if (idx === -1) return note;
        const outList = semitones >= 0 ? this.musical.keys : this.musical.flats;
        return outList[(idx + semitones + 12) % 12];
    },

    // --- TEXT INJECTION LOGIC (Keeps your v15.2 layout style exactly) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));
                const rawLines = (replacement || '').split(/\r?\n/);

                if (placeholder !== '[Lyrics and Chords]') {
                    const escapedText = rawLines.map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    return shapeXml.replace(phRegex, escapedText);
                }

                let injectedXml = `</a:t></a:r></a:p>`;
                rawLines.forEach((line) => {
                    const trimmed = line.trim();
                    if (trimmed === '') {
                        injectedXml += `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                        return;
                    }
                    const isTag = trimmed.startsWith('[') && trimmed.endsWith(']');
                    const chords = line.match(chordRegex) || [];
                    const words = trimmed.split(/\s+/).filter(w => w.length > 0);
                    const isChordLine = chords.length > 0 && !isTag && (chords.length >= words.length * 0.3 || words.length < 3);
                    let alignment = isChordLine ? 'l' : 'ctr';
                    let lineStyle = isChordLine ? style.replace(/sz="\d+"/, 'sz="1800"').replace('<a:rPr', '<a:rPr sz="1800"') : style;
                    injectedXml += `<a:p><a:pPr algn="${alignment}"><a:buNone/></a:pPr><a:r>${lineStyle}<a:t xml:space="preserve">${this.escXml(line).replace(/ /g, '\u00A0')}</a:t></a:r></a:p>`;
                });
                return shapeXml.replace(phRegex, () => injectedXml + `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r>${style}<a:t xml:space="preserve">`);
            }
            return shapeXml;
        });
    },

    createNotesSlideXml(text) {
        const lines = text.split(/\r?\n/).map(line => 
            `<a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>${this.escXml(line)}</a:t></a:r></a:p>`
        ).join('');
        return `<?xml version="1.0" encoding="UTF-8"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${lines}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
    },

    getPlaceholderRegexStr(ph) { 
        const pts = ph.replace(/[\[\]]/g, '').trim().split('');
        return '\\[' + pts.map(p => this.escRegex(p)).join('(?:<[^>]+>|\\s)*') + '\\]'; 
    },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; }
};

App.init();
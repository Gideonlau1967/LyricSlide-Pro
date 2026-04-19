/* LyricSlide Pro - Core Logic v15.3 (Full Integrated Version) */

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

    originalSlides: [],   
    selectedTemplateFile: null, 

    init() {
        if (!this.elements.generateBtn) {
            console.error("UI Elements not found. Check your HTML IDs.");
            return;
        }
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        console.log("App Initialized. Version 15.3 (Background & Notes Fix)");
    },

    // --- THEME MANAGEMENT ---
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

    // --- UI HELPERS ---
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

    // --- TEMPLATE LIBRARY ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        const dirName = document.getElementById('dirName');
        try {
            const res = await fetch('./templates.json');
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const names = await res.json();
            if(dirName) dirName.textContent = `${names.length} template${names.length !== 1 ? 's' : ''} available`;
            const entries = names.map(name => ({
                name,
                getFile: async () => {
                    const r = await fetch(`./${encodeURIComponent(name)}`);
                    const blob = await r.blob();
                    return new File([blob], name, { type: blob.type });
                }
            }));
            this.renderTemplateGallery(entries);
        } catch (e) {
            if(dirName) dirName.textContent = 'Could not load templates';
            gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Templates failed to load. Ensure you are using a local server (e.g. Live Server).</div>`;
        }
    },

    renderTemplateGallery(entries) {
        const gallery = document.getElementById('templateGallery');
        gallery.innerHTML = '';
        const grid = document.createElement('div');
        grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div');
            card.className = 'template-card';
            const thumbSrc = entry.name.replace(/\.pptx$/i, '.png');
            const img = document.createElement('img');
            img.className = 'template-thumb';
            img.src = thumbSrc;
            img.addEventListener('error', () => {
                const ph = document.createElement('div');
                ph.className = 'template-thumb-placeholder';
                ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>';
                img.replaceWith(ph);
            });
            const nameDiv = document.createElement('div');
            nameDiv.className = 'template-card-name';
            nameDiv.textContent = entry.name.replace(/\.pptx$/i, '');
            card.appendChild(img);
            card.appendChild(nameDiv);
            card.addEventListener('click', async () => {
                this.selectedTemplateFile = await entry.getFile();
                document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                card.classList.add('selected');
                document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                document.getElementById('selectedTemplateName').textContent = entry.name;
            });
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },

    // --- GENERATION LOGIC ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';

        if (!file) return alert('Please select a template first.');
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

                const noteName = `notesSlideGen${i + 1}.xml`;
                newZip.file(`ppt/notesSlides/${noteName}`, this.createNotesSlideXml(sectionContent));

                let relsDoc = new DOMParser().parseFromString(originalRelsXml, 'application/xml');
                let noteRel = relsDoc.createElement('Relationship');
                noteRel.setAttribute('Id', 'rIdNotesCustom'); 
                noteRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide');
                noteRel.setAttribute('Target', `../notesSlides/${noteName}`);
                relsDoc.documentElement.appendChild(noteRel);
                
                newZip.file(`ppt/slides/_rels/${name}.rels`, new XMLSerializer().serializeToString(relsDoc));

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
            let el = relsDoc.createElement('Relationship');
            el.setAttribute('Id', s.rid);
            el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            el.setAttribute('Target', `slides/${s.name}`);
            relsDoc.documentElement.appendChild(el);
        });
        newZip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));

        const ctFile = await newZip.file('[Content_Types].xml').async('string');
        let ctDoc = new DOMParser().parseFromString(ctFile, 'application/xml');
        let types = ctDoc.documentElement;
        let overrides = types.getElementsByTagName('Override');
        for (let i = overrides.length - 1; i >= 0; i--) {
            const pn = overrides[i].getAttribute('PartName');
            if (pn.includes('/ppt/slides/song_gen') || pn.includes('/ppt/notesSlides/notesSlideGen')) {
                overrides[i].parentNode.removeChild(overrides[i]);
            }
        }
        generated.forEach(s => {
            let sld = ctDoc.createElement('Override');
            sld.setAttribute('PartName', `/${s.path}`);
            sld.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            types.appendChild(sld);
            let nte = ctDoc.createElement('Override');
            nte.setAttribute('PartName', `/ppt/notesSlides/${s.noteName}`);
            nte.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
            types.appendChild(nte);
        });
        newZip.file('[Content_Types].xml', new XMLSerializer().serializeToString(ctDoc));
    },

    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file) return alert('Select a PPTX file.');
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
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11
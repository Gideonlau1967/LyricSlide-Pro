/* LyricSlide Pro - Core Logic v18.1 (Version-Stamped & Table-Safe) */

const App = {
    version: "v18.1", // The version that will appear in your PPTX

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
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
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
                const p = document.getElementById('picker-' + key.replace('--', '').replace('-color', ''));
                if (p) p.value = val;
            });
            document.querySelectorAll('.color-picker-input').forEach(p => p.addEventListener('input', (e) => { App.theme.setVariable(App.theme.getVarNameFromPicker(e.target.id), e.target.value); App.theme.save(); }));
        },
        getVarNameFromPicker(id) { const map = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' }; return map[id]; },
        setVariable(name, val) { document.documentElement.style.setProperty(name, val); },
        save() { const c = {}; Object.keys(this.defaults).forEach(k => { c[k] = getComputedStyle(document.documentElement).getPropertyValue(k).trim(); }); localStorage.setItem('lyric_theme', JSON.stringify(c)); }
    },

    // --- PREVIEW ENGINE (Reads Tables) ---
    async loadForPreview(file) {
        try {
            this.showLoading('Reading slides...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml')).sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            this.originalSlides = [];
            let globalSongTitle = "";

            for (const path of slideFiles) {
                const xml = await zip.file(path).async('string');
                const slideData = [];
                const elRegex = /<(p:sp|a:tc)>([\s\S]*?)<\/\1>/g;
                let m;
                while ((m = elRegex.exec(xml)) !== null) {
                    const content = m[2];
                    const ph = content.match(/<p:ph[^>]*type="(?:title|ctrTitle|ftr|dt|sldNum)"/);
                    const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g;
                    let pm;
                    while ((pm = pRegex.exec(content)) !== null) {
                        const tags = /<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g;
                        let text = '', tm;
                        while ((tm = tags.exec(pm[1])) !== null) { text += tm[0].startsWith('<a:br') ? '\n' : this.unescXml(tm[2] || ''); }
                        let align = pm[1].includes('algn="ctr"') ? 'center' : 'left';
                        if (ph && (ph[0].includes('title')) && text.trim() && !globalSongTitle) globalSongTitle = text.trim();
                        slideData.push({ text, alignment: align, isTitle: !!ph });
                    }
                }
                this.originalSlides.push(slideData);
            }
            this.songTitle = globalSongTitle;
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides Loaded`;
            this.updatePreview(0);
            this.hideLoading();
        } catch (e) { this.hideLoading(); alert("Error loading preview"); }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';
        if (this.originalSlides.length === 0) return;
        this.originalSlides.forEach((slide, idx) => {
            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 font-black">Slide ${idx + 1}</div>`;
            const content = document.createElement('div');
            content.className = 'slide-content';
            slide.forEach(p => {
                if (p.text.trim() && !p.isTitle && !/©|Copyright|CCLI/i.test(p.text)) {
                    const d = document.createElement('div');
                    d.style.textAlign = p.alignment;
                    d.innerHTML = this.renderChordHTML(this.transposeLine(p.text, semitones));
                    content.appendChild(d);
                }
            });
            if (content.children.length > 0) { card.appendChild(content); container.appendChild(card); }
        });
        this.updateZoom();
    },

    updateZoom(val) {
        if (val === undefined) val = document.getElementById('zoomSlider').value;
        document.getElementById('zoomVal').textContent = val + '%';
        const contents = document.getElementsByClassName('slide-content');
        for(let c of contents) c.style.transform = `scale(${val / 100})`;
    },

    // --- GENERATION ENGINE ---
    async generate() {
        const file = this.selectedTemplateFile;
        const lyrics = this.elements.lyricsInput.value;
        if (!file || !lyrics) return alert('Select template and enter lyrics.');

        try {
            this.showLoading('Generating PPTX...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[this.getSlideIds(presXml)[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const slideFileName = templateRelPath.split('/').pop();
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${slideFileName}.rels`).async('string');
            
            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            if (sections.length === 0 && lyrics.trim() !== '') sections = [lyrics.trim()];
            
            const generated = [];

            // Add the version stamp to the copyright info
            const copyrightWithVersion = (this.elements.copyrightInfo.value ? this.elements.copyrightInfo.value + " | " : "") + "Generated by LyricSlide Pro " + this.version;

            for (let i = 0; i < sections.length; i++) {
                let slideXml = templateXml;
                slideXml = this.lockInStyleAndReplace(slideXml, '[Title]', this.elements.songTitle.value);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyrightWithVersion);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sections[i].trim());

                const name = `song_gen_${i + 1}.xml`;
                zip.file(`ppt/slides/${name}`, slideXml);
                
                let notesPath = null;
                if (templateNotesXml) {
                    notesPath = `ppt/notesSlides/notes_gen_${i + 1}.xml`;
                    const formattedNotes = this.escXml(sections[i].trim()).replace(/\r?\n/g, '</a:t></a:r><a:br/><a:r><a:t xml:space="preserve">');
                    zip.file(notesPath, templateNotesXml.replace(/\[Presenter Note\]/g, formattedNotes));
                    zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml.replace(/Target="..\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/notes_gen_${i+1}.xml"`));
                    zip.file(`ppt/notesSlides/_rels/notes_gen_${i+1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${name}"/></Relationships>`);
                } else {
                    zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                }
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path: `ppt/slides/${name}`, notesPath });
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${(this.elements.songTitle.value || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (e) { console.error(e); this.hideLoading(); alert("Error generating file."); }
    },

    // --- COORDINATE-SAFE TABLE FILLER (v18.1) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const createFuzzyRegex = (ph) => {
            const inner = ph.replace(/[\[\]]/g, '').trim();
            const fuzzy = inner.split('').map(char => 
                char === ' ' ? '\\s+' : `${this.escRegex(char)}(?:<[^>]+>)*`
            ).join('(?:<[^>]+>)*');
            return new RegExp('\\[' + '(?:<[^>]+>)*' + fuzzy + '(?:<[^>]+>)*' + '\\]', 'gi');
        };

        const phRegex = createFuzzyRegex(placeholder);
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        return xml.replace(/<(p:sp|p:graphicFrame)>([\s\S]*?)<\/\1>/g, (fullFrame) => {
            phRegex.lastIndex = 0;
            if (phRegex.test(fullFrame)) {
                const latinMatch = fullFrame.match(/<a:latin typeface="([^"]+)"/);
                const templateFont = latinMatch ? latinMatch[1] : "Arial";
                const sizeMatch = fullFrame.match(/sz="(\d+)"/);
                const templateSize = sizeMatch ? sizeMatch[1] : "2400"; 

                if (!/Lyrics/i.test(placeholder)) {
                    const rPrMatch = fullFrame.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                    let style = (rPrMatch ? rPrMatch[0] : '<a:rPr lang="en-US"/>');
                    const escapedText = (replacement || '').split(/\r?\n/).map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    return fullFrame.replace(phRegex, escapedText);
                }

                const lines = (replacement || '').split(/\r?\n/);
                let tableRowsXml = '';

                lines.forEach((line) => {
                    let trimmed = line.trim();
                    if (trimmed === '') {
                        tableRowsXml += this.createTableCellXml(" ", templateFont, templateSize, "ctr", 150000);
                        return;
                    }
                    const isTag = trimmed.startsWith('[') && trimmed.endsWith(']');
                    const hasChords = line.match(chordRegex);

                    if (isTag) {
                        tableRowsXml += this.createTableCellXml(trimmed, templateFont, Math.round(templateSize * 0.8), "ctr", 400000);
                    } else if (hasChords) {
                        const esc = this.escXml(line).replace(/ /g, '&#160;');
                        tableRowsXml += this.createTableCellXml(esc, "Courier New", templateSize, "l", 400000);
                    } else {
                        tableRowsXml += this.createTableCellXml(this.escXml(line), templateFont, templateSize, "ctr", 450000);
                    }
                });

                // COORDINATE SAFE: Splicing rows into existing shell
                const xmlParts = fullFrame.split(/<a:tr[\s\S]*?<\/a:tr>/);
                const header = xmlParts[0];
                const footer = xmlParts[xmlParts.length - 1];

                return `<p:graphicFrame>${header}${tableRowsXml}${footer}</p:graphicFrame>`;
            }
            return fullFrame;
        });
    },

    createTableCellXml(text, font, size, align, height) {
        return `<a:tr h="${height}"><a:tc><a:txBody><a:bodyPr vert="ctr" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/><a:p><a:pPr algn="${align}"/><a:r><a:rPr sz="${size}" lang="en-US"><a:latin typeface="${font}"/><a:cs typeface="${font}"/></a:rPr><a:t xml:space="preserve">${text}</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB><a:solidFill><a:noFill/></a:solidFill></a:tcPr></a:tc></a:tr>`;
    },

    // --- TEMPLATE LIB ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            document.getElementById('dirName').textContent = `${names.length} templates available`;
            const grid = document.createElement('div'); grid.className = 'template-grid';
            names.forEach(name => {
                const card = document.createElement('div'); card.className = 'template-card';
                card.innerHTML = `<img class="template-thumb" src="${name.replace('.pptx','.png')}" onerror="this.src='data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%2
/* LyricSlide Pro - Core Logic v19.1 (Visible Table Fix) */

const App = {
    version: "v19.1 (Visible Table)",

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
        const verEl = document.getElementById('appVersion');
        if (verEl) verEl.textContent = this.version;

        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        console.log(`App Initialized. ${this.version}`);
    },

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
            document.querySelectorAll('.color-picker-input').forEach(picker => {
                picker.addEventListener('input', (e) => {
                    const id = e.target.id;
                    const map = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' };
                    document.documentElement.style.setProperty(map[id], e.target.value);
                });
            });
        }
    },

    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
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
                    const r = await fetch(`./${encodeURIComponent(name)}`);
                    const blob = await r.blob();
                    this.selectedTemplateFile = new File([blob], name, { type: blob.type });
                    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                    card.classList.add('selected');
                    document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                    document.getElementById('selectedTemplateName').textContent = name;
                };
                grid.appendChild(card);
            });
            gallery.appendChild(grid);
        } catch (e) { gallery.innerHTML = 'Library load failed.'; }
    },

    async generate() {
        if (!this.selectedTemplateFile || !this.elements.lyricsInput.value) return alert('Select template and enter lyrics.');
        try {
            this.showLoading('Generating...');
            const zip = await JSZip.loadAsync(this.selectedTemplateFile);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[this.getSlideIds(presXml)[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templateRelPath.split('/').pop()}.rels`).async('string');
            
            const sections = ("\n" + this.elements.lyricsInput.value).split(/\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/).filter(s => s.trim() !== '');
            const generated = [];
            const copyrightText = (this.elements.copyrightInfo.value ? this.elements.copyrightInfo.value + " | " : "") + "Generated by " + this.version;

            for (let i = 0; i < sections.length; i++) {
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', this.elements.songTitle.value);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyrightText);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sections[i].trim());

                const name = `song_gen_${i + 1}.xml`;
                zip.file(`ppt/slides/${name}`, slideXml);
                zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path: `ppt/slides/${name}` });
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            saveAs(await zip.generateAsync({ type: 'blob' }), `${(this.elements.songTitle.value || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (e) { this.hideLoading(); alert("Error during generation."); }
    },

    // --- RE-BUILT REPLACEMENT ENGINE (Fixes the "Invisible Table" Bug) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        // Build regex that ignores case AND ignores internal XML tags between letters
        const createFuzzyRegex = (ph) => {
            const inner = ph.replace(/[\[\]]/g, '').trim();
            // Handle optional spaces around the word (to match "[ TITLE ]" or "[Title]")
            const fuzzy = inner.split('').map(c => 
                c === ' ' ? '\\s*' : `${this.escRegex(c)}(?:<[^>]+>)*`
            ).join('(?:<[^>]+>)*');
            return new RegExp('\\[' + '(?:<[^>]+>|\\s)*' + fuzzy + '(?:<[^>]+>|\\s)*' + '\\]', 'gi');
        };

        const phRegex = createFuzzyRegex(placeholder);
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        return xml.replace(/<(p:sp|p:graphicFrame)>([\s\S]*?)<\/\1>/g, (fullFrame, tagName, innerContent) => {
            phRegex.lastIndex = 0;
            if (phRegex.test(innerContent)) {
                
                const latinMatch = innerContent.match(/<a:latin typeface="([^"]+)"/);
                const templateFont = latinMatch ? latinMatch[1] : "Arial";
                const sizeMatch = innerContent.match(/sz="(\d+)"/);
                const templateSize = sizeMatch ? sizeMatch[1] : "2400"; 

                // Non-lyrics (Title/Copyright)
                if (!/Lyrics/i.test(placeholder)) {
                    const rPrMatch = innerContent.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                    let style = (rPrMatch ? rPrMatch[0] : '<a:rPr lang="en-US"/>');
                    const escapedText = (replacement || '').split(/\r?\n/).map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    return `<${tagName}>${innerContent.replace(phRegex, escapedText)}</${tagName}>`;
                }

                // Table Logic
                const lines = (replacement || '').split(/\r?\n/);
                let tableRowsXml = '';
                lines.forEach((line) => {
                    let trimmed = line.trim();
                    if (trimmed === '') { tableRowsXml += this.createTableCellXml(" ", templateFont, templateSize, "ctr", 150000); return; }
                    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
                        tableRowsXml += this.createTableCellXml(trimmed, templateFont, Math.round(templateSize * 0.8), "ctr", 400000);
                    } else if (line.match(chordRegex)) {
                        tableRowsXml += this.createTableCellXml(this.escXml(line).replace(/ /g, '&#160;'), "Courier New", templateSize, "l", 400000);
                    } else {
                        tableRowsXml += this.createTableCellXml(this.escXml(line), templateFont, templateSize, "ctr", 450000);
                    }
                });

                // FIXED: Split the content ONLY, and re-wrap in the existing tagName to avoid double tags
                const rowsSplit = innerContent.split(/<a:tr[\s\S]*?<\/a:tr>/);
                const newInner = `${rowsSplit[0]}${tableRowsXml}${rowsSplit[rowsSplit.length - 1]}`;
                return `<${tagName}>${newInner}</${tagName}>`;
            }
            return fullFrame;
        });
    },

    createTableCellXml(text, font, size, align, height) {
        return `<a:tr h="${height}"><a:tc><a:txBody><a:bodyPr vert="ctr" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/><a:p><a:pPr algn="${align}"/><a:r><a:rPr sz="${size}" lang="en-US"><a:latin typeface="${font}"/><a:cs typeface="${font}"/></a:rPr><a:t xml:space="preserve">${text}</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB><a:solidFill><a:noFill/></a:solidFill></a:tcPr></a:tc></a:tr>`;
    },

    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s||'').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    syncPresentationRegistry(zip, pres, rels, gen) {
        const list = '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, list));
        let doc = new DOMParser().parseFromString(rels, 'application/xml');
        let rs = doc.getElementsByTagName('Relationship');
        for (let j = rs.length - 1; j >= 0; j--) if (rs[j].getAttribute('Type').endsWith('slide')) rs[j].parentNode.removeChild(rs[j]);
        gen.forEach(s => { let el = doc.createElement('Relationship'); el.setAttribute('Id', s.rid); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); el.setAttribute('Target', `slides/${s.name}`); doc.documentElement.appendChild(el); });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(doc));
        const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>';
        let entries = gen.map(s => `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
        zip.file('[Content_Types].xml', (head + entries + '</Types>').replace('><Override', '><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'));
    },
    async transpose() { /* Placeholder */ },
    transposeLine(t, s) { return t; }
};

App.init();
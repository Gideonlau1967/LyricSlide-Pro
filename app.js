/* LyricSlide Pro - Version 3.3.0 (Hybrid Inheritance Engine) */

const App = {
    version: "Version 3.3.0",
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
        if (alignSelect) alignSelect.addEventListener('change', () => { if (this.originalSlides.length > 0) this.updatePreview(0); });
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
    },

    // --- MS COMPLIANCE REGISTRY (WPS STYLE) ---
    async rebuildRegistry(zip, slideCount, hasNotes) {
        // 1. Rewrite Content Types
        let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
        ct += `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`;
        ct += `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`;
        ct += `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`;
        ct += `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`;
        ct += `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`;
        for (let i = 1; i <= slideCount; i++) {
            ct += `<Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
            if (hasNotes) ct += `<Override PartName="/ppt/notesSlides/notesSlide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
        }
        ct += `</Types>`;
        zip.file('[Content_Types].xml', ct);

        // 2. Rewrite Presentation Rels
        let pr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
        pr += `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`;
        pr += `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`;
        for (let i = 1; i <= slideCount; i++) pr += `<Relationship Id="s${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i}.xml"/>`;
        pr += `</Relationships>`;
        zip.file('ppt/_rels/presentation.xml.rels', pr);

        // 3. Rewrite Presentation.xml
        let pres = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>`;
        for (let i = 1; i <= slideCount; i++) pres += `<p:sldId id="${255+i}" r:id="s${i}"/>`;
        pres += `</p:sldIdLst><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
        zip.file('ppt/presentation.xml', pres);
    },

    // --- CORE GENERATION (WITH STYLE INHERITANCE) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Select template and input lyrics.');

        try {
            this.showLoading('Extracting Styles & Generating...');
            const zip = await JSZip.loadAsync(file);
            
            // 1. Find the blueprint slide
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const slideIdMatch = presXml.match(/<p:sldId[^>]+r:id="([^"]+)"/);
            const firstSlideRid = slideIdMatch ? slideIdMatch[1] : 'rId1';
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slidePathMatch = presRelsXml.match(new RegExp(`Id="${firstSlideRid}"[^>]+Target="([^"]+)"`));
            const templatePath = slidePathMatch ? `ppt/${slidePathMatch[1]}` : 'ppt/slides/slide1.xml';
            
            const templateXml = await zip.file(templatePath).async('string');

            // 2. Clean existing slides
            Object.keys(zip.files).forEach(f => {
                if (f.startsWith('ppt/slides/slide') || f.startsWith('ppt/notesSlides/notesSlide')) zip.remove(f);
            });

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');

            for (let i = 0; i < sections.length; i++) {
                const num = i + 1;
                const sectionText = sections[i].trim();
                
                // INHERITANCE: Replace text while keeping the exact <p:sp> (Shape) from the template
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);

                zip.file(`ppt/slides/slide${num}.xml`, slideXml);
                
                // Rels
                zip.file(`ppt/slides/_rels/slide${num}.xml.rels`, `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${num}.xml"/></Relationships>`);
                
                // Notes
                const notesXml = `<?xml version="1.0" encoding="UTF-8"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">${this.escXml(sectionText)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
                zip.file(`ppt/notesSlides/notesSlide${num}.xml`, notesXml);
                zip.file(`ppt/notesSlides/_rels/notesSlide${num}.xml.rels`, `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${num}.xml"/></Relationships>`);
            }

            await this.rebuildRegistry(zip, sections.length, true);
            saveAs(await zip.generateAsync({ type: 'blob' }), `${title.replace(/[^a-z0-9]/gi, '_') || 'Song'}.pptx`);
            this.hideLoading();
        } catch (err) { console.error(err); alert(err.message); this.hideLoading(); }
    },

    // REPLACEMENT: Replaces the entire <a:p> block to keep formatting intact
    lockInStyleAndReplace(xml, ph, replacement, align = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(ph), 'gi');
        return xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (pMatch) => {
            if (!phRegex.test(pMatch)) return pMatch;

            // Inherit formatting from template
            const style = pMatch.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr lang="en-US"/>';
            const alignment = align === 'ctr' ? 'ctr' : 'l';

            if (ph !== '[Lyrics and Chords]') {
                return `<a:p><a:pPr algn="${alignment}"/><a:r>${style}<a:t xml:space="preserve">${this.escXml(replacement)}</a:t></a:r></a:p>`;
            }

            // Multi-line injection for Lyrics
            let newParagraphs = "";
            const lines = replacement.split('\n');
            for (let i = 0; i < lines.length; i++) {
                let line = lines[i], next = lines[i+1];
                if (this.isChordLine(line) && next && !this.isChordLine(next) && !next.trim().startsWith('[')) {
                    newParagraphs += this.makePptLine(line, this.getChordStyle(style), alignment) + this.makePptLine(next, style, alignment);
                    i++;
                } else {
                    newParagraphs += this.makePptLine(line, line.trim().startsWith('[') ? style.replace(/sz="\d+"/, 'sz="2000"') : style, alignment);
                }
            }
            return newParagraphs;
        });
    },

    makePptLine(text, style, align) {
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="80000"/></a:lnSpc></a:pPr><a:r>${style}<a:t xml:space="preserve">${this.escXml(text)}</a:t></a:r></a:p>`;
    },

    // (Include all original helper functions below: escXml, isChordLine, getChordStyle, getPlaceholderRegexStr, theme, loadDefaultTemplates, showLoading, etc.)
    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    isChordLine(lineStr) {
        if (!lineStr || typeof lineStr !== 'string') return false;
        const trimmed = lineStr.trim();
        if (trimmed === '' || /^(A|I|The|And|Then|They|We|He|She)\s+[a-zA-Z]{2,}/i.test(trimmed)) return false;
        return (trimmed.match(this.chordRegex) || []).length > 0;
    },
    getChordStyle(lyricStyle) {
        let s = lyricStyle.includes('sz=') ? lyricStyle.replace(/sz="\d+"/, 'sz="1800"') : lyricStyle.replace('<a:rPr', '<a:rPr sz="1800"');
        const greyFill = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>';
        return s.includes('<a:solidFill>') ? s.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, greyFill) : s.replace('</a:rPr>', greyFill + '</a:rPr>');
    },
    theme: {
        defaults: {'--primary-color': '#334155', '--bg-start': '#f8fafc', '--bg-end': '#f8fafc', '--text-main': '#1e293b', '--card-accent': '#e2e8f0', '--preview-card-bg': '#ffffff', '--preview-chord-color': '#334155', '--preview-lyrics-color': '#1e293b'},
        init() {
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                document.documentElement.style.setProperty(key, val);
            });
        }
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
        const gallery = document.getElementById('templateGallery');
        gallery.innerHTML = '';
        const grid = document.createElement('div');
        grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div');
            card.className = 'template-card';
            const nameDiv = document.createElement('div');
            nameDiv.className = 'template-card-name'; nameDiv.textContent = entry.name.replace(/\.pptx$/i, '');
            card.appendChild(nameDiv);
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
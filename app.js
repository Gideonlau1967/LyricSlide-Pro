/* LyricSlide Pro - Core Logic v16 (Table-Template Integration) */

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
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        console.log("App Initialized. Version 16.0 (Table Engine)");
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
            const map = {
                'picker-primary': '--primary-color',
                'picker-bg-start': '--bg-start',
                'picker-bg-end': '--bg-end',
                'picker-text': '--text-main',
                'picker-card-accent': '--card-accent',
                'picker-preview-bg': '--preview-card-bg',
                'picker-chord': '--preview-chord-color',
                'picker-lyrics': '--preview-lyrics-color'
            };
            return map[id];
        },

        setVariable(name, val) {
            document.documentElement.style.setProperty(name, val);
            if (name === '--primary-color') {
                document.documentElement.style.setProperty('--primary-gradient', val);
            }
        },

        save() {
            const current = {};
            Object.keys(this.defaults).forEach(key => {
                current[key] = getComputedStyle(document.documentElement).getPropertyValue(key).trim();
            });
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

    // --- UI HELPERS ---
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
        for(let content of contents) {
            content.style.transform = `scale(${scale})`;
        }
    },

    async changeSemitones(delta) {
        const current = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const next = Math.max(-11, Math.min(11, current + delta));
        this.elements.semitoneDisplay.textContent = (next > 0 ? '+' : '') + next;
        if (this.originalSlides.length > 0) {
            this.updatePreview(next);
        }
    },

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
                
                // --- UPDATED PREVIEW REGEX: Match Shapes OR Table Cells ---
                const elementRegex = /<(p:sp|a:tc)>([\s\S]*?)<\/\1>/g;
                let elMatch;
                
                while ((elMatch = elementRegex.exec(xml)) !== null) {
                    const elContent = elMatch[2];
                    const phMatch = elContent.match(/<p:ph[^>]*type="(?:title|ctrTitle|ftr|dt|sldNum)"/);
                    const isExcludedShape = !!phMatch;

                    const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g;
                    let pMatch;
                    
                    while ((pMatch = pRegex.exec(elContent)) !== null) {
                        const pContent = pMatch[1];
                        const tagRegex = /<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g;
                        let pText = '';
                        let match;
                        while ((match = tagRegex.exec(pContent)) !== null) {
                            if (match[0].startsWith('<a:br')) pText += '\n';
                            else pText += this.unescXml(match[2] || '');
                        }
                        
                        let alignment = 'left';
                        const algMatch = pContent.match(/algn="([^"]+)"/);
                        if (algMatch && algMatch[1] === 'ctr') alignment = 'center';

                        const isPlaceholderTitle = phMatch && (phMatch[0].includes('title') || phMatch[0].includes('ctrTitle'));
                        if (isPlaceholderTitle && pText.trim() && !globalSongTitle) {
                            globalSongTitle = pText.trim();
                        }

                        slideData.push({ text: pText, alignment, isTitle: isExcludedShape });
                    }
                }
                this.originalSlides.push(slideData);
            }
            this.songTitle = globalSongTitle; 
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides Loaded`;
            this.updatePreview(0);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error loading preview: " + err.message);
            this.hideLoading();
        }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';
        if (this.originalSlides.length === 0) {
            container.innerHTML = '<div class="md:col-span-2 lg:col-span-3 text-center py-20 text-slate-500 italic">No slides found.</div>';
            return;
        }
        const songTitle = this.songTitle || "";
        this.originalSlides.forEach((slideData, idx) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left sticky left-0">Slide ${idx + 1}</div>`;
            const contentDiv = document.createElement('div');
            contentDiv.className = 'slide-content'; 
                slideData.forEach((para) => {
                    const text = para.text;
                    const isTitle = para.isTitle || (songTitle && text.trim().toLowerCase() === songTitle.toLowerCase());
                    const isMetadata = /©|Copyright|Words:|Music:|Lyrics:|Chris Tomlin|CCLI|DAYEG AMBASSADOR/i.test(text);
                    if (text.trim() && !isMetadata && !isTitle) {
                        const lineDiv = document.createElement('div');
                        lineDiv.style.textAlign = para.alignment;
                        lineDiv.style.minHeight = '1.2em';
                        const transposed = this.transposeLine(para.text, semitones);
                        lineDiv.innerHTML = this.renderChordHTML(transposed);
                        contentDiv.appendChild(lineDiv);
                    }
                });
            if (contentDiv.children.length > 0) {
                card.appendChild(contentDiv);
                wrapper.appendChild(card);
                container.appendChild(wrapper);
            }
        });
        this.updateZoom();
    },

    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },

    renderChordHTML(text) {
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        let html = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return html.replace(chordRegex, '<span class="chord">$&</span>');
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
        try {
            const res = await fetch('./templates.json');
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const names = await res.json();
            document.getElementById('dirName').textContent = `${names.length} template${names.length !== 1 ? 's' : ''} available`;
            const entries = names.map(name => ({
                name,
                getFile: async () => {
                    const r = await fetch(`./${encodeURIComponent(name)}`);
                    if (!r.ok) throw new Error(`Could not load ${name}`);
                    const blob = await r.blob();
                    return new File([blob], name, { type: blob.type });
                }
            }));
            this.renderTemplateGallery(entries);
        } catch (e) {
            console.warn('templates.json load failed:', e.message);
            gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Could not read templates.json.</div>`;
        }
    },

    renderTemplateGallery(entries) {
        const gallery = document.getElementById('templateGallery');
        gallery.innerHTML = '';
        if (!entries.length) return;
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
                try {
                    card.style.opacity = '0.6';
                    const file = await entry.getFile();
                    card.style.opacity = '1';
                    this.selectTemplate({ name: entry.name, file }, card);
                } catch (e) {
                    card.style.opacity = '1';
                    alert('Could not load template');
                }
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
    },

    // --- GENERATION LOGIC ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';

        if (!file) return alert('Select a template first.');
        if (!lyrics) return alert('Lyrics are required.');

        try {
            this.showLoading('Generating...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateSlidePath = `ppt/${templateRelPath}`;
            const templateXml = await zip.file(templateSlidePath).async('string');
            const slideFileName = templateRelPath.split('/').pop();
            const relsPath = `ppt/slides/_rels/${slideFileName}.rels`;
            const templateRelsXml = zip.file(relsPath) ? await zip.file(relsPath).async('string') : null;
            
            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            if (sections.length === 0 && lyrics.trim() !== '') sections = [lyrics.trim()];
            
            const newZip = zip;
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionText = sections[i].trim();
                let slideXml = templateXml;
                slideXml = this.lockInStyleAndReplace(slideXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText);

                const name = `song_gen_${i + 1}.xml`;
                const path = `ppt/slides/${name}`;
                newZip.file(path, slideXml);
                
                if (templateNotesXml) {
                    const notesName = `notes_gen_${i + 1}.xml`;
                    const notesPath = `ppt/notesSlides/${notesName}`;
                    const formattedNotes = this.escXml(sectionText).replace(/\r?\n/g, '</a:t></a:r><a:br/><a:r><a:t xml:space="preserve">');
                    let newNotesXml = templateNotesXml.replace(/\[Presenter Note\]/g, formattedNotes);
                    newZip.file(notesPath, newNotesXml);

                    let newSlideRels = templateRelsXml.replace(/Target="..\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/${notesName}"`);
                    newZip.file(`ppt/slides/_rels/${name}.rels`, newSlideRels);

                    const notesRelXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${name}"/>
                    </Relationships>`;
                    newZip.file(`ppt/notesSlides/_rels/${notesName}.rels`, notesRelXml);
                    generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path, notesPath });
                } else {
                    if (templateRelsXml) newZip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                    generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path, notesPath: null });
                }
            }

            this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);
            const finalBlob = await newZip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            this.hideLoading();
        }
    },

    // --- REWORKED TABLE LOGIC ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        // Search in both Shapes and GraphicFrames (Tables)
        return xml.replace(/<(p:sp|p:graphicFrame)>([\s\S]*?)<\/\1>/g, (fullTag, tagName, innerContent) => {
            if (phRegex.test(innerContent)) {
                
                const xMatch = innerContent.match(/<a:off x="(\d+)"/);
                const yMatch = innerContent.match(/<a:off [^>]*y="(\d+)"/);
                const cxMatch = innerContent.match(/<a:ext cx="(\d+)"/);
                const posX = xMatch ? xMatch[1] : "0";
                const posY = yMatch ? yMatch[1] : "1000000";
                const posWidth = cxMatch ? cxMatch[1] : "9144000"; 

                const latinMatch = innerContent.match(/<a:latin typeface="([^"]+)"/);
                const templateFont = latinMatch ? latinMatch[1] : "Arial";
                const sizeMatch = innerContent.match(/sz="(\d+)"/);
                const templateSize = sizeMatch ? sizeMatch[1] : "2400"; 

                if (placeholder !== '[Lyrics and Chords]') {
                    const rPrMatch = innerContent.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                    let style = (rPrMatch ? rPrMatch[0] : '<a:rPr lang="en-US"/>');
                    const escapedText = (replacement || '').split(/\r?\n/)
                        .map(l => this.escXml(l))
                        .join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    return fullTag.replace(phRegex, escapedText);
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
                    const chords = line.match(chordRegex) || [];
                    const isChordLine = chords.length > 0 && !isTag;

                    if (isTag) {
                        tableRowsXml += this.createTableCellXml(trimmed, templateFont, templateSize, "ctr", 400000);
                    } else if (isChordLine) {
                        const chordLineEscaped = this.escXml(line).replace(/ /g, '&#160;');
                        tableRowsXml += this.createTableCellXml(chordLineEscaped, "Courier New", templateSize, "l", 400000);
                    } else {
                        tableRowsXml += this.createTableCellXml(this.escXml(line), templateFont, templateSize, "ctr", 450000);
                    }
                });

                return `
                <p:graphicFrame>
                    <p:nvGraphicFramePr><p:cNvPr id="1025" name="LyricsTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
                    <p:xfm><a:off x="${posX}" y="${posY}"/><a:ext cx="${posWidth}" cy="5000000"/></p:xfm>
                    <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                            <a:tbl>
                                <a:tblPr firstRow="0" bandRow="0"><a:tableStyleId>{5C22544A-7EE6-4342-B051-7303C2061113}</a:tableStyleId></a:tblPr>
                                <a:tblGrid><a:gridCol w="${posWidth}"/></a:tblGrid>
                                ${tableRowsXml}
                            </a:tbl>
                        </a:graphicData>
                    </a:graphic>
                </p:graphicFrame>`;
            }
            return fullTag;
        });
    },

    createTableCellXml(text, font, size, align, height) {
        return `
        <a:tr h="${height}">
            <a:tc>
                <a:txBody>
                    <a:bodyPr vert="ctr" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/>
                    <a:p>
                        <a:pPr algn="${align}"><a:buNone/></a:pPr>
                        <a:r>
                            <a:rPr sz="${size}" lang="en-US"><a:latin typeface="${font}"/><a:cs typeface="${font}"/></a:rPr>
                            <a:t xml:space="preserve">${text}</a:t>
                        </a:r>
                    </a:p>
                </a:txBody>
                <a:tcPr>
                    <a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR>
                    <a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB>
                    <a:solidFill><a:noFill/></a:solidFill>
                </a:tcPr>
            </a:tc>
        </a:tr>`;
    },

    // --- OTHER HELPERS ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file) return alert('Select file.');
        try {
            this.showLoading('Applying...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');
                content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                zip.file(path, content);
            }
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, file.name.replace('.pptx', `_transposed.pptx`));
            this.hideLoading();
        } catch (e) { this.hideLoading(); }
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        return text.split('\n').map(line => {
            if (!line.match(chordRegex)) return line;
            let result = line;
            let offset = 0;
            const matches = [...line.matchAll(chordRegex)];
            for (const m of matches) {
                const originalChord = m[0];
                const pos = m.index + offset;
                const newChord = this.shiftNote(m[1], semitones) + (m[2] || '') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '');
                result = result.substring(0, pos) + newChord + result.substring(pos + originalChord.length);
                offset += (newChord.length - originalChord.length);
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

    syncPresentationRegistry(newZip, presXml, presRelsXml, generated) {
        const sldIdLst = '<p:sldIdLst>' + generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        newZip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));
        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let rels = relsDoc.getElementsByTagName('Relationship');
        for (let j = rels.length - 1; j >= 0; j--) if (rels[j].getAttribute('Type').endsWith('slide')) rels[j].parentNode.removeChild(rels[j]);
        generated.forEach(s => {
            let el = relsDoc.createElement('Relationship');
            el.setAttribute('Id', s.rid);
            el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            el.setAttribute('Target', `slides/${s.name}`);
            relsDoc.documentElement.appendChild(el);
        });
        newZip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));
        const ctHead = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>';
        let ctEntries = generated.map(s => `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` + (s.notesPath ? `<Override PartName="/${s.notesPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` : '')).join('');
        newZip.file('[Content_Types].xml', (ctHead + ctEntries + '</Types>').replace('><Override', '><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'));
    },

    getPlaceholderRegexStr(ph) {
        const inner = ph.replace(/[\[\]]/g, '').trim();
        const pts = inner.split('');
        return '\\[' + '(?:<[^>]+>|\\s)*' + pts.map((p, i) => (p === ' ' ? '\\s+' : this.escRegex(p)) + (i < pts.length - 1 ? '(?:<[^>]+>|\\s)*' : '')).join('') + '(?:<[^>]+>|\\s)*' + '\\]';
    },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(slideRelsXml) {
        if (!slideRelsXml) return null;
        const m = slideRelsXml.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/);
        return m ? `ppt/notesSlides/${m[1]}` : null;
    }
};

App.init();
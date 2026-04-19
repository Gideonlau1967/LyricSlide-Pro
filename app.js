/* LyricSlide Pro - Core Logic v15.2 (Integrated Generation, Transposition & Presenter Notes) */

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
        this.loadDefaultTemplates(); // Auto-load from templates.json
        window.LyricApp = this;
        console.log("App Initialized. Version 15.2 (Notes & Background Fix)");
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
        },

        adjustColor(hex, percent) {
            const num = parseInt(hex.replace('#',''), 16),
                  amt = Math.round(2.55 * percent),
                  R = (num >> 16) + amt,
                  G = (num >> 8 & 0x00FF) + amt,
                  B = (num & 0x0000FF) + amt;
            return '#' + (0x1000000 + (R<255?R<0?0:R:255)*0x10000 + (G<255?G<0?0:G:255)*0x100 + (B<255?B<0?0:B:255)).toString(16).slice(1);
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

    toggleThemeSidebar() {
        document.getElementById('themeSidebar').classList.toggle('open');
        document.getElementById('sidebarBackdrop').classList.toggle('open');
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
                const spRegex = /<p:sp>([\s\S]*?)<\/p:sp>/g;
                let spMatch;
                while ((spMatch = spRegex.exec(xml)) !== null) {
                    const spContent = spMatch[1];
                    const phMatch = spContent.match(/<p:ph[^>]*type="(?:title|ctrTitle|ftr|dt|sldNum)"/);
                    const isExcludedShape = !!phMatch;
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
                        let alignment = 'left';
                        const algMatch = pContent.match(/algn="([^"]+)"/);
                        if (algMatch && algMatch[1] === 'ctr') alignment = 'center';
                        const isPlaceholderTitle = phMatch && (phMatch[0].includes('title') || phMatch[0].includes('ctrTitle'));
                        if (isPlaceholderTitle && pText.trim() && !globalSongTitle) globalSongTitle = pText.trim();
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
        const zoomSlider = document.getElementById('zoomSlider');
        this.updateZoom(zoomSlider ? zoomSlider.value : 100);
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
                    const blob = await r.blob();
                    return new File([blob], name, { type: blob.type });
                }
            }));
            this.renderTemplateGallery(entries);
        } catch (e) {
            document.getElementById('dirName').textContent = 'Could not load templates';
            gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Templates load failed.</div>`;
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

    // --- GENERATION LOGIC (v15.2) ---
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

                // --- NEW: Create Presenter Notes ---
                const noteName = `notesSlideGen${i + 1}.xml`;
                newZip.file(`ppt/notesSlides/${noteName}`, this.createNotesSlideXml(sectionContent));

                // Rel: Slide -> Notes & Layout (FIXES BACKGROUND)
                const slideRelXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/relationships">
                    <Relationship Id="rIdNotes1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/${noteName}"/>
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>`;
                newZip.file(`ppt/slides/_rels/${name}.rels`, slideRelXml);

                // Rel: NotesSlide -> Master (Crucial for visibility)
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

    // --- REGISTRY SYNC (CRITICAL FOR BACKGROUNDS & NOTES) ---
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

    // --- TRANSPOSITION LOGIC ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const getSz = id => { const v = parseFloat(document.getElementById(id).value); return (!isNaN(v) && v > 0) ? Math.round(v * 100) : null; };
        const getFont = id => document.getElementById(id).value.trim();
        const titleFont = getFont('fontTitle'); const titleSize = getSz('fontSizeTitle');
        const lyricsFont = getFont('fontLyrics'); const lyricsSize = getSz('fontSizeLyrics');
        const copyFont = getFont('fontCopyright'); const copySize = getSz('fontSizeCopyright');
        const anyFontChange = titleFont || titleSize || lyricsFont || lyricsSize || copyFont || copySize;

        if (!file) return alert('Select a PPTX file to transpose.');
        try {
            this.showLoading('Applying changes...');
            const zip = await JSZip.loadAsync(file);

            // Transpose Slides
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');
                if (semitones !== 0) {
                    content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                }
                if (anyFontChange) {
                    content = content.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shapeMatch, shapeContent) => {
                        const isTitle = /<p:ph[^>]*type="(?:title|ctrTitle)"/.test(shapeContent);
                        const isFooter = /<p:ph[^>]*type="ftr"/.test(shapeContent);
                        const isSkipped = /<p:ph[^>]*type="(?:dt|sldNum)"/.test(shapeContent);
                        if (isSkipped) return shapeMatch;
                        const plainText = shapeContent.replace(/<[^>]+>/g, '');
                        const isCopyright = isFooter || /©|copyright|ccli/i.test(plainText);
                        let tf, ts;
                        if (isTitle) { tf = titleFont; ts = titleSize; }
                        else if (isCopyright) { tf = copyFont; ts = copySize; }
                        else { tf = lyricsFont; ts = lyricsSize; }
                        if (!tf && !ts) return shapeMatch;
                        return `<p:sp>${this.applyFontToShapeXml(shapeContent, tf, ts)}</p:sp>`;
                    });
                }
                zip.file(path, content);
            }

            // --- NEW: Transpose Notes ---
            if (semitones !== 0) {
                const noteFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/notesSlides/notesSlide') && k.endsWith('.xml'));
                for (const path of noteFiles) {
                    let noteXml = await zip.file(path).async('string');
                    noteXml = noteXml.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                    zip.file(path, noteXml);
                }
            }

            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, file.name.replace('.pptx', `_modified.pptx`));
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    applyFontToShapeXml(shapeXml, fontFamily, fontSizeHundredths) {
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)\/>/g, (_, attrs) => {
            const newAttrs = this.applyFontSizeToAttrs(attrs, fontSizeHundredths);
            return `<a:rPr${newAttrs}>${this.buildFontTags(fontFamily)}</a:rPr>`;
        });
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/g, (_, attrs, inner) => {
            const newAttrs = this.applyFontSizeToAttrs(attrs, fontSizeHundredths);
            if (fontFamily) {
                inner = inner.replace(/<a:(latin|ea|cs)[^>]*\/>/g, '').replace(/<a:(latin|ea|cs)[^>]*>[\s\S]*?<\/a:\1>/g, '');
            }
            return `<a:rPr${newAttrs}>${this.buildFontTags(fontFamily)}${inner}</a:rPr>`;
        });
        return shapeXml;
    },

    applyFontSizeToAttrs(attrs, fontSizeHundredths) {
        if (!fontSizeHundredths) return attrs;
        if (/\bsz="[^"]+"/.test(attrs)) return attrs.replace(/\bsz="[^"]+"/, `sz="${fontSizeHundredths}"`);
        return attrs + ` sz="${fontSizeHundredths}"`;
    },

    buildFontTags(fontFamily) {
        if (!fontFamily) return '';
        const e = this.escXml(fontFamily);
        return `<a:latin typeface="${e}"/><a:ea typeface="${e}"/><a:cs typeface="${e}"/>`;
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        const lines = text.split('\n');
        return lines.map(line => {
            const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
            const words = line.split(/\s+/).filter(w => w.length > 0);
            const matches = [...line.matchAll(chordRegex)];
            if (matches.length === 0 || matches.length < words.length * 0.4) return line;
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
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                <p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                    <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${lines}</a:p></p:txBody>
                </p:sp>
            </p:spTree></p:cSld>
            <p:clrMapOver r:id="rIdNotesMaster1"/>
        </p:notes>`;
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
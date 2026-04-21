/* LyricSlide Pro */

const App = {
    version: "Version 2.0",
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

        const versionEl = document.getElementById('appVersion');
        if (versionEl) {
            versionEl.textContent = this.version;
        }
        
        console.log(`App Initialized. Version ${this.version}`);
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

        this.originalSlides.forEach((slideData, idx) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left sticky left-0">Slide ${idx + 1}</div>`;

            const contentDiv = document.createElement('div');
            contentDiv.className = 'slide-content'; 
            
            // Preview Pairing Logic
            for (let i = 0; i < slideData.length; i++) {
                const para = slideData[i];
                const text = para.text;
                const isMetadata = /©|Copyright|Words:|Music:|Lyrics:| Chris Tomlin|CCLI|DAYEG AMBASSADOR/i.test(text);
                
                if (text.trim() && !isMetadata && !para.isTitle) {
                    const lineDiv = document.createElement('div');
                    lineDiv.style.textAlign = 'center'; // Preview defaults to center
                    lineDiv.style.minHeight = '1.2em';
                    const transposed = this.transposeLine(text, semitones);
                    lineDiv.innerHTML = this.renderChordHTML(transposed);
                    contentDiv.appendChild(lineDiv);
                }
            }
            
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
            document.getElementById('dirName').textContent = 'Could not load templates';
            gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Could not read templates.json.</div>`;
        }
    },

    renderTemplateGallery(entries) {
        const gallery = document.getElementById('templateGallery');
        gallery.innerHTML = '';
        if (!entries.length) {
            gallery.innerHTML = '<div class="text-center py-8 text-slate-400 text-xs italic">No templates found.</div>';
            return;
        }
        const grid = document.createElement('div');
        grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div');
            card.className = 'template-card';
            card.title = entry.name;
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
                    alert('Could not load template: ' + e.message);
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

    clearTemplate() {
        this.selectedTemplateFile = null;
        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
        document.getElementById('selectedTemplateInfo').classList.add('hidden');
    },

    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;

        if (!file) return alert('Please select a template from the Template Library first.');
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
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);

                const name = `song_gen_${i + 1}.xml`;
                const path = `ppt/slides/${name}`;
                newZip.file(path, slideXml);
                
                let notesPath = null;
                if (templateNotesXml) {
                    const notesName = `notes_gen_${i + 1}.xml`;
                    notesPath = `ppt/notesSlides/${notesName}`;
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
                } else {
                    if (templateRelsXml) newZip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                }
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path, notesPath });
            }

            this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);
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

    // --- PRO-ALIGNMENT REPLACEMENT LOGIC ---
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
                    const escapedText = rawLines.map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    return shapeXml.replace(phRegex, escapedText);
                }

                let injectedXml = `</a:t></a:r></a:p>`;
                
                for (let i = 0; i < rawLines.length; i++) {
                    let line = rawLines[i];
                    let nextLine = rawLines[i + 1];

                    // Identify Musical Pair (Chords + Lyrics)
                    if (this.isChordLine(line) && nextLine !== undefined && !this.isChordLine(nextLine) && !nextLine.trim().startsWith('[')) {
                        
                        let chordText = line;
                        let lyricText = nextLine;
                        let finalAlign = (userAlign === 'smart' || userAlign === 'ctr') ? 'ctr' : 'l';

                        // Symmetric Balancing for Centered Modes
                        if (finalAlign === 'ctr') {
                            const overhangRight = Math.max(0, chordText.length - lyricText.length);
                            const leftPadding = " ".repeat(overhangRight);
                            chordText = leftPadding + chordText;
                            lyricText = leftPadding + lyricText + " ".repeat(overhangRight);
                        }

                        injectedXml += this.makePptLine(chordText, style.replace('<a:rPr', '<a:rPr sz="1600"'), finalAlign);
                        injectedXml += this.makePptLine(lyricText, style, finalAlign);
                        i++; 
                    } 
                    else {
                        if (line.trim() !== "") {
                            let singleAlign = (userAlign === 'smart' || userAlign === 'ctr') ? 'ctr' : 'l';
                            injectedXml += this.makePptLine(line.trim(), style, singleAlign);
                        } else {
                            injectedXml += `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                        }
                    }
                }

                injectedXml += `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r>${style}<a:t xml:space="preserve">`;
                let result = shapeXml.replace(phRegex, () => injectedXml);
                result = result.replace(/<a:p><a:pPr[^>]*><a:buNone\/><\/a:pPr><a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r><\/a:p>/g, '');
                
                if (!result.includes('Autofit')) {
                    result = result.replace('</a:bodyPr>', '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/></a:bodyPr>');
                }
                return result;
            }
            return shapeXml;
        });
    },

    makePptLine(text, style, align) {
        const escaped = this.escXml(text).replace(/ /g, '\u00A0');
        return `<a:p>
            <a:pPr algn="${align}"><a:buNone/></a:pPr>
            <a:r>${style}<a:t xml:space="preserve">${escaped}</a:t></a:r>
        </a:p>`;
    },

    isChordLine(text) {
        if (!text || text.trim() === '') return false;
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        const matches = text.match(chordRegex) || [];
        const words = text.trim().split(/\s+/).filter(w => w.length > 0);
        return matches.length > 0 && (matches.length >= words.length * 0.4 || words.length < 3);
    },

    // --- TRANSPOSITION LOGIC ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const getSz = id => { const v = parseFloat(document.getElementById(id).value); return (!isNaN(v) && v > 0) ? Math.round(v * 100) : null; };
        const getFont = id => document.getElementById(id).value.trim();

        const titleFont = getFont('fontTitle');    const titleSize = getSz('fontSizeTitle');
        const lyricsFont = getFont('fontLyrics');   const lyricsSize = getSz('fontSizeLyrics');
        const copyFont = getFont('fontCopyright'); const copySize = getSz('fontSizeCopyright');
        const anyFontChange = titleFont || titleSize || lyricsFont || lyricsSize || copyFont || copySize;

        if (!file) return alert('Select a PPTX file to transpose.');
        if (semitones === 0 && !anyFontChange) return alert('Please select a transposition amount and/or choose font settings.');

        try {
            this.showLoading('Applying changes...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files)
                .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
                .sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

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
                        let targetFont, targetSize;
                        if (isTitle) { targetFont = titleFont; targetSize = titleSize; }
                        else if (isCopyright) { targetFont = copyFont; targetSize = copySize; }
                        else { targetFont = lyricsFont; targetSize = lyricsSize; }
                        if (!targetFont && !targetSize) return shapeMatch;
                        return `<p:sp>${this.applyFontToShapeXml(shapeContent, targetFont, targetSize)}</p:sp>`;
                    });
                }
                zip.file(path, content);
            }
            if (semitones !== 0) {
                const notesFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/notesSlides/notesSlide') && k.endsWith('.xml'));
                for (const path of notesFiles) {
                    let notesContent = await zip.file(path).async('string');
                    notesContent = notesContent.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                    zip.file(path, notesContent);
                }
            }
            this.showLoading('Downloading...');
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            const suffix = [semitones !== 0 ? `${semitones > 0 ? 'plus' : 'minus'}${Math.abs(semitones)}` : '', anyFontChange ? 'fontChanged' : ''].filter(Boolean).join('_');
            saveAs(finalBlob, file.name.replace('.pptx', `_${suffix || 'modified'}.pptx`));
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error: " + err.message);
            this.hideLoading();
        }
    },

    applyFontToShapeXml(shapeXml, fontFamily, fontSizeHundredths) {
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)\/>/g, (_, attrs) => {
            const newAttrs = this.applyFontSizeToAttrs(attrs, fontSizeHundredths);
            return `<a:rPr${newAttrs}>${this.buildFontTags(fontFamily)}</a:rPr>`;
        });
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/g, (_, attrs, inner) => {
            const newAttrs = this.applyFontSizeToAttrs(attrs, fontSizeHundredths);
            if (fontFamily) {
                inner = inner.replace(/<a:latin[^>]*\/>/g, '').replace(/<a:ea[^>]*\/>/g, '').replace(/<a:cs[^>]*\/>/g, '')
                             .replace(/<a:latin[^>]*>[\s\S]*?<\/a:latin>/g, '').replace(/<a:ea[^>]*>[\s\S]*?<\/a:ea>/g, '').replace(/<a:cs[^>]*>[\s\S]*?<\/a:cs>/g, '');
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
        const e = fontFamily.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        return `<a:latin typeface="${e}"/><a:ea typeface="${e}"/><a:cs typeface="${e}"/>`;
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        const lines = text.split('\n');
        return lines.map(line => {
            const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
            const words = line.split(/\s+/).filter(w => w.length > 0);
            const chordCount = (line.match(chordRegex) || []).length;
            if (chordCount === 0 || chordCount < words.length * 0.4) return line;
            let result = line;
            let offset = 0;
            const matches = [...line.matchAll(chordRegex)];
            for (const m of matches) {
                const originalChord = m[0];
                const pos = m.index + offset;
                const newRoot = this.shiftNote(m[1], semitones);
                const newBass = m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '';
                const newChord = newRoot + (m[2] || '') + newBass;
                const diff = newChord.length - originalChord.length;
                result = result.substring(0, pos) + newChord + result.substring(pos + originalChord.length);
                if (diff > 0) {
                    let spaceMatch = result.substring(pos + newChord.length).match(/^ +/);
                    if (spaceMatch && spaceMatch[0].length >= diff) result = result.substring(0, pos + newChord.length) + result.substring(pos + newChord.length + diff);
                    else offset += diff;
                } else if (diff < 0) {
                    result = result.substring(0, pos + newChord.length) + " ".repeat(Math.abs(diff)) + result.substring(pos + newChord.length);
                }
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(note);
        if (idx === -1) {
            list = (list === this.musical.keys ? this.musical.flats : this.musical.keys);
            idx = list.indexOf(note);
        }
        if (idx === -1) return note;
        let newIdx = (idx + semitones + 12) % 12;
        return (semitones >= 0 ? this.musical.keys : this.musical.flats)[newIdx];
    },

    syncPresentationRegistry(newZip, presXml, presRelsXml, generated) {
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
        const ctXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>';
        let ctEntries = generated.map(s => `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` + (s.notesPath ? `<Override PartName="/${s.notesPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` : '')).join('');
        newZip.file('[Content_Types].xml', (ctXml + ctEntries + '</Types>').replace('><Override', '><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'));
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

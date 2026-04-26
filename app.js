/* LyricSlide Pro */

const App = {
    version: "2.5.0",
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
        // This follows your fixed convention:
        preferred: ['C', 'C#', 'D', 'Eb', 'E', 'F', 'F#', 'G', 'Ab', 'A', 'Bb', 'B']
    },

    originalSlides: [],   
    selectedTemplateFile: null, 

    
    // IMPROVED REGEX: Fixed to capture the entire suffix group without overwriting
    chordRegex: /\b([A-G][b#]?)((?:m|maj|dim|aug|sus|add|[245679]|11|13|[\(\)])*)(\/[A-G][b#]?)?(?=\s|$|[\(\)\[\]\s,])/g,

    init() {
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        
        document.getElementById('alignmentSelect').addEventListener('change', () => {
            if (this.originalSlides.length > 0) this.updatePreview(0);
        });

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
            this.showLoading('Extracting notes...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files)
                .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
                .sort((a, b) => {
                    const numA = parseInt(a.match(/\d+/)[0]);
                    const numB = parseInt(b.match(/\d+/)[0]);
                    return numA - numB;
                });

            this.originalSlides = [];
            for (const path of slideFiles) {
                const slideFileName = path.split('/').pop();
                const relsPath = `ppt/slides/_rels/${slideFileName}.rels`;
                const relsXml = zip.file(relsPath) ? await zip.file(relsPath).async('string') : null;
                const notesPath = this.getNotesRelPath(relsXml);
                
                // --- START OF INSERTED/MODIFIED SECTION ---
                let slideData = [];  // For the visual slide layout
                let notesText = "";   // For the hidden presenter notes text

                if (notesPath && zip.file(notesPath)) {
                    const notesXml = await zip.file(notesPath).async('string');
                    const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g;
                    let pMatch;
                    while ((pMatch = pRegex.exec(notesXml)) !== null) {
                        const pContent = pMatch[1];
                        const tagRegex = /<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g;
                        let pText = '';
                        let match;
                        while ((match = tagRegex.exec(pContent)) !== null) {
                            if (match[0].startsWith('<a:br')) pText += '\n';
                            else pText += this.unescXml(match[2] || '');
                        }
                        
                        if (pText.trim()) {
                            // We add it to the visual slide array
                            slideData.push({ text: pText, alignment: 'left', isTitle: false });
                            // And we append it to the full notes string
                            notesText += pText + '\n';
                        }
                    }
                }

                // Push an object containing both instead of just an array
                this.originalSlides.push({
                    slideContent: slideData.length > 0 ? slideData : [{ text: "[No Content]", alignment: 'left' }],
                    notes: notesText.trim()
                });
                // --- END OF INSERTED/MODIFIED SECTION ---
            }
            
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides`;
            this.updatePreview(0);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error: " + err.message);
            this.hideLoading();
        }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        const userAlign = document.getElementById('alignmentSelect').value;
        container.innerHTML = '';
    
        if (this.originalSlides.length === 0) {
            container.innerHTML = '<div class="text-center py-20 text-slate-500 italic">No slides loaded.</div>';
            return;
        }
    
        this.originalSlides.forEach((slideObj, idx) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'preview-card-wrapper';
            
            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left">Slide ${idx + 1}</div>`;
    
            // 1. Render Slide Content (Lyrics)
            const contentDiv = document.createElement('div');
            contentDiv.className = 'slide-content'; 
            
            slideObj.slideContent.forEach(para => {
                const originalText = para.text;
                // Filter out metadata and titles
                const isMetadata = /©|Copyright|Words:|Music:|Lyrics:|Chris Tomlin|CCLI|DAYEG AMBASSADOR/i.test(originalText);
                
                if (originalText.trim() && !isMetadata && !para.isTitle) {
                    const lineDiv = document.createElement('div');
                    lineDiv.style.textAlign = (userAlign === 'l' ? 'left' : 'center'); 
                    
                    // Transpose and highlight chords
                    const transposed = this.transposeLine(originalText, semitones);
                    lineDiv.innerHTML = this.renderChordHTML(transposed);
                    contentDiv.appendChild(lineDiv);
                }
            });
            card.appendChild(contentDiv);
    
            // 2. Render Presenter Notes (Transposed)
            if (slideObj.notes) {
                const notesDiv = document.createElement('div');
                // Adding a visual border and distinct color for the notes area
                notesDiv.className = 'mt-4 pt-2 border-t border-slate-200 text-left bg-slate-50/50 -mx-4 px-4 pb-2';
                
                const transposedNotes = this.transposeLine(slideObj.notes, semitones);
                
                notesDiv.innerHTML = `
                    <div class="text-[9px] text-amber-600 font-bold uppercase mb-1 tracking-wider">Presenter Notes</div>
                    <div class="text-[11px] text-slate-500 font-mono whitespace-pre-wrap leading-relaxed">${this.renderChordHTML(transposedNotes)}</div>
                `;
                card.appendChild(notesDiv);
            }
    
            wrapper.appendChild(card);
            container.appendChild(wrapper);
        });
    
        this.updateZoom();
    },

    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },

    renderChordHTML(text) {
        let html = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return html.replace(this.chordRegex, '<span class="chord">$&</span>');
    },

    showLoading(text) {
        this.elements.loadingText.textContent = text;
        this.elements.loadingOverlay.style.display = 'flex';
    },

    hideLoading() {
        this.elements.loadingOverlay.style.display = 'none';
    },

    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            document.getElementById('dirName').textContent = `${names.length} templates`;
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
            gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Template library unavailable.</div>`;
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
            const img = document.createElement('img');
            img.className = 'template-thumb';
            img.src = entry.name.replace(/\.pptx$/i, '.png');
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
                const file = await entry.getFile();
                this.selectTemplate({ name: entry.name, file }, card);
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

    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;

        if (!file || !lyrics) return alert('Select a template and input lyrics.');

        try {
            this.showLoading('Generating PPTX...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateSlidePath = `ppt/${templateRelPath}`;
            const templateXml = await zip.file(templateSlidePath).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templateRelPath.split('/').pop()}.rels`).async('string');
            
            const templateNotesPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesPath ? await zip.file(templateNotesPath).async('string') : null;

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            if (sections.length === 0 && lyrics.trim() !== '') sections = [lyrics.trim()];
            
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
                    const styleMatch = templateNotesXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/);
                    const notesStyle = styleMatch ? styleMatch[0] : '<a:rPr lang="en-US" sz="1600"/>';
                    const formattedNotes = this.escXml(sectionText).replace(/\r?\n/g, `</a:t></a:r><a:br/><a:r>${notesStyle}<a:t xml:space="preserve">`);
                    const notesRegex = new RegExp(this.getPlaceholderRegexStr('[Presenter Note]'), 'gi');
                    newZip.file(notesPath, templateNotesXml.replace(notesRegex, formattedNotes));
                    newZip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml.replace(/Target="..\/notesSlides\/notesSlide\d+\.xml"/, `Target="../notesSlides/${notesName}"`));
                    newZip.file(`ppt/notesSlides/_rels/${notesName}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/${name}"/></Relationships>`);
                    generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path, notesPath });
                } else {
                    newZip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                    generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path });
                }
            }

            this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);
            const finalBlob = await newZip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error: " + err.message);
            this.hideLoading();
        }
    },

    lockInStyleAndReplace(xml, placeholder, replacement, userAlign = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(placeholder), 'gi');
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
                const isCenter = (userAlign === 'ctr');
                for (let i = 0; i < rawLines.length; i++) {
                    let line = rawLines[i], nextLine = rawLines[i + 1];
                    
                    // Logic for Chord + Lyric pairs
                    if (this.isChordLine(line) && nextLine !== undefined && !this.isChordLine(nextLine) && !nextLine.trim().startsWith('[')) {
                        if (isCenter) {
                            const maxLen = Math.max(line.length, nextLine.length);
                            injectedXml += this.makeGhostAlignmentLine(line.padEnd(maxLen, ' '), nextLine.padEnd(maxLen, ' '), style, 'ctr');
                            injectedXml += this.makePptLine(nextLine.padEnd(maxLen, ' '), style, 'ctr');
                        } else {
                            injectedXml += this.makePptLine(line, this.getChordStyle(style), 'l');
                            injectedXml += this.makePptLine(nextLine, style, 'l');
                        }
                        i++;
                    } else {
                        const text = line.trim();
                        const alignTag = isCenter ? 'ctr' : 'l';
                        
                        // NEW: Detect section tags like [Verse], [Bridge]
                        const isSectionTag = text.startsWith('[') && text.endsWith(']');
                        let currentStyle = style;
                        
                        if (isSectionTag) {
                            // Set font size to 20pt (2000 in PPT XML)
                            currentStyle = currentStyle.includes('sz=') 
                                ? currentStyle.replace(/sz="\d+"/, 'sz="2000"') 
                                : currentStyle.replace('<a:rPr', '<a:rPr sz="2000"');
                        }

                        if (text !== "") {
                            injectedXml += this.makePptLine(text, currentStyle, alignTag);
                        } else {
                            injectedXml += `<a:p><a:pPr algn="${alignTag}"><a:buNone/></a:pPr><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                        }
                    }
                }
                injectedXml += `<a:p><a:pPr algn="${isCenter ? 'ctr' : 'l'}"><a:buNone/></a:pPr><a:r>${style}<a:t xml:space="preserve">`;
                return shapeXml.replace(phRegex, () => injectedXml).replace(/<a:p><a:pPr[^>]*><a:buNone\/><\/a:pPr><a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r><\/a:p>/g, '')
                                .replace('</a:bodyPr>', '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/></a:bodyPr>');
            }
            return shapeXml;
        });
    },

    getChordStyle(lyricStyle) {
        let s = lyricStyle;
        if (s.endsWith('/>')) s = s.replace('/>', '></a:rPr>');
        
        // 1. Set chord font size to 18pt
        if (s.includes('sz=')) {
            s = s.replace(/sz="\d+"/, 'sz="1800"');
        } else {
            s = s.replace('<a:rPr', '<a:rPr sz="1800"');
        }

        // 2. Set chord color to Medium Grey (#808080)
        const greyFill = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>';
        if (s.includes('<a:solidFill>')) {
            s = s.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, greyFill);
        } else {
            s = s.replace('</a:rPr>', greyFill + '</a:rPr>');
        }

        return s;
    },

    makeGhostAlignmentLine(chordLine, lyricLine, lyricStyle, align) {
        const chordStyle = this.getChordStyle(lyricStyle);
        let ghostStyle = lyricStyle;
        if (ghostStyle.endsWith('/>')) ghostStyle = ghostStyle.replace('/>', '></a:rPr>');
        
        // Ensure ghost text is invisible
        ghostStyle = ghostStyle.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, '').replace('<a:rPr', '<a:rPr><a:noFill/>');
        
        let runsXml = "";
        for (let i = 0; i < chordLine.length; i++) {
            const chordChar = chordLine[i], lyricChar = lyricLine[i] || '\u00A0';
            if (chordChar === ' ' || chordChar === '\u00A0') {
                runsXml += `<a:r>${ghostStyle}<a:t xml:space="preserve">${this.escXml(lyricChar).replace(/ /g, '\u00A0')}</a:t></a:r>`;
            } else {
                runsXml += `<a:r>${chordStyle}<a:t xml:space="preserve">${this.escXml(chordChar).replace(/ /g, '\u00A0')}</a:t></a:r>`;
            }
        }
        // Spacing set to 50% (50000)
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc><a:buNone/></a:pPr>${runsXml}</a:p>`;
    },

    makePptLine(text, style, align) {
        const escapedText = this.escXml(text).replace(/ /g, '\u00A0'); 
        let finalStyle = style;
        if (finalStyle.endsWith('/>')) finalStyle = finalStyle.replace('/>', '></a:rPr>');
        
        // Spacing set to 50% (50000)
        return `<a:p><a:pPr algn="${align}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc><a:buNone/></a:pPr><a:r>${finalStyle}<a:t xml:space="preserve">${escapedText}</a:t></a:r></a:p>`;
    },

    isChordLine(line) {
        if (!line || line.trim() === '') return false;
        const words = line.trim().split(/\s+/);
        const chords = line.match(this.chordRegex) || [];
        return chords.length >= words.length * 0.6 || (chords.length > 0 && words.length <= 2);
    },

    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file) return alert('Select file.');
        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file);
            const slides = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
            for (const path of slides) {
                let content = await zip.file(path).async('string');
                content = this.transposeParagraphs(content, semitones);
                zip.file(path, content);
            }
            const notes = Object.keys(zip.files).filter(k => k.startsWith('ppt/notesSlides/notesSlide') && k.endsWith('.xml'));
            for (const path of notes) {
                let nc = await zip.file(path).async('string');
                nc = this.transposeParagraphs(nc, semitones);
                zip.file(path, nc);
            }
            saveAs(await zip.generateAsync({ type: 'blob' }), file.name.replace('.pptx', `_transposed.pptx`));
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    transposeParagraphs(xml, semitones) {
        return xml.replace(/<a:p[^>]*>([\s\S]*?)<\/a:p>/g, (matchFull, pXml) => {
            let logicLine = ""; 
            let characterMetadata = []; // Stores { isGhost: bool, originalChar: string, style: string }

            const runRegex = /<a:r>([\s\S]*?)<\/a:r>|<a:br\/>/g;
            let m;
            
            while ((m = runRegex.exec(pXml)) !== null) {
                if (m[0] === '<a:br/>') {
                    logicLine += "\n";
                    characterMetadata.push({ isBr: true });
                    continue;
                }
                
                const rPrMatch = m[1].match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/);
                const rStyle = rPrMatch ? rPrMatch[0] : '<a:rPr/>';
                const isGhost = rStyle.includes('<a:noFill/>');
                
                const tMatch = m[1].match(/<a:t[^>]*>(.*?)<\/a:t>/);
                const text = tMatch ? this.unescXml(tMatch[1]) : "";

                for (let char of text) {
                    characterMetadata.push({
                        isGhost: isGhost,
                        originalChar: char,
                        style: rStyle,
                        isBr: false
                    });
                    // If it's a ghost, we treat it as a space for the chord-detector
                    logicLine += isGhost ? " " : char;
                }
            }

            if (!logicLine.trim()) return matchFull;

            const transposedLogic = this.transposeLine(logicLine, semitones);
            if (transposedLogic === logicLine) return matchFull;

            // RECONSTRUCTION
            const pPrMatch = pXml.match(/<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/);
            const pPr = pPrMatch ? pPrMatch[0] : '';
            const pTagMatch = matchFull.match(/^<a:p[^>]*>/);
            const pTagOpen = pTagMatch ? pTagMatch[0] : '<a:p>';

            let newRuns = "";
            let metaIdx = 0;

            for (let i = 0; i < transposedLogic.length; i++) {
                const newChar = transposedLogic[i];
                
                if (newChar === "\n") {
                    newRuns += "<a:br/>";
                    // Skip the newline in metadata
                    while(metaIdx < characterMetadata.length && !characterMetadata[metaIdx].isBr) metaIdx++;
                    metaIdx++; 
                    continue;
                }

                const meta = characterMetadata[metaIdx] || { isGhost: false, style: '<a:rPr sz="1800"/>' };

                if (meta.isGhost && newChar === " ") {
                    // This position was a ghost character (lyric) and remains a space in the chord line
                    // We must use the original style (Large) to keep alignment
                    newRuns += `<a:r>${meta.style}<a:t xml:space="preserve">${this.escXml(meta.originalChar).replace(/ /g, '\u00A0')}</a:t></a:r>`;
                } else {
                    // This is a visible chord character
                    // We force the Chord Style (18pt)
                    const chordStyle = this.getChordStyle(meta.style).replace('<a:noFill/>', ''); 
                    newRuns += `<a:r>${chordStyle}<a:t xml:space="preserve">${this.escXml(newChar).replace(/ /g, '\u00A0')}</a:t></a:r>`;
                }
                metaIdx++;
            }

            return `${pTagOpen}${pPr}${newRuns}</a:p>`;
        });
    },

    transposeLine(text, semitones) {
        // If no transposition is needed or text is empty, return original
        if (semitones === 0 || !text) return text;
    
        const lines = text.split('\n');
        return lines.map(line => {
            // IMPORTANT: We removed "if (!this.isChordLine(line)) return line;"
            // This allows chords inside Presenter Notes (like "[G] Amazing Grace") 
            // to be transposed even though the line is mostly lyrics.
    
            let result = line;
            let offset = 0;
            
            // Find all chord matches in the current line
            const matches = [...line.matchAll(this.chordRegex)];
            
            // If no chords are found in this specific line, return it as is
            if (matches.length === 0) return line;
    
            for (const m of matches) {
                const originalFull = m[0];       // e.g., "C#m7/G"
                const rootNote = m[1];           // e.g., "C#"
                const suffix = m[2] || '';       // e.g., "m7"
                const bassPart = m[3] || '';     // e.g., "/G"
    
                // 1. Shift the main root note
                const newRoot = this.shiftNote(rootNote, semitones);
    
                // 2. Shift the bass note if it exists (e.g., the "G" in "C/G")
                let newBass = '';
                if (bassPart) {
                    const bassNote = bassPart.substring(1); // Remove the "/"
                    newBass = '/' + this.shiftNote(bassNote, semitones);
                }
    
                const newChord = newRoot + suffix + newBass;
    
                // 3. Calculate position with current offset (offset changes as string length changes)
                const position = m.index + offset;
                const lengthDiff = newChord.length - originalFull.length;
    
                let pre = result.substring(0, position);
                let suf = result.substring(position + originalFull.length);
    
                // 4. Alignment Logic (mainly for chords-over-lyrics on slides)
                // If chord got longer and there is a space after it, remove a space
                if (lengthDiff > 0 && suf.startsWith(' ')) {
                    suf = suf.substring(1);
                    offset--;
                } 
                // If chord got shorter, add spaces to keep the rest of the line aligned
                else if (lengthDiff < 0) {
                    suf = ' '.repeat(Math.abs(lengthDiff)) + suf;
                    offset += Math.abs(lengthDiff);
                }
    
                // Construct the new line
                result = pre + newChord + suf;
                offset += lengthDiff;
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        // 1. Find the numeric index of the current note (0-11)
        let idx = this.musical.keys.indexOf(note);
        if (idx === -1) idx = this.musical.flats.indexOf(note);
        
        // If note not found (shouldn't happen with chordRegex), return as is
        if (idx === -1) return note;

        // 2. Calculate the new index based on semitones
        let newIdx = (idx + semitones) % 12;
        if (newIdx < 0) newIdx += 12;

        // 3. Always return the note from your fixed convention list
        return this.musical.preferred[newIdx];
    },

    syncPresentationRegistry(newZip, presXml, presRelsXml, generated) {
        const sldIdLst = '<p:sldIdLst>' + generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        newZip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));
        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let rels = relsDoc.getElementsByTagName('Relationship');
        for (let j = rels.length - 1; j >= 0; j--) if (rels[j].getAttribute('Type').endsWith('slide')) rels[j].parentNode.removeChild(rels[j]);
        generated.forEach(s => {
            let el = relsDoc.createElement('Relationship');
            el.setAttribute('Id', s.rid); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); el.setAttribute('Target', `slides/${s.name}`);
            relsDoc.documentElement.appendChild(el);
        });
        newZip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));
    },

    getPlaceholderRegexStr(ph) {
        const inner = ph.replace(/[\[\]]/g, '').trim(), pts = inner.split('');
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

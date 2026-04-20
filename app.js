/* LyricSlide Pro - Core Logic v12 (Integrated Generation & Transposition) */

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
        console.log("App Initialized. Version 15.0 (Auto-Template)");
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
            // Load saved theme
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                this.setVariable(key, val);
                
                // Update picker UI
                const pickerId = 'picker-' + key.replace('--', '').replace('-color', '');
                const picker = document.getElementById(pickerId);
                if (picker) picker.value = val;
            });

            // Set up listeners
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
                // Keep it flat for "no accent" look
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
                
                // Match shapes to identify placeholders
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
                        
                        // Detect alignment
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
            this.songTitle = globalSongTitle; // Store globally

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

        // Add Header Container for Title (REMOVED)

        this.originalSlides.forEach((slideData, idx) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'preview-card-wrapper';

            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left sticky left-0">Slide ${idx + 1}</div>`;

            const contentDiv = document.createElement('div');
            contentDiv.className = 'slide-content'; // TARGET FOR ZOOM
                slideData.forEach((para, pIdx) => {
                    const text = para.text;
                    const isTitle = para.isTitle || (songTitle && text.trim().toLowerCase() === songTitle.toLowerCase());
                    const isMetadata = /©|Copyright|Words:|Music:|Lyrics:|Chris Tomlin|CCLI|DAYEG AMBASSADOR/i.test(text);
                    
                    if (text.trim() && !isMetadata && !isTitle) {
                        const lineDiv = document.createElement('div');
                        lineDiv.style.textAlign = para.alignment;
                        lineDiv.style.minHeight = '1.2em';
                        const transposed = this.transposeLine(para.text, semitones);
                        // Wrap chords in span for styling
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

        // Re-apply zoom/scaling (v13 uses updateZoom)
        const zoomSlider = document.getElementById('zoomSlider');
        if (typeof updateZoom === 'function') updateZoom(zoomSlider ? zoomSlider.value : 100);
    },

    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },

    renderChordHTML(text) {
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        // Escape existing HTML just in case
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
            gallery.innerHTML = `
                <div class="text-center py-8 text-slate-400 text-xs italic">
                    <i class="fas fa-exclamation-circle mr-1"></i>
                    Could not read templates.json.<br>
                    Make sure this page is served via HTTP (not file://).
                </div>`;
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

            // Thumbnail — try PNG with same base name, fall back to icon
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
        const infoBar = document.getElementById('selectedTemplateInfo');
        infoBar.classList.remove('hidden');
        document.getElementById('selectedTemplateName').textContent = item.name;
    },

    clearTemplate() {
        this.selectedTemplateFile = null;
        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
        document.getElementById('selectedTemplateInfo').classList.add('hidden');
    },

    // --- GENERATION LOGIC (v11) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';

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
            
            // --- MODIFIED: Detect Presenter Notes associated with the template slide ---
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
                
                let notesPath = null;
                // --- MODIFIED: Clone presenter notes and replace placeholder ---
                if (templateNotesXml) {
                    const notesName = `notes_gen_${i + 1}.xml`;
                    notesPath = `ppt/notesSlides/${notesName}`;
                    
                    // Format text to maintain PPT XML line breaks instead of letting \n break raw strings
                    const formattedNotes = this.escXml(sectionText).replace(/\r?\n/g, '</a:t></a:r><a:br/><a:r><a:t xml:space="preserve">');
                    let newNotesXml = templateNotesXml.replace(/\[Presenter Note\]/g, formattedNotes);
                    newZip.file(notesPath, newNotesXml);

                    // Update slide rels to reference the newly generated note slide file
                    let newSlideRels = templateRelsXml.replace(
                        /Target="..\/notesSlides\/notesSlide\d+\.xml"/, 
                        `Target="../notesSlides/${notesName}"`
                    );
                    newZip.file(`ppt/slides/_rels/${name}.rels`, newSlideRels);

                    // Create the mapping for the note pointing back to the specific slide
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
    // --- TRANSPOSITION LOGIC ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;

        // Read per-section font settings
        const getSz = id => { const v = parseFloat(document.getElementById(id).value); return (!isNaN(v) && v > 0) ? Math.round(v * 100) : null; };
        const getFont = id => document.getElementById(id).value.trim();

        const titleFont   = getFont('fontTitle');    const titleSize   = getSz('fontSizeTitle');
        const lyricsFont  = getFont('fontLyrics');   const lyricsSize  = getSz('fontSizeLyrics');
        const copyFont    = getFont('fontCopyright');const copySize    = getSz('fontSizeCopyright');

        const anyFontChange = titleFont || titleSize || lyricsFont || lyricsSize || copyFont || copySize;

        if (!file) return alert('Select a PPTX file to transpose.');
        if (semitones === 0 && !anyFontChange) return alert('Please select a transposition amount and/or choose font settings.');

        try {
            this.showLoading('Applying changes...');
            const zip = await JSZip.loadAsync(file);
            
            // 1. Process Slide Files (Text + Optional Fonts)
            const slideFiles = Object.keys(zip.files)
                .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
                .sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');

                // Transpose chords in slides
                if (semitones !== 0) {
                    content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) =>
                        `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                }

                // Apply font changes to slides
                if (anyFontChange) {
                    content = content.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shapeMatch, shapeContent) => {
                        const isTitle    = /<p:ph[^>]*type="(?:title|ctrTitle)"/.test(shapeContent);
                        const isFooter   = /<p:ph[^>]*type="ftr"/.test(shapeContent);
                        const isSkipped  = /<p:ph[^>]*type="(?:dt|sldNum)"/.test(shapeContent);
                        if (isSkipped) return shapeMatch;

                        const plainText  = shapeContent.replace(/<[^>]+>/g, '');
                        const isCopyright = isFooter || /©|copyright|ccli/i.test(plainText);

                        let targetFont, targetSize;
                        if (isTitle)          { targetFont = titleFont;  targetSize = titleSize;  }
                        else if (isCopyright) { targetFont = copyFont;   targetSize = copySize;   }
                        else                  { targetFont = lyricsFont; targetSize = lyricsSize; }

                        if (!targetFont && !targetSize) return shapeMatch;
                        return `<p:sp>${this.applyFontToShapeXml(shapeContent, targetFont, targetSize)}</p:sp>`;
                    });
                }
                zip.file(path, content);
            }

            // 2. NEW: Process Presenter Notes (Transpose Chords)
            if (semitones !== 0) {
                const notesFiles = Object.keys(zip.files)
                    .filter(k => k.startsWith('ppt/notesSlides/notesSlide') && k.endsWith('.xml'));

                for (const path of notesFiles) {
                    let notesContent = await zip.file(path).async('string');
                    
                    // Transpose chords inside the <a:t> tags of the presenter notes
                    notesContent = notesContent.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) =>
                        `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                    
                    zip.file(path, notesContent);
                }
            }

            this.showLoading('Downloading...');
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            const suffix = [
                semitones !== 0 ? `${semitones > 0 ? 'plus' : 'minus'}${Math.abs(semitones)}` : '',
                anyFontChange ? 'fontChanged' : ''
            ].filter(Boolean).join('_');
            saveAs(finalBlob, file.name.replace('.pptx', `_${suffix || 'modified'}.pptx`));
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error: " + err.message);
            this.hideLoading();
        }
    },

    applyFontToShapeXml(shapeXml, fontFamily, fontSizeHundredths) {
        // Handle self-closing <a:rPr .../>
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)\/>/g, (_, attrs) => {
            const newAttrs = this.applyFontSizeToAttrs(attrs, fontSizeHundredths);
            return `<a:rPr${newAttrs}>${this.buildFontTags(fontFamily)}</a:rPr>`;
        });
        // Handle open <a:rPr ...>...</a:rPr>
        shapeXml = shapeXml.replace(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/g, (_, attrs, inner) => {
            const newAttrs = this.applyFontSizeToAttrs(attrs, fontSizeHundredths);
            if (fontFamily) {
                inner = inner
                    .replace(/<a:latin[^>]*\/>/g, '')
                    .replace(/<a:ea[^>]*\/>/g, '')
                    .replace(/<a:cs[^>]*\/>/g, '')
                    .replace(/<a:latin[^>]*>[\s\S]*?<\/a:latin>/g, '')
                    .replace(/<a:ea[^>]*>[\s\S]*?<\/a:ea>/g, '')
                    .replace(/<a:cs[^>]*>[\s\S]*?<\/a:cs>/g, '');
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
            
            // If it's not a chord line, return as is (preserving all spaces)
            if (chordCount === 0 || chordCount < words.length * 0.4) return line;

            let result = line;
            let offset = 0;
            const matches = [...line.matchAll(chordRegex)];
            for (const m of matches) {
                const originalChord = m[0];
                const pos = m.index + offset;
                const root = m[1];
                const suffix = m[2] || '';
                const bass = m[3] || '';
                
                const newRoot = this.shiftNote(root, semitones);
                const newBass = bass ? '/' + this.shiftNote(bass.substring(1), semitones) : '';
                const newChord = newRoot + suffix + newBass;

                const diff = newChord.length - originalChord.length;
                result = result.substring(0, pos) + newChord + result.substring(pos + originalChord.length);
                
                if (diff > 0) {
                    let spaceMatch = result.substring(pos + newChord.length).match(/^ +/);
                    if (spaceMatch && spaceMatch[0].length >= diff) {
                        result = result.substring(0, pos + newChord.length) + 
                                 result.substring(pos + newChord.length + diff);
                    } else {
                        offset += diff;
                    }
                } else if (diff < 0) {
                    result = result.substring(0, pos + newChord.length) + 
                             " ".repeat(Math.abs(diff)) + 
                             result.substring(pos + newChord.length);
                }
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = this.musical.keys;
        if (note.includes('b')) list = this.musical.flats;
        
        let idx = list.indexOf(note);
        if (idx === -1) {
            // Try the other list if not found (e.g. seeking D# in a flat song)
            list = (list === this.musical.keys ? this.musical.flats : this.musical.keys);
            idx = list.indexOf(note);
        }
        
        if (idx === -1) return note; // Give up
        
        let newIdx = (idx + semitones + 12) % 12;
        // Use sharps for positive shifts, flats for negative, or match input
        const outList = semitones >= 0 ? this.musical.keys : this.musical.flats;
        return outList[newIdx];
    },

// --- TABLE METHOD: FULL-WIDTH CENTERED WITH LENGTH NORMALIZATION ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                
                // 1. EXTRACT TEMPLATE SETTINGS
                const latinMatch = shapeXml.match(/<a:latin typeface="([^"]+)"/);
                const templateFont = latinMatch ? latinMatch[1] : "Arial";
                const sizeMatch = shapeXml.match(/sz="(\d+)"/);
                const templateSize = sizeMatch ? sizeMatch[1] : "2400"; 

                if (placeholder !== '[Lyrics and Chords]') {
                    const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                    let style = (rPrMatch ? rPrMatch[0] : '<a:rPr lang="en-US"/>');
                    const escapedText = (replacement || '').split(/\r?\n/)
                        .map(l => this.escXml(l))
                        .join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    return shapeXml.replace(phRegex, escapedText);
                }

                // 2. WIDESCREEN CONSTANTS (16:9)
                const MAX_SLIDE_WIDTH = 12192000; 
                const lines = (replacement || '').split(/\r?\n/);
                let tableRowsXml = '';

                for (let i = 0; i < lines.length; i++) {
                    let line = lines[i];
                    let trimmed = line.trim();
                    
                    if (trimmed === '') {
                        tableRowsXml += `<a:tr h="150000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t> </a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>`;
                        continue;
                    }

                    const isTag = trimmed.startsWith('[') && trimmed.endsWith(']');
                    const chords = line.match(chordRegex) || [];
                    const isChordLine = chords.length > 0 && !isTag;

                    let typeface = templateFont;
                    let fontSize = templateSize;
                    let processedText = line;

                    // 3. PRECISION ALIGNMENT LOGIC (LENGTH NORMALIZATION)
                    if (isChordLine) {
                        typeface = "Courier New"; // Monospace for 1:1 character alignment
                        fontSize = Math.round(parseInt(templateSize) * 0.8);
                        
                        // If there is a lyric line below, make them the same length
                        if (lines[i+1] && !lines[i+1].trim().startsWith('[')) {
                            const lyricLine = lines[i+1];
                            const diff = lyricLine.length - line.length;
                            if (diff > 0) {
                                processedText = line + " ".repeat(diff); // Pad chord line
                            }
                        }
                    } else if (!isTag && i > 0) {
                        // Check if the chord line ABOVE was longer than this lyric
                        const lineAbove = lines[i-1];
                        if (lineAbove.match(chordRegex)) {
                            const diff = lineAbove.length - line.length;
                            if (diff > 0) {
                                processedText = line + " ".repeat(diff); // Pad lyric line
                            }
                        }
                    }

                    // Convert spaces to Non-Breaking Spaces for XML stability
                    const escapedLine = this.escXml(processedText).replace(/ /g, '&#160;');

                    tableRowsXml += `
                        <a:tr h="400000">
                            <a:tc>
                                <a:txBody>
                                    <a:bodyPr vert="ctr" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/>
                                    <a:p>
                                        <a:pPr algn="ctr"><a:buNone/></a:pPr>
                                        <a:r>
                                            <a:rPr sz="${fontSize}" lang="en-US">
                                                <a:latin typeface="${typeface}"/>
                                                <a:cs typeface="${typeface}"/>
                                            </a:rPr>
                                            <a:t xml:space="preserve">${escapedLine}</a:t>
                                        </a:r>
                                    </a:p>
                                </a:txBody>
                                <a:tcPr/>
                            </a:tc>
                        </a:tr>`;
                }

                // 4. GENERATE FULL-WIDTH TABLE
                return `
                <p:graphicFrame>
                    <p:nvGraphicFramePr>
                        <p:cNvPr id="1025" name="LyricsTable"/>
                        <p:cNvGraphicFramePr/><p:nvPr/>
                    </p:nvGraphicFramePr>
                    <p:xfm>
                        <a:off x="0" y="1000000"/> 
                        <a:ext cx="${MAX_SLIDE_WIDTH}" cy="5000000"/>
                    </p:xfm>
                    <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                            <a:tbl>
                                <a:tblPr firstRow="0" bandRow="0"><a:tableStyleId>{5C22544A-7EE6-4342-B051-7303C2061113}</a:tableStyleId></a:tblPr>
                                <a:tblGrid><a:gridCol w="${MAX_SLIDE_WIDTH}"/></a:tblGrid>
                                ${tableRowsXml}
                            </a:tbl>
                        </a:graphicData>
                    </a:graphic>
                </p:graphicFrame>`;
            }
            return shapeXml;
        });
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
        
        // --- MODIFIED: Injecting notesSlide overrides into the content list dynamically ---
        let ctEntries = generated.map(s => {
            let entries = `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
            if (s.notesPath) {
                entries += `<Override PartName="/${s.notesPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
            }
            return entries;
        }).join('');

        // We actually need to keep the themes and masters in [Content_Types].xml. Simplified approach:
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
    
    // --- NEW HELPER: Extract associated Notes XML path from a slide's relationships ---
    getNotesRelPath(slideRelsXml) {
        if (!slideRelsXml) return null;
        const m = slideRelsXml.match(/Relationship[^>]+Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/);
        return m ? `ppt/notesSlides/${m[1]}` : null;
    }
};

App.init();

/* LyricSlide Pro - Core Logic v15.6 (Rigid Block-Lock & GitHub Optimized) */

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
        console.log("App Initialized. Version 15.6 (Rigid Block-Lock Centering)");
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
        if (this.originalSlides.length > 0) this.updatePreview(next);
    },

    async loadForPreview(file) {
        try {
            this.showLoading('Extracting slide text...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files)
                .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
                .sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

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
                        if (phMatch && (phMatch[0].includes('title') || phMatch[0].includes('ctrTitle')) && pText.trim() && !globalSongTitle) {
                            globalSongTitle = pText.trim();
                        }
                        slideData.push({ text: pText, alignment, isTitle: !!phMatch });
                    }
                }
                this.originalSlides.push(slideData);
            }
            this.songTitle = globalSongTitle;
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides Loaded`;
            this.updatePreview(0);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';
        if (this.originalSlides.length === 0) return;
        this.originalSlides.forEach((slideData, idx) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left">Slide ${idx + 1}</div>`;
            const contentDiv = document.createElement('div');
            contentDiv.className = 'slide-content';
            slideData.forEach((para) => {
                if (para.text.trim() && !para.isTitle && !/©|Copyright|CCLI/i.test(para.text)) {
                    const lineDiv = document.createElement('div');
                    lineDiv.style.textAlign = para.alignment;
                    lineDiv.innerHTML = this.renderChordHTML(this.transposeLine(para.text, semitones));
                    contentDiv.appendChild(lineDiv);
                }
            });
            card.appendChild(contentDiv);
            wrapper.appendChild(card);
            container.appendChild(wrapper);
        });
        this.updateZoom();
    },

    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    renderChordHTML(text) { return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g, '<span class="chord">$&</span>'); },
    showLoading(text) { this.elements.loadingText.textContent = text; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },

    // --- TEMPLATE LIBRARY (GitHub Path Fix Preserved) ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        const basePath = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1);
        try {
            const res = await fetch(`${basePath}templates.json`);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const names = await res.json();
            document.getElementById('dirName').textContent = `${names.length} templates available`;
            const entries = names.map(name => ({
                name,
                getFile: async () => {
                    const r = await fetch(`${basePath}${encodeURIComponent(name)}`);
                    const blob = await r.blob();
                    return new File([blob], name, { type: blob.type });
                }
            }));
            this.renderTemplateGallery(entries);
        } catch (e) {
            gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Could not read templates.json.</div>`;
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
            card.innerHTML = `<img class="template-thumb" src="${entry.name.replace(/\.pptx$/i, '.png')}" onerror="this.src='https://placehold.co/200x120?text=PPTX'"><div class="template-card-name">${entry.name.replace(/\.pptx$/i, '')}</div>`;
            card.onclick = async () => {
                const file = await entry.getFile();
                this.selectedTemplateFile = file;
                document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                card.classList.add('selected');
                document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                document.getElementById('selectedTemplateName').textContent = entry.name;
            };
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },

    // --- GENERATION LOGIC ---
    async generate() {
        const file = this.selectedTemplateFile;
        const lyrics = this.elements.lyricsInput.value;
        if (!file || !lyrics) return alert('Select template and enter lyrics.');

        try {
            this.showLoading('Locking Chords...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const slideFileName = templateRelPath.split('/').pop();
            const templateRelsXml = zip.file(`ppt/slides/_rels/${slideFileName}.rels`) ? await zip.file(`ppt/slides/_rels/${slideFileName}.rels`).async('string') : null;

            const sections = ("\n" + lyrics).split(/\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/).filter(s => s.trim() !== '');
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                let sXml = templateXml;
                sXml = this.lockInStyleAndReplace(sXml, '[Title]', this.elements.songTitle.value);
                sXml = this.lockInStyleAndReplace(sXml, '[Copyright Info]', this.elements.copyrightInfo.value);
                sXml = this.lockInStyleAndReplace(sXml, '[Lyrics and Chords]', sections[i].trim());

                const name = `slide_gen_${i + 1}.xml`;
                zip.file(`ppt/slides/${name}`, sXml);
                if (templateRelsXml) zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path: `ppt/slides/${name}` });
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            const blob = await zip.generateAsync({ type: 'blob' });
            saveAs(blob, `${(this.elements.songTitle.value || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    // --- REPLACEMENT ENGINE (RIGID BLOCK-LOCK) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');

        return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));

                let rawLines = (replacement || '').split(/\r?\n/);
                
                if (placeholder === '[Lyrics and Chords]') {
                    // Logic: Find longest line and pad right with NBSPs to lock vertical alignment
                    const maxLen = Math.max(...rawLines.map(l => l.length));
                    rawLines = rawLines.map(l => {
                        const paddingCount = maxLen - l.length;
                        return l.replace(/ /g, '\u00A0') + '\u00A0'.repeat(paddingCount);
                    });
                }

                const lines = rawLines.map(l => this.escXml(l));
                let injected = '';
                lines.forEach((line, idx) => {
                    if (idx > 0) injected += `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`;
                    injected += line;
                });

                let result = shapeXml.replace(phRegex, () => {
                    return `</a:t></a:r><a:r>${style}<a:t xml:space="preserve">${injected}</a:t></a:r><a:r>${style}<a:t xml:space="preserve">`;
                });

                // Force Paragraph Centering for Lyrics
                if (placeholder === '[Lyrics and Chords]') {
                    if (result.includes('<a:pPr')) {
                        result = result.replace(/<a:pPr([^>]*)>/, (m, attrs) => attrs.includes('algn=') ? m.replace(/algn="[^"]*"/, 'algn="ctr"') : `<a:pPr${attrs} algn="ctr">`);
                    } else {
                        result = result.replace(/<a:p>/g, '<a:p><a:pPr algn="ctr"/>');
                    }
                }

                result = result.replace(/<a:t xml:space="preserve"><\/a:t>/g, '').replace(/<a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r>/g, '');
                if (!result.includes('Autofit')) result = result.replace('</a:bodyPr>', '<a:normAutofit fontScale="85000" lnSpcReduction="15000"/></a:bodyPr>');
                return result;
            }
            return shapeXml;
        });
    },

    // --- TRANSPOSITION ---
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
                content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                zip.file(path, content);
            }
            const blob = await zip.generateAsync({ type: 'blob' });
            saveAs(blob, file.name.replace('.pptx', '_transposed.pptx'));
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        return text.split('\n').map(line => {
            const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
            const matches = [...line.matchAll(chordRegex)];
            if (matches.length === 0) return line;
            let result = line, offset = 0;
            for (const m of matches) {
                const newC = this.shiftNote(m[1], semitones) + (m[2] || '') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '');
                const diff = newC.length - m[0].length;
                result = result.substring(0, m.index + offset) + newC + result.substring(m.index + offset + m[0].length);
                offset += diff;
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(note);
        if (idx === -1) { list = (list === this.musical.keys ? this.musical.flats : this.musical.keys); idx = list.indexOf(note); }
        return (semitones >= 0 ? this.musical.keys : this.musical.flats)[(idx + semitones + 12) % 12];
    },

    syncPresentationRegistry(zip, presXml, presRelsXml, gen) {
        const sldIdLst = '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));
        let rDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let rs = rDoc.getElementsByTagName('Relationship');
        for (let j = rs.length - 1; j >= 0; j--) if (rs[j].getAttribute('Type').endsWith('slide')) rs[j].parentNode.removeChild(rs[j]);
        gen.forEach(s => {
            let el = rDoc.createElement('Relationship');
            el.setAttribute('Id', s.rid); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); el.setAttribute('Target', `slides/${s.name}`);
            rDoc.documentElement.appendChild(el);
        });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(rDoc));
    },

    getPlaceholderRegexStr(ph) {
        const inner = ph.replace(/[\[\]]/g, '').trim();
        return '\\[' + inner.split('').map(p => p === ' ' ? '\\s+' : this.escRegex(p)).join('(?:<[^>]+>|\\s)*') + '\\]';
    },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; }
};

App.init();
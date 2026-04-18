/* LyricSlide Pro - Core Logic v21 (Restored Loader + Block-Lock) */

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
        console.log("App Initialized. v21.0 (Restored GitHub Pathing)");
    },

    // --- THEME & UI (RESTORED ORIGINAL) ---
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
                    this.setVariable(this.getVarNameFromPicker(e.target.id), e.target.value);
                    this.save();
                });
            });
        },
        getVarNameFromPicker(id) {
            const map = { 'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end', 'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg', 'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color' };
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

    // --- TEMPLATE LIBRARY (RESTORED ORIGINAL WORKING VERSION) ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const names = await res.json();

            document.getElementById('dirName').textContent = `${names.length} templates available`;

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
            gallery.innerHTML = `<div class="text-center py-8 text-slate-400 text-xs italic">Could not load templates.json</div>`;
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
            card.innerHTML = `
                <img class="template-thumb" src="${entry.name.replace(/\.pptx$/i, '.png')}" onerror="this.src='https://placehold.co/200x120?text=PPTX'">
                <div class="template-card-name">${entry.name.replace(/\.pptx$/i, '')}</div>
            `;
            card.addEventListener('click', async () => {
                try {
                    const file = await entry.getFile();
                    this.selectedTemplateFile = file;
                    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                    card.classList.add('selected');
                    document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                    document.getElementById('selectedTemplateName').textContent = entry.name;
                } catch (e) { alert(e.message); }
            });
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },

    // --- CORE REPLACEMENT ENGINE (AMENDED WITH BLOCK-LOCK) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');

        return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));

                let rawLines = (replacement || '').split(/\r?\n/);
                
                // NEW: Block-Lock Logic for Monospaced Alignment
                if (placeholder === '[Lyrics and Chords]') {
                    const maxLen = Math.max(...rawLines.map(l => l.length));
                    rawLines = rawLines.map(l => {
                        const paddingCount = maxLen - l.length;
                        // Replace ALL spaces with Non-Breaking Spaces (\u00A0)
                        const coreText = l.replace(/ /g, '\u00A0'); 
                        return coreText + '\u00A0'.repeat(paddingCount);
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

                // Force Centering for lyrics
                if (placeholder === '[Lyrics and Chords]') {
                    if (result.includes('<a:pPr')) {
                        result = result.replace(/<a:pPr([^>]*)>/, (m, attrs) => {
                            return attrs.includes('algn=') ? m.replace(/algn="[^"]*"/, 'algn="ctr"') : `<a:pPr${attrs} algn="ctr">`;
                        });
                    } else {
                        result = result.replace(/<a:p>/g, '<a:p><a:pPr algn="ctr"/>');
                    }
                }

                result = result.replace(/<a:t xml:space="preserve"><\/a:t>/g, '').replace(/<a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r>/g, '');
                if (!result.includes('Autofit')) {
                    result = result.replace('</a:bodyPr>', '<a:normAutofit fontScale="85000" lnSpcReduction="15000"/></a:bodyPr>');
                }
                return result;
            }
            return shapeXml;
        });
    },

    // --- PPTX GENERATION & TRANSPOSITION ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        if (!file || !lyrics) return alert('Select template and enter lyrics.');

        try {
            this.showLoading('Generating...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            
            const sections = ("\n" + lyrics).split(/\r?\n(?=\s*\[[^\]]+\])/).filter(s => s.trim() !== '');
            const newZip = zip;
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                let sXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                sXml = this.lockInStyleAndReplace(sXml, '[Copyright Info]', copyright);
                sXml = this.lockInStyleAndReplace(sXml, '[Lyrics and Chords]', sections[i].trim());
                const name = `slide_gen_${i+1}.xml`;
                newZip.file(`ppt/slides/${name}`, sXml);
                generated.push({ id: 5000+i, rid: `rIdGen${i+1}`, name, path: `ppt/slides/${name}` });
            }
            this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);
            const blob = await newZip.generateAsync({ type: 'blob' });
            saveAs(blob, `${title.replace(/[^a-z0-9]/gi, '_') || 'Song'}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file || semitones === 0) return alert('Select file and semitones.');
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

    // --- HELPERS (STABLE) ---
    transposeLine(text, semitones) {
        return text.split('\n').map(line => {
            const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
            const matches = [...line.matchAll(chordRegex)];
            if (matches.length === 0) return line;
            let res = line, off = 0;
            for (const m of matches) {
                const newC = this.shiftNote(m[1], semitones) + (m[2] || '') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '');
                const diff = newC.length - m[0].length;
                res = res.substring(0, m.index + off) + newC + res.substring(m.index + off + m[0].length);
                off += diff;
            }
            return res;
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
        return '\\[' + '(?:<[^>]+>|\\s)*' + inner.split('').map((p, i) => (p === ' ' ? '\\s+' : this.escRegex(p)) + (i < inner.length - 1 ? '(?:<[^>]+>|\\s)*' : '')).join('') + '(?:<[^>]+>|\\s)*' + '\\]';
    },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; }
};

App.init();
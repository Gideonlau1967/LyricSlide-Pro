/* LyricSlide Pro - Core Logic v17 (Centered & Locked Chords Fix) */

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
        console.log("App Initialized. Version 17.0 (Equal-Length Padding Fix)");
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
                const p = document.getElementById('picker-' + key.replace('--', '').replace('-color', ''));
                if (p) p.value = val;
            });
            document.querySelectorAll('.color-picker-input').forEach(p => {
                p.addEventListener('input', (e) => {
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
            const cur = {};
            Object.keys(this.defaults).forEach(k => cur[k] = getComputedStyle(document.documentElement).getPropertyValue(k).trim());
            localStorage.setItem('lyric_theme', JSON.stringify(cur));
        }
    },

    // --- UI HELPERS ---
    setMode(mode) {
        const isGen = mode === 'gen';
        document.getElementById('viewGen').classList.toggle('hidden', !isGen);
        document.getElementById('viewTrans').classList.toggle('hidden', isGen);
    },

    updateZoom(val) {
        if (val === undefined) val = document.getElementById('zoomSlider').value;
        document.getElementById('zoomVal').textContent = val + '%';
        const scale = val / 100;
        const contents = document.getElementsByClassName('slide-content');
        for(let c of contents) c.style.transform = `scale(${scale})`;
    },

    async changeSemitones(delta) {
        const current = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const next = Math.max(-11, Math.min(11, current + delta));
        this.elements.semitoneDisplay.textContent = (next > 0 ? '+' : '') + next;
        if (this.originalSlides.length > 0) this.updatePreview(next);
    },

    async loadForPreview(file) {
        try {
            this.showLoading('Extracting text...');
            const zip = await JSZip.loadAsync(file);
            const files = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml')).sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
            this.originalSlides = [];
            for (const path of files) {
                const xml = await zip.file(path).async('string');
                const data = [];
                const spRegex = /<p:sp>([\s\S]*?)<\/p:sp>/g;
                let m;
                while ((m = spRegex.exec(xml)) !== null) {
                    const sp = m[1];
                    const ph = sp.match(/<p:ph[^>]*type="(?:title|ctrTitle|ftr|dt|sldNum)"/);
                    const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g;
                    let pm;
                    while ((pm = pRegex.exec(sp)) !== null) {
                        const pText = pm[1].replace(/<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g, (tag, t, content) => tag.startsWith('<a:br') ? '\n' : this.unescXml(content || ''));
                        let alg = 'left';
                        if (pm[1].includes('algn="ctr"')) alg = 'center';
                        data.push({ text: pText, alignment: alg, isTitle: !!ph });
                    }
                }
                this.originalSlides.push(data);
            }
            this.updatePreview(0);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';
        this.originalSlides.forEach((slide, idx) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div');
            card.className = 'preview-card';
            card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2 uppercase font-black">Slide ${idx + 1}</div>`;
            const content = document.createElement('div');
            content.className = 'slide-content';
            slide.forEach(p => {
                if (p.text.trim() && !p.isTitle && !/©|Copyright|CCLI/i.test(p.text)) {
                    const line = document.createElement('div');
                    line.style.textAlign = p.alignment;
                    line.innerHTML = this.renderChordHTML(this.transposeLine(p.text, semitones));
                    content.appendChild(line);
                }
            });
            card.appendChild(content);
            wrapper.appendChild(card);
            container.appendChild(wrapper);
        });
        this.updateZoom();
    },

    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    renderChordHTML(text) { return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g, '<span class="chord">$&</span>'); },
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },

    async loadDefaultTemplates() {
        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            const entries = names.map(name => ({
                name,
                getFile: async () => {
                    const r = await fetch(`./${encodeURIComponent(name)}`);
                    const b = await r.blob();
                    return new File([b], name, { type: b.type });
                }
            }));
            this.renderTemplateGallery(entries);
        } catch (e) { console.warn("Templates load fail"); }
    },

    renderTemplateGallery(entries) {
        const gallery = document.getElementById('templateGallery');
        const grid = document.createElement('div');
        grid.className = 'template-grid';
        entries.forEach(entry => {
            const card = document.createElement('div');
            card.className = 'template-card';
            card.innerHTML = `<img class="template-thumb" src="${entry.name.replace(/\.pptx$/i, '.png')}" onerror="this.src='https://placehold.co/200x120?text=PPTX'"><div class="template-card-name">${entry.name.replace(/\.pptx$/i, '')}</div>`;
            card.addEventListener('click', async () => {
                const file = await entry.getFile();
                this.selectedTemplateFile = file;
                document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                card.classList.add('selected');
                document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                document.getElementById('selectedTemplateName').textContent = entry.name;
            });
            grid.appendChild(card);
        });
        gallery.appendChild(grid);
    },

    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        if (!file || !lyrics) return alert('Select a template and enter lyrics.');
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
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sections[i].trim());
                const name = `song_gen_${i + 1}.xml`;
                newZip.file(`ppt/slides/${name}`, slideXml);
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path: `ppt/slides/${name}` });
            }
            this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);
            const finalBlob = await newZip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${title.replace(/[^a-z0-9]/gi, '_') || 'Song'}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    // --- UPDATED REPLACEMENT ENGINE (CENTERING FIX) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');

        return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));

                let rawLines = (replacement || '').split(/\r?\n/);
                
                // PADDING LOGIC: Make all lines in the section the same character length
                if (placeholder === '[Lyrics and Chords]') {
                    const maxCharCount = Math.max(...rawLines.map(l => l.length));
                    rawLines = rawLines.map(l => l + ' '.repeat(maxCharCount - l.length));
                }

                const lines = rawLines.map(l => this.escXml(l));
                let injected = '';
                lines.forEach((line, idx) => {
                    if (idx > 0) injected += `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`;
                    injected += line;
                });

                if (placeholder === '[Lyrics and Chords]' && lines.length > 10) {
                    const szMatch = style.match(/sz=\"(\d+)\"/);
                    if (szMatch) {
                        const scale = Math.max(0.6, 1 - (lines.length - 10) * 0.05);
                        style = style.replace(/sz=\"\d+\"/, `sz="${Math.floor(parseInt(szMatch[1]) * scale)}"`);
                    }
                }

                let result = shapeXml.replace(phRegex, () => {
                    return `</a:t></a:r><a:r>${style}<a:t xml:space="preserve">${injected}</a:t></a:r><a:r>${style}<a:t xml:space="preserve">`;
                });

                // FORCE CENTER ALIGNMENT in Paragraph Properties
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
                if (!result.includes('Autofit')) result = result.replace('</a:bodyPr>', '<a:normAutofit fontScale="75000" lnSpcReduction="15000"/></a:bodyPr>');
                return result;
            }
            return shapeXml;
        });
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        return text.split('\n').map(line => {
            const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
            const matches = [...line.matchAll(chordRegex)];
            if (matches.length === 0 || matches.length < line.split(/\s+/).filter(w => w.length > 0).length * 0.4) return line;
            let res = line, off = 0;
            for (const m of matches) {
                const newC = this.shiftNote(m[1], semitones) + (m[2] || '') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '');
                const diff = newC.length - m[0].length;
                res = res.substring(0, m.index + off) + newC + res.substring(m.index + off + m[0].length);
                if (diff > 0) {
                    let sm = res.substring(m.index + off + newC.length).match(/^ +/);
                    if (sm && sm[0].length >= diff) res = res.substring(0, m.index + off + newC.length) + res.substring(m.index + off + newC.length + diff);
                    else off += diff;
                } else if (diff < 0) {
                    res = res.substring(0, m.index + off + newC.length) + " ".repeat(Math.abs(diff)) + res.substring(m.index + off + newC.length);
                }
            }
            return res;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(note);
        if (idx === -1) { list = (list === this.musical.keys ? this.musical.flats : this.musical.keys); idx = list.indexOf(note); }
        if (idx === -1) return note;
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
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; }
};

App.init();
/* LyricSlide Pro - Core Logic v15.0 (Table Integration & Version Display) */

const App = {
    version: "15.0", // <--- VERSION DEFINED HERE

    elements: {
        songTitle: document.getElementById('songTitle'),
        lyricsInput: document.getElementById('lyricsInput'),
        copyrightInfo: document.getElementById('copyrightInfo'),
        generateBtn: document.getElementById('generateBtn'),
        
        versionDisplay: document.getElementById('versionDisplay'), // <--- ELEMENT LINKED
        
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
        
        // --- UI VERSION UPDATE ---
        if (this.elements.versionDisplay) {
            this.elements.versionDisplay.textContent = `v${this.version}`;
        }
        // --------------------------

        this.theme.init();
        this.loadDefaultTemplates();
        window.LyricApp = this;
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
                'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start', 'picker-bg-end': '--bg-end',
                'picker-text': '--text-main', 'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg',
                'picker-chord': '--preview-chord-color', 'picker-lyrics': '--preview-lyrics-color'
            };
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
        },
        reset() {
            if (confirm('Reset theme?')) {
                Object.keys(this.defaults).forEach(key => {
                    this.setVariable(key, this.defaults[key]);
                    const picker = document.getElementById('picker-' + key.replace('--', '').replace('-color', ''));
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
        for(let content of contents) content.style.transform = `scale(${scale})`;
    },
    async changeSemitones(delta) {
        const current = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const next = Math.max(-11, Math.min(11, current + delta));
        this.elements.semitoneDisplay.textContent = (next > 0 ? '+' : '') + next;
    },

    // --- COMPLEX TABLE GENERATION LOGIC ---
    createTableXml(lines, baseStyle) {
        let rowsXml = '';
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        for (let i = 0; i < lines.length; i++) {
            let line = lines[i];
            const trimmed = line.trim();
            if (trimmed === '' && i === lines.length - 1) continue;

            const isTag = trimmed.startsWith('[') && trimmed.endsWith(']');
            const chords = line.match(chordRegex) || [];
            const words = trimmed.split(/\s+/).filter(w => w.length > 0);
            const isChords = chords.length > 0 && !isTag && (chords.length >= words.length * 0.3 || words.length < 3);

            let rowStyle = baseStyle;
            let processedText = line;

            if (isTag) {
                processedText = trimmed.replace(/[\[\]]/g, ''); 
            } else if (isChords) {
                const lyricLine = lines[i + 1] || "";
                const maxLength = Math.max(line.length, lyricLine.length);
                
                // SYNC-BLOCK PADDING
                processedText = line.padEnd(maxLength, ' ');
                lines[i + 1] = lyricLine.padEnd(maxLength, ' ');

                // Force Courier New 18pt for Chords
                rowStyle = rowStyle.replace(/sz="\d+"/, 'sz="1800"')
                                   .replace(/typeface="[^"]+"/, 'typeface="Courier New"');
            }

            rowsXml += `
                <a:tr h="450000">
                    <a:tc>
                        <a:txBody>
                            <a:bodyPr vert="ctr" anchor="ctr" />
                            <a:p>
                                <a:pPr algn="ctr"><a:buNone/></a:pPr>
                                <a:r>${rowStyle}<a:t xml:space="preserve">${this.escXml(processedText)}</a:t></a:r>
                            </a:p>
                        </a:txBody>
                    </a:tc>
                </a:tr>`;
        }

        return `
            <p:graphicFrame>
                <p:nvGraphicFramePr><p:cNvPr id="1025" name="Table" /><p:cNvGraphicFramePr /><p:nvPr /></p:nvGraphicFramePr>
                <p:xfrm><a:off x="400000" y="1200000" /><a:ext cx="8328000" cy="4500000" /></p:xfrm>
                <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                            <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>
                            <a:tblGrid><a:gridCol w="8328000" /></a:tblGrid>
                            ${rowsXml}
                        </a:tbl>
                    </a:graphicData>
                </a:graphic>
            </p:graphicFrame>`;
    },

    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';

        if (!file || !lyrics) return alert('Select template and enter lyrics.');

        try {
            this.showLoading('Generating Tables...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');

            const styleMatch = templateXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/);
            const baseStyle = styleMatch ? styleMatch[0] : '<a:rPr lang="en-US" sz="2400"/>';

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = lyrics.split(splitRegex).filter(s => s.trim() !== '');

            const generated = [];
            for (let i = 0; i < sections.length; i++) {
                let slideXml = templateXml;
                const lines = sections[i].trim().split(/\r?\n/);

                const tableXml = this.createTableXml(lines, baseStyle);
                slideXml = slideXml.replace(/<p:sp>[\s\S]*?\[Lyrics and Chords\][\s\S]*?<\/p:sp>/, tableXml);
                
                slideXml = slideXml.replace(/\[Title\]/g, this.escXml(title));
                slideXml = slideXml.replace(/\[Copyright Info\]/g, this.escXml(copyright));

                const name = `song_gen_${i + 1}.xml`;
                zip.file(`ppt/slides/${name}`, slideXml);
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path: `ppt/slides/${name}` });
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${title.replace(/[^a-z0-9]/gi, '_') || 'Song'}.pptx`);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            this.hideLoading();
        }
    },

    // --- TRANSPOSITION LOGIC ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file || semitones === 0) return alert('Select file and semitone shift.');

        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
            for (const path of slideFiles) {
                let content = await zip.file(path).async('string');
                content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                zip.file(path, content);
            }
            const notesFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/notesSlides/notesSlide') && k.endsWith('.xml'));
            for (const path of notesFiles) {
                let content = await zip.file(path).async('string');
                content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => `<a:t>${this.transposeLine(text, semitones)}</a:t>`);
                zip.file(path, content);
            }
            const finalBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, file.name.replace('.pptx', `_${semitones}.pptx`));
            this.hideLoading();
        } catch (err) { console.error(err); this.hideLoading(); }
    },

    transposeLine(text, semitones) {
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        return text.split('\n').map(line => {
            if (!line.match(chordRegex)) return line;
            let result = line; let offset = 0;
            const matches = [...line.matchAll(chordRegex)];
            for (const m of matches) {
                const original = m[0]; const root = m[1]; const suffix = m[2] || ''; const bass = m[3] || '';
                const newRoot = this.shiftNote(root, semitones);
                const newBass = bass ? '/' + this.shiftNote(bass.substring(1), semitones) : '';
                const newChord = newRoot + suffix + newBass;
                result = result.substring(0, m.index + offset) + newChord + result.substring(m.index + offset + original.length);
                offset += (newChord.length - original.length);
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let list = note.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(note);
        if (idx === -1) { list = (list === this.musical.keys ? this.musical.flats : this.musical.keys); idx = list.indexOf(note); }
        if (idx === -1) return note;
        let newIdx = (idx + semitones + 12) % 12;
        return (semitones >= 0 ? this.musical.keys : this.musical.flats)[newIdx];
    },

    // --- SYSTEM HELPERS ---
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    
    syncPresentationRegistry(zip, presXml, presRelsXml, generated) {
        const sldIdLst = '<p:sldIdLst>' + generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));
        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let rels = relsDoc.getElementsByTagName('Relationship');
        for(let i=rels.length-1; i>=0; i--) if(rels[i].getAttribute('Type').endsWith('slide')) rels[i].parentNode.removeChild(rels[i]);
        generated.forEach(s => {
            let el = relsDoc.createElement('Relationship');
            el.setAttribute('Id', s.rid);
            el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            el.setAttribute('Target', `slides/${s.name}`);
            relsDoc.documentElement.appendChild(el);
        });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));
    },

    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            gallery.innerHTML = '';
            names.forEach(name => {
                const card = document.createElement('div');
                card.className = 'template-card';
                card.innerHTML = `<div class="template-card-name">${name.replace('.pptx','')}</div>`;
                card.onclick = async () => {
                    const r = await fetch(`./${encodeURIComponent(name)}`);
                    const blob = await r.blob();
                    this.selectedTemplateFile = new File([blob], name);
                    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                    card.classList.add('selected');
                };
                gallery.appendChild(card);
            });
        } catch (e) { console.warn("Templates load failed."); }
    },

    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; }
};

App.init();
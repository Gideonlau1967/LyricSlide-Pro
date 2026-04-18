/**
 * LyricSlide Pro - v15.3 (Auto-Load Templates)
 */

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

    selectedTemplateFile: null,

    init() {
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        
        // AUTO-LOAD STARTUP
        this.loadTemplatesFromDirectory();
        
        window.LyricApp = this;
        console.log("LyricSlide Pro v15.3: Ready");
    },

    // --- TEMPLATE LOADING ---
    
    // Attempt 1: Auto-load from templates.json in the same folder
    async loadTemplatesFromDirectory() {
        try {
            const response = await fetch('./templates.json');
            if (!response.ok) throw new Error("No manifest");
            
            const filenames = await response.json();
            const galleryData = filenames.map(name => {
                const baseName = name.replace(/\.pptx$/i, '');
                return {
                    name: name,
                    thumbUrl: `./${encodeURIComponent(baseName)}.png`, // Expects PNG in same folder
                    getFile: async () => {
                        const res = await fetch(`./${encodeURIComponent(name)}`);
                        const blob = await res.blob();
                        return new File([blob], name);
                    }
                };
            });
            this.renderTemplateGallery(galleryData);
        } catch (e) {
            console.warn("Auto-load failed. Waiting for manual folder selection.");
            document.getElementById('templateGallery').innerHTML = `
                <div class="text-center py-8 text-slate-400 text-[10px] leading-relaxed">
                    templates.json not found.<br>Click the folder icon above to load manually.
                </div>`;
        }
    },

    // Attempt 2: Manual Folder Picker (Backup)
    async pickTemplateDir() {
        try {
            const dirHandle = await window.showDirectoryPicker();
            const entries = [];
            const imageMap = new Map();

            this.showLoading('Scanning...');

            for await (const entry of dirHandle.values()) {
                if (entry.kind === 'file') {
                    const low = entry.name.toLowerCase();
                    if (low.endsWith('.pptx')) entries.push(entry);
                    else if (low.endsWith('.png') || low.endsWith('.jpg')) {
                        const base = entry.name.split('.').slice(0, -1).join('.');
                        imageMap.set(base, entry);
                    }
                }
            }

            const galleryData = await Promise.all(entries.map(async (handle) => {
                const baseName = handle.name.replace(/\.pptx$/i, '');
                let thumbUrl = null;
                if (imageMap.has(baseName)) {
                    const imgFile = await imageMap.get(baseName).getFile();
                    thumbUrl = URL.createObjectURL(imgFile);
                }
                return { name: handle.name, thumbUrl, getFile: () => handle.getFile() };
            }));

            this.renderTemplateGallery(galleryData);
            this.hideLoading();
        } catch (err) {
            this.hideLoading();
            if (err.name !== 'AbortError') alert("Folder access failed.");
        }
    },

    renderTemplateGallery(entries) {
        const container = document.getElementById('templateGallery');
        container.innerHTML = '';
        const grid = document.createElement('div');
        grid.className = 'template-grid';

        entries.forEach(entry => {
            const card = document.createElement('div');
            card.className = 'template-card';
            
            const thumb = document.createElement('img');
            thumb.className = 'template-thumb';
            thumb.src = entry.thumbUrl;
            thumb.onerror = () => {
                const ph = document.createElement('div');
                ph.className = 'template-thumb-placeholder';
                ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>';
                thumb.replaceWith(ph);
            };

            const name = document.createElement('div');
            name.className = 'template-card-name';
            name.textContent = entry.name.replace(/\.pptx$/i, '');

            card.appendChild(thumb);
            card.appendChild(name);
            card.onclick = async () => {
                this.showLoading('Loading Template...');
                this.selectedTemplateFile = await entry.getFile();
                document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                card.classList.add('selected');
                document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                document.getElementById('selectedTemplateName').textContent = entry.name;
                this.hideLoading();
            };
            grid.appendChild(card);
        });
        container.appendChild(grid);
    },

    clearTemplate() {
        this.selectedTemplateFile = null;
        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
        document.getElementById('selectedTemplateInfo').classList.add('hidden');
    },

    // --- ALIGNMENT ENGINE (CHORDS + LYRICS) ---
    prepareCenteredLines(text) {
        const rawLines = text.split(/\r?\n/);
        const processed = [];
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        for (let i = 0; i < rawLines.length; i++) {
            let line = rawLines[i];
            let next = rawLines[i+1] || "";
            const isChord = (line.match(chordRegex) || []).length > 0 && 
                            (line.match(chordRegex) || []).length >= line.trim().split(/\s+/).length * 0.3;

            if (isChord && next.trim() !== "" && !(next.match(chordRegex))) {
                const len = Math.max(line.length, next.length);
                processed.push(line.padEnd(len, ' '));
                processed.push(next.padEnd(len, ' '));
                i++; 
            } else {
                processed.push(line);
            }
        }
        return processed;
    },

    // --- GENERATION ---
    async generate() {
        if (!this.selectedTemplateFile) return alert("Select a template first.");
        const lyrics = this.elements.lyricsInput.value.trim();
        if (!lyrics) return alert("Enter some lyrics.");

        try {
            this.showLoading("Generating PPTX...");
            const zip = await JSZip.loadAsync(this.selectedTemplateFile);
            
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            
            const templatePath = `ppt/${slideRels[slideIds[0].rid]}`;
            const templateXml = await zip.file(templatePath).async('string');
            const relsPath = `ppt/slides/_rels/${templatePath.split('/').pop()}.rels`;
            const templateRelsXml = zip.file(relsPath) ? await zip.file(relsPath).async('string') : null;

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            const sections = lyrics.split(splitRegex).filter(s => s.trim() !== '');
            
            const gen = [];
            for (let i = 0; i < sections.length; i++) {
                let sXml = templateXml;
                sXml = this.injectContent(sXml, '[Title]', this.elements.songTitle.value);
                sXml = this.injectContent(sXml, '[Copyright Info]', this.elements.copyrightInfo.value);
                sXml = this.injectContent(sXml, '[Lyrics and Chords]', sections[i].trim());

                const name = `s_${i+1}.xml`;
                const path = `ppt/slides/${name}`;
                zip.file(path, sXml);
                if (templateRelsXml) zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                gen.push({ id: 7000 + i, rid: `rIdGen${i+1}`, name, path });
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, gen);
            const blob = await zip.generateAsync({ type: 'blob' });
            saveAs(blob, `${this.elements.songTitle.value || 'Song'}.pptx`);
            this.hideLoading();
        } catch (e) {
            console.error(e);
            alert("Error: " + e.message);
            this.hideLoading();
        }
    },

    injectContent(xml, placeholder, replacement) {
        const regex = new RegExp(this.getPlaceholderRegexStr(placeholder), 'gi');
        return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
            if (regex.test(shapeXml)) {
                const isLyrics = placeholder === '[Lyrics and Chords]';
                const alignment = '<a:pPr algn="ctr"/>'; 

                let lines = (replacement || '').split(/\n/);
                if (isLyrics) lines = this.prepareCenteredLines(replacement || '');

                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                let style = rPrMatch ? rPrMatch[0] : '<a:rPr lang="en-US"/>';

                // Auto-scale font for lyrics
                if (isLyrics && lines.length > 8) {
                    const sz = style.match(/sz=\"(\d+)\"/);
                    if (sz) {
                        const newSz = Math.floor(parseInt(sz[1]) * Math.max(0.5, 1 - (lines.length - 8) * 0.06));
                        style = style.replace(/sz=\"\d+\"/, `sz="${newSz}"`);
                    }
                }

                const paras = lines.map(l => 
                    `<a:p>${alignment}<a:r>${style}<a:t xml:space="preserve">${this.escXml(l)}</a:t></a:r></a:p>`
                ).join('');

                return shapeXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, 
                    `<p:txBody><a:bodyPr anchor="ctr" wrap="none"><a:normAutofit fontScale="92000"/></a:bodyPr><a:lstStyle/>${paras}</p:txBody>`);
            }
            return shapeXml;
        });
    },

    // --- TRANSPOSE ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semi = parseInt(this.elements.semitoneDisplay.textContent);
        if (!file) return alert("Select a file.");
        this.showLoading("Transposing...");
        const zip = await JSZip.loadAsync(file);
        const slides = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
        for (const path of slides) {
            let xml = await zip.file(path).async('string');
            xml = xml.replace(/<a:t>(.*?)<\/a:t>/g, (_, txt) => `<a:t>${this.transposeLine(txt, semi)}</a:t>`);
            zip.file(path, xml);
        }
        const blob = await zip.generateAsync({ type: 'blob' });
        saveAs(blob, file.name.replace('.pptx', '_shifted.pptx'));
        this.hideLoading();
    },

    transposeLine(text, semi) {
        if (semi === 0) return text;
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
        return text.split('\n').map(line => {
            if (!(line.match(chordRegex))) return line;
            let result = line;
            let offset = 0;
            const matches = [...line.matchAll(chordRegex)];
            for (const m of matches) {
                const pos = m.index + offset;
                const newC = this.shiftNote(m[1], semi) + (m[2]||'') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semi) : '');
                result = result.substring(0, pos) + newC + result.substring(pos + m[0].length);
                offset += (newC.length - m[0].length);
            }
            return result;
        }).join('\n');
    },

    shiftNote(n, s) {
        let list = n.includes('b') ? this.musical.flats : this.musical.keys;
        let idx = list.indexOf(n);
        if (idx === -1) idx = (list === this.musical.keys ? this.musical.flats : this.musical.keys).indexOf(n);
        if (idx === -1) return n;
        return (s >= 0 ? this.musical.keys : this.musical.flats)[(idx + s + 12) % 12];
    },

    // --- UTILS ---
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    getPlaceholderRegexStr(ph) {
        const inner = ph.replace(/[\[\]]/g, '').trim();
        return '\\[' + inner.split('').map(p => (p === ' ' ? '\\s+' : this.escRegex(p))).join('(?:<[^>]+>|\\s)*') + '\\]';
    },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    getSlideIds(x) { let ids=[], m, r=/<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while(m=r.exec(x)) ids.push({id:m[1], rid:m[2]}); return ids; },
    getSlideRels(x) { let rels={}, m, r=/<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while(m=r.exec(x)) rels[m[1]]=m[2]; return rels; },
    syncPresentationRegistry(zip, presXml, presRelsXml, gen) {
        const sldIdLst = '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));
        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let relationships = relsDoc.getElementsByTagName('Relationship');
        for (let j = relationships.length - 1; j >= 0; j--) if (relationships[j].getAttribute('Type').endsWith('slide')) relationships[j].remove();
        gen.forEach(s => {
            let el = relsDoc.createElement('Relationship');
            el.setAttribute('Id', s.rid);
            el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            el.setAttribute('Target', `slides/${s.name}`);
            relsDoc.documentElement.appendChild(el);
        });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));
    },

    setMode(m) {
        const isG = m === 'gen';
        document.getElementById('modeGen').classList.toggle('active', isG);
        document.getElementById('modeTrans').classList.toggle('active', !isG);
        document.getElementById('viewGen').classList.toggle('hidden', !isG);
        document.getElementById('viewTrans').classList.toggle('hidden', isG);
    },
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; }
};

App.init();
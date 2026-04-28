/* LyricSlide Pro - Version 3.4.1 (Full Master Integrity Build) */

const App = {
    version: "Version 3.4.1",
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
        preferred: ['C', 'C#', 'D', 'Eb', 'E', 'F', 'F#', 'G', 'Ab', 'A', 'Bb', 'B']
    },

    chordRegex: /(?:\[)?\b([A-G][b#]?)((?:m|maj|dim|aug|sus|add|[245679]|11|13|[\(\)])*)(\/[A-G][b#]?)?\b(?:\])?/g,

    originalSlides: [],   
    selectedTemplateFile: null, 
    
    init() {
        if (this.elements.generateBtn) this.elements.generateBtn.addEventListener('click', () => this.generate());
        if (this.elements.transposeBtn) this.elements.transposeBtn.addEventListener('click', () => this.transpose());
        if (this.elements.transFileInput) {
            this.elements.transFileInput.addEventListener('change', (e) => {
                if (e.target.files[0]) this.loadForPreview(e.target.files[0]);
            });
        }
        
        const alignSelect = document.getElementById('alignmentSelect');
        if (alignSelect) {
            alignSelect.addEventListener('change', () => {
                if (this.originalSlides.length > 0) this.updatePreview(0);
            });
        }

        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;

        const vEl = document.getElementById('appVersion');
        if (vEl) vEl.textContent = this.version;
    },

    // --- THEME MANAGEMENT ---
    theme: {
        defaults: {
            '--primary-color': '#334155', '--bg-start': '#f8fafc', '--bg-end': '#f8fafc',
            '--text-main': '#1e293b', '--card-accent': '#e2e8f0', '--preview-card-bg': '#ffffff',
            '--preview-chord-color': '#334155', '--preview-lyrics-color': '#1e293b'
        },
        init() {
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                this.setVariable(key, val);
                const picker = document.getElementById('picker-' + key.replace('--', '').replace('-color', ''));
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
                'picker-primary': '--primary-color', 'picker-bg-start': '--bg-start',
                'picker-bg-end': '--bg-end', 'picker-text': '--text-main',
                'picker-card-accent': '--card-accent', 'picker-preview-bg': '--preview-card-bg',
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
        const zoomValEl = document.getElementById('zoomVal');
        if (zoomValEl) zoomValEl.textContent = val + '%';
        const scale = val / 100;
        const contents = document.getElementsByClassName('preview-card');
        for(let content of contents) {
            content.style.transform = `scale(${scale})`;
            content.style.transformOrigin = 'top center';
        }
    },

    async changeSemitones(delta) {
        const current = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        const next = Math.max(-11, Math.min(11, current + delta));
        this.elements.semitoneDisplay.textContent = (next > 0 ? '+' : '') + next;
        if (this.originalSlides.length > 0) this.updatePreview(next);
    },

    // --- PREVIEW / TRANSPOSITION LOGIC ---
    async loadForPreview(file) {
        try {
            this.showLoading('Extracting notes...');
            const zip = await JSZip.loadAsync(file);
            const slideFiles = Object.keys(zip.files)
                .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
                .sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
    
            this.originalSlides = [];
            for (const path of slideFiles) {
                const slideFileName = path.split('/').pop();
                const relsPath = `ppt/slides/_rels/${slideFileName}.rels`;
                const relsXml = zip.file(relsPath) ? await zip.file(relsPath).async('string') : null;
                const notesPath = this.getNotesRelPath(relsXml);
                let notesText = ""; 
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
                            else pText += this.unescXml(match[2] || '').replace(/\u00A0/g, ' ');
                        }
                        notesText += pText + '\n';
                    }
                }
                this.originalSlides.push({ path, notesPath, notes: notesText.trim() });
            }
            document.getElementById('slideCount').textContent = `${this.originalSlides.length} Slides`;
            this.updatePreview(0);
            this.hideLoading();
        } catch (err) { alert("Error: " + err.message); this.hideLoading(); }
    },

    updatePreview(semitones) {
        const container = document.getElementById('previewContainer');
        container.innerHTML = '';
        if (this.originalSlides.length === 0) return;
        this.originalSlides.forEach((slide, idx) => {
            const wrapper = document.createElement('div');
            wrapper.className = 'preview-card-wrapper';
            const card = document.createElement('div');
            card.className = 'preview-card';
            const transposedText = this.transposeLine(slide.notes, semitones);
            card.innerHTML = `
                <div class="text-[10px] text-slate-400 mb-2 uppercase font-black text-left">Slide ${idx + 1}</div>
                <div class="whitespace-pre font-mono text-[11px] leading-snug text-left">${this.renderChordHTML(transposedText)}</div>
            `;
            wrapper.appendChild(card);
            container.appendChild(wrapper);
        });
        this.updateZoom();
    },

    renderChordHTML(text) {
        let html = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return html.replace(this.chordRegex, '<span class="chord">$&</span>');
    },

    // --- TEMPLATE GALLERY ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        if(!gallery) return;
        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            const entries = names.map(name => ({
                name, getFile: async () => {
                    const r = await fetch(`./${encodeURIComponent(name)}`);
                    const blob = await r.blob();
                    return new File([blob], name, { type: blob.type });
                }
            }));
            this.renderTemplateGallery(entries);
        } catch (e) { gallery.innerHTML = `<div class="text-center py-8 text-slate-400 italic">No templates found.</div>`; }
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
                const ph = document.createElement('div'); ph.className = 'template-thumb-placeholder';
                ph.innerHTML = '<i class="fas fa-file-powerpoint"></i>'; img.replaceWith(ph);
            });
            const nameDiv = document.createElement('div');
            nameDiv.className = 'template-card-name'; nameDiv.textContent = entry.name.replace(/\.pptx$/i, '');
            card.appendChild(img); card.appendChild(nameDiv);
            card.addEventListener('click', async () => {
                const file = await entry.getFile(); this.selectTemplate({ name: entry.name, file }, card);
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

    // --- CORE GENERATION (MASTER INTEGRITY BUILD) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const copyright = this.elements.copyrightInfo.value || '';
        const userAlign = document.getElementById('alignmentSelect').value;
        const lyrics = (this.elements.lyricsInput.value || '').trim();
        if (!file || !lyrics) return alert('Input lyrics and select template.');

        try {
            this.showLoading('Analyzing Template Architecture...');
            const zip = await JSZip.loadAsync(file);

            // 1. DATA EXTRACTION
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const templatePath = slideRels[this.getSlideIds(presXml)[0].rid];
            const templateXml = await zip.file(`ppt/${templatePath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templatePath.split('/').pop()}.rels`).async('string');
            
            const templateNotesRelPath = this.getNotesRelPath(templateRelsXml);
            const templateNotesXml = templateNotesRelPath ? await zip.file(templateNotesRelPath).async('string') : null;
            const templateNotesRelsXml = templateNotesRelPath ? await zip.file(`ppt/notesSlides/_rels/${templateNotesRelPath.split('/').pop()}.rels`).async('string') : null;

            // 2. CONTENT GENERATION
            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/i;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionText = sections[i].trim();
                const sName = `slide_gen_${i + 1}.xml`;
                
                let slideXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionText, userAlign);
                zip.file(`ppt/slides/${sName}`, slideXml);
                
                zip.file(`ppt/slides/_rels/${sName}.rels`, templateRelsXml.replace(/notesSlide\d+\.xml/g, `notes_gen_${i + 1}.xml`));

                if (templateNotesXml) {
                    const nName = `notes_gen_${i + 1}.xml`;
                    const noteLines = sectionText.split(/\n/).map(l => this.isChordLine(l) ? l.replace(this.chordRegex, m => `[${m.replace(/[\[\]]/g,'')}]`) : l);
                    const formattedNotes = this.escXml(noteLines.join('\n')).replace(/\n/g, `</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">`);
                    const newNotesContent = templateNotesXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r><a:rPr sz="1200"/><a:t xml:space="preserve">${formattedNotes}</a:t></a:r></a:p>`);
                    
                    zip.file(`ppt/notesSlides/${nName}`, newNotesContent);
                    zip.file(`ppt/notesSlides/_rels/${nName}.rels`, templateNotesRelsXml.replace(/slide\d+\.xml/g, sName));
                }
                generated.push({ id: 1000 + i, rid: `rIdGen${i + 1}`, name: sName });
            }

            // 3. FINAL REPAIR & SYNC
            await this.finalizeAndRepair(zip, presXml, presRelsXml, generated);
            
            const finalBlob = await zip.generateAsync({ type: 'blob', mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
            saveAs(finalBlob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    async finalizeAndRepair(zip, presXml, presRelsXml, generated) {
        const serializer = new XMLSerializer();
        const parser = new DOMParser();

        // A. Sync Slide ID List
        const presDoc = parser.parseFromString(presXml, 'application/xml');
        const sldIdLst = presDoc.getElementsByTagName('p:sldIdLst')[0];
        while (sldIdLst.firstChild) sldIdLst.removeChild(sldIdLst.firstChild);
        generated.forEach(s => {
            const node = presDoc.createElement('p:sldId');
            node.setAttribute('id', (256 + s.id).toString());
            node.setAttribute('r:id', s.rid);
            sldIdLst.appendChild(node);
        });
        zip.file('ppt/presentation.xml', serializer.serializeToString(presDoc));

        // B. Sync Presentation Relationships
        const relsDoc = parser.parseFromString(presRelsXml, 'application/xml');
        const rRoot = relsDoc.documentElement;
        [...rRoot.getElementsByTagName('Relationship')].forEach(r => r.getAttribute('Type').endsWith('slide') && r.remove());
        generated.forEach(s => {
            const e = relsDoc.createElement('Relationship');
            e.setAttribute('Id', s.rid);
            e.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            e.setAttribute('Target', `slides/${s.name}`);
            rRoot.appendChild(e);
        });
        zip.file('ppt/_rels/presentation.xml.rels', serializer.serializeToString(relsDoc));

        // C. Clean [Content_Types].xml
        const ctDoc = parser.parseFromString(await zip.file('[Content_Types].xml').async('string'), 'application/xml');
        const ctRoot = ctDoc.documentElement;
        [...ctRoot.getElementsByTagName('Override')].forEach(ov => {
            const pn = ov.getAttribute('PartName');
            if (pn.includes('/ppt/slides/') || pn.includes('/ppt/notesSlides/')) ov.remove();
        });
        generated.forEach(s => {
            const sEl = ctDoc.createElement('Override');
            sEl.setAttribute('PartName', `/ppt/slides/${s.name}`);
            sEl.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            ctRoot.appendChild(sEl);
            const nName = s.name.replace('slide_gen_', 'notes_gen_');
            if (zip.file(`ppt/notesSlides/${nName}`)) {
                const nEl = ctDoc.createElement('Override');
                nEl.setAttribute('PartName', `/ppt/notesSlides/${nName}`);
                nEl.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
                ctRoot.appendChild(nEl);
            }
        });
        zip.file('[Content_Types].xml', serializer.serializeToString(ctDoc));

        // D. Update Metadata
        if (zip.file('docProps/app.xml')) {
            let appXml = await zip.file('docProps/app.xml').async('string');
            appXml = appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${generated.length}</Slides>`)
                           .replace(/<I4>\d+<\/I4>/, `<I4>${generated.length}</I4>`);
            zip.file('docProps/app.xml', appXml);
        }
    },

    // --- TRANSPOSITION ENGINE ---
    async transpose() {
        const file = this.elements.transFileInput.files[0];
        const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;
        if (!file || this.originalSlides.length === 0) return alert('Select file and wait for load.');

        try {
            this.showLoading('Transposing...');
            const zip = await JSZip.loadAsync(file);
            for (const slide of this.originalSlides) {
                let slideXml = await zip.file(slide.path).async('string');
                slideXml = this.transposeParagraphs(slideXml, semitones);
                zip.file(slide.path, slideXml);
                if (slide.notesPath) {
                    const transposedNotes = this.transposeLine(slide.notes, semitones);
                    let notesXml = await zip.file(slide.notesPath).async('string');
                    const style = '<a:rPr lang="en-US" sz="1200"/>';
                    const formatted = this.escXml(transposedNotes).replace(/\n/g, `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                    notesXml = notesXml.replace(/<a:p>[\s\S]*?<\/a:p>/, `<a:p><a:r>${style}<a:t xml:space="preserve">${formatted}</a:t></a:r></a:p>`);
                    zip.file(slide.notesPath, notesXml);
                }
            }
            saveAs(await zip.generateAsync({ type: 'blob' }), file.name.replace('.pptx', `_transposed.pptx`));
            this.hideLoading();
        } catch (err) { alert(err.message); this.hideLoading(); }
    },

    transposeParagraphs(xml, semitones) {
        return xml.replace(/<a:p[^>]*>([\s\S]*?)<\/a:p>/g, (matchFull, pXml) => {
            let logicLine = "", charMeta = []; 
            const runRegex = /<a:r>([\s\S]*?)<\/a:r>|<a:br\/>/g;
            let m;
            while ((m = runRegex.exec(pXml)) !== null) {
                if (m[0] === '<a:br/>') { logicLine += "\n"; charMeta.push({ isBr: true }); continue; }
                const rStyle = m[1].match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr/>';
                const isGhost = rStyle.includes('<a:noFill/>');
                const text = this.unescXml(m[1].match(/<a:t[^>]*>(.*?)<\/a:t>/)?.[1] || "");
                for (let char of text) { charMeta.push({ isGhost, originalChar: char, style: rStyle, isBr: false }); logicLine += isGhost ? " " : char; }
            }
            if (!logicLine.trim()) return matchFull;
            const transposedLogic = this.transposeLine(logicLine, semitones);
            if (transposedLogic === logicLine) return matchFull;
            const pPr = pXml.match(/<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/)?.[0] || '';
            const pTagOpen = matchFull.match(/^<a:p[^>]*>/)?.[0] || '<a:p>';
            let newRuns = "", metaIdx = 0;
            for (let i = 0; i < transposedLogic.length; i++) {
                const newChar = transposedLogic[i];
                if (newChar === "\n") { newRuns += "<a:br/>"; while(metaIdx < charMeta.length && !charMeta[metaIdx].isBr) metaIdx++; metaIdx++; continue; }
                const meta = charMeta[metaIdx] || { isGhost: false, style: '<a:rPr sz="1800"/>' };
                if (meta.isGhost && newChar === " ") { newRuns += `<a:r>${meta.style}<a:t xml:space="preserve">${this.escXml(meta.originalChar).replace(/ /g, '\u00A0')}</a:t></a:r>`; }
                else { newRuns += `<a:r>${this.getChordStyle(meta.style).replace('<a:noFill/>', '')}<a:t xml:space="preserve">${this.escXml(newChar).replace(/ /g, '\u00A0')}</a:t></a:r>`; }
                metaIdx++;
            }
            return `${pTagOpen}${pPr}${newRuns}</a:p>`;
        });
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        return text.split('\n').map(line => {
            if (!this.isChordLine(line)) return line;
            let result = line, offset = 0;
            const matches = [...line.matchAll(this.chordRegex)];
            for (const m of matches) {
                const nr = this.shiftNote(m[1], semitones); 
                const nb = m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '';
                const nf = nr + (m[2] || '') + nb;
                const p = m.index + offset, d = nf.length - m[0].length;
                let pre = result.substring(0, p), suf = result.substring(p + m[0].length);
                if (d > 0 && suf.startsWith(' ')) { suf = suf.substring(1); offset--; }
                else if (d < 0) { suf = ' '.repeat(Math.abs(d)) + suf; offset += Math.abs(d); }
                result = pre + nf + suf; offset += d;
            }
            return result;
        }).join('\n');
    },

    shiftNote(note, semitones) {
        let idx = this.musical.keys.indexOf(note);
        if (idx === -1) idx = this.musical.flats.indexOf(note);
        if (idx === -1) return note;
        let newIdx = (idx + semitones) % 12;
        if (newIdx < 0) newIdx += 12;
        return this.musical.preferred[newIdx];
    },

    // --- XML REPLACEMENT ENGINE ---
    lockInStyleAndReplace(xml, ph, replacement, align = 'ctr') {
        const phRegex = new RegExp(this.getPlaceholderRegexStr(ph), 'gi');
        return xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (shape) => {
            if (!phRegex.test(shape)) return shape;
            const style = shape.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] || '<a:rPr lang="en-US"/>';
            if (!ph.toLowerCase().includes('lyrics')) {
                const escaped = replacement.split('\n').map(l => this.escXml(l)).join(`</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`);
                return shape.replace(phRegex, escaped);
            }
            let injected = `</a:t></a:r></a:p>`;
            const rawLines = replacement.split('\n');
            for (let i = 0; i < rawLines.length; i++) {
                let line = rawLines[i], next = rawLines[i+1];
                if (this.isChordLine(line) && next && !this.isChordLine(next) && !next.trim().startsWith('[')) {
                    const max = Math.max(line.length, next.length);
                    injected += (align === 'ctr') ? this.makeGhostAlignmentLine(line.padEnd(max,' '), next.padEnd(max,' '), style, 'ctr') + this.makePptLine(next.padEnd(max,' '), style, 'ctr') 
                                                : this.makePptLine(line, this.getChordStyle(style), 'l') + this.makePptLine(next, style, 'l');
                    i++;
                } else {
                    const text = line.trim(), isTag = text.startsWith('[') && text.endsWith(']');
                    let curStyle = isTag ? style.replace(/sz="\d+"/, 'sz="2000"') : style;
                    injected += text ? this.makePptLine(text, curStyle, align === 'ctr' ? 'ctr' : 'l') : `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t> </a:t></a:r></a:p>`;
                }
            }
            return shape.replace(phRegex, injected + `<a:p><a:pPr algn="${align === 'ctr' ? 'ctr' : 'l'}"/><a:r>${style}<a:t xml:space="preserve">`)
                        .replace('</a:bodyPr>', '<a:normAutofit fontScale="92000" lnSpcReduction="10000"/></a:bodyPr>');
        });
    },

    // --- HELPERS ---
    makeGhostAlignmentLine(c, l, s, a) {
        let gh = s.replace('<a:rPr', '<a:rPr><a:noFill/>').replace(/<a:solidFill>.*?<\/a:solidFill>/g, '');
        let xml = "";
        for (let i = 0; i < c.length; i++) xml += (c[i] === ' ') ? `<a:r>${gh}<a:t xml:space="preserve">${this.escXml(l[i] || ' ')}</a:t></a:r>` : `<a:r>${this.getChordStyle(s)}<a:t xml:space="preserve">${this.escXml(c[i])}</a:t></a:r>`;
        return `<a:p><a:pPr algn="${a}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr>${xml}</a:p>`;
    },
    makePptLine(t, s, a) { return `<a:p><a:pPr algn="${a}"><a:lnSpc><a:spcPct val="50000"/></a:lnSpc></a:pPr><a:r>${s}<a:t xml:space="preserve">${this.escXml(t)}</a:t></a:r></a:p>`; },
    getPlaceholderRegexStr(ph) { return '\\[' + ph.replace(/[\[\]]/g, '').split('').map(c => (c === ' ' ? '\\s+' : this.escRegex(c))).join('(?:<[^>]+>|\\s)*') + '\\]'; },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    getNotesRelPath(rXml) { const m = rXml?.match(/Type="[^"]+notesSlide"[^>]+Target="..\/notesSlides\/(notesSlide\d+\.xml)"/); return m ? `ppt/notesSlides/${m[1]}` : null; },
    isChordLine(l) { if(!l) return false; const t = l.trim(), w = t.split(/\s+/), c = t.match(this.chordRegex) || []; return c.length >= w.length * 0.5 || (c.length > 0 && w.length <= 2); },
    getChordStyle(s) { const f = '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>'; let res = s.includes('sz=') ? s.replace(/sz="\d+"/, 'sz="1800"') : s.replace('<a:rPr', '<a:rPr sz="1800"'); return res.includes('<a:solidFill>') ? res.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, f) : res.replace('</a:rPr>', f + '</a:rPr>'); },
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; }
};

App.init();
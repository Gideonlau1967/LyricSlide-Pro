/* LyricSlide Pro - Core Logic v20 (Monospaced Block-Lock Edition) */

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
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
        console.log("App Initialized. v20.0 - Monospaced Right-Padding Lock");
    },

    // --- REPLACEMENT ENGINE (THE BLOCK-LOCK) ---
    lockInStyleAndReplace(xml, placeholder, replacement) {
        const phRegexStr = this.getPlaceholderRegexStr(placeholder);
        const phRegex = new RegExp(phRegexStr, 'gi');

        return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
            if (phRegex.test(shapeXml)) {
                // Get style from template
                const rPrMatch = shapeXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g);
                const defRPrMatch = shapeXml.match(/<a:defRPr[^>]*>[\s\S]*?<\/a:defRPr>/g);
                let style = (rPrMatch ? rPrMatch[0] : (defRPrMatch ? defRPrMatch[0].replace('defRPr', 'rPr') : '<a:rPr lang="en-US"/>'));

                let rawLines = (replacement || '').split(/\r?\n/);
                
                if (placeholder === '[Lyrics and Chords]') {
                    // 1. Find the longest line in this section to be the anchor
                    const maxLen = Math.max(...rawLines.map(l => l.length));
                    
                    // 2. Rigid Block Logic: Pad right with Non-Breaking Spaces
                    // We also convert internal spaces to NBSPs to ensure mathematical alignment in Monospaced fonts
                    rawLines = rawLines.map(l => {
                        const paddingCount = maxLen - l.length;
                        const core = l.replace(/ /g, '\u00A0'); 
                        return core + '\u00A0'.repeat(paddingCount);
                    });
                }

                const lines = rawLines.map(l => this.escXml(l));
                let injected = '';
                lines.forEach((line, idx) => {
                    if (idx > 0) injected += `</a:t></a:r><a:br/><a:r>${style}<a:t xml:space="preserve">`;
                    injected += line;
                });

                // Handle Autofit Font Shrink for long sections
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

                // Clean up empty tags
                result = result.replace(/<a:t xml:space="preserve"><\/a:t>/g, '').replace(/<a:r><a:rPr[^>]*><a:t xml:space="preserve"><\/a:t><\/a:r>/g, '');
                
                // Ensure Autofit is enabled so text stays inside the box
                if (!result.includes('Autofit')) {
                    result = result.replace('</a:bodyPr>', '<a:normAutofit fontScale="85000" lnSpcReduction="15000"/></a:bodyPr>');
                }

                return result;
            }
            return shapeXml;
        });
    },

    // --- PPTX GENERATION ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';

        if (!file || !lyrics) return alert('Please select a template and enter lyrics.');

        try {
            this.showLoading('Locking Chords & Generating...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            
            // Split lyrics by sections like [Verse] or [Chorus]
            const sections = ("\n" + lyrics).split(/\r?\n(?=\s*\[[^\]]+\])/).filter(s => s.trim() !== '');
            
            const newZip = zip;
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                let sXml = this.lockInStyleAndReplace(templateXml, '[Title]', title);
                sXml = this.lockInStyleAndReplace(sXml, '[Copyright Info]', copyright);
                sXml = this.lockInStyleAndReplace(sXml, '[Lyrics and Chords]', sections[i].trim());
                
                const name = `slide_gen_${i+1}.xml`;
                newZip.file(`ppt/slides/${name}`, sXml);
                generated.push({ id: 5000 + i, rid: `rIdGen${i+1}`, name, path: `ppt/slides/${name}` });
            }

            this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);
            const blob = await newZip.generateAsync({ type: 'blob' });
            saveAs(blob, `${title.replace(/[^a-z0-9]/gi, '_') || 'Song'}.pptx`);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error: " + err.message);
            this.hideLoading();
        }
    },

    // --- HELPERS ---
    getPlaceholderRegexStr(ph) {
        const inner = ph.replace(/[\[\]]/g, '').trim();
        return '\\[' + '(?:<[^>]+>|\\s)*' + inner.split('').map((p, i) => (p === ' ' ? '\\s+' : this.escRegex(p)) + (i < inner.length - 1 ? '(?:<[^>]+>|\\s)*' : '')).join('') + '(?:<[^>]+>|\\s)*' + '\\]';
    },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    escXml(s) { return (s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    unescXml(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    
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

    loadDefaultTemplates() {
        fetch('./templates.json').then(r => r.json()).then(names => {
            const gallery = document.getElementById('templateGallery');
            const grid = document.createElement('div'); grid.className = 'template-grid';
            names.forEach(name => {
                const card = document.createElement('div'); card.className = 'template-card';
                card.innerHTML = `<img class="template-thumb" src="${name.replace(/\.pptx$/i, '.png')}" onerror="this.src='https://placehold.co/200x120?text=PPTX'"><div class="template-card-name">${name.replace(/\.pptx$/i, '')}</div>`;
                card.onclick = () => {
                    fetch(`./${encodeURIComponent(name)}`).then(r => r.blob()).then(b => {
                        this.selectedTemplateFile = new File([b], name);
                        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                        card.classList.add('selected');
                        document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                        document.getElementById('selectedTemplateName').textContent = name;
                    });
                };
                grid.appendChild(card);
            });
            gallery.appendChild(grid);
        }).catch(e => console.log("Templates not found locally."));
    },

    transposeLine(text, semitones) {
        if (semitones === 0) return text;
        return text.split('\n').map(line => {
            const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
            const matches = [...line.matchAll(chordRegex)];
            if (matches.length === 0) return line;
            let res = line, off = 0;
            for (const m of matches) {
                const root = this.shiftNote(m[1], semitones);
                const bass = m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '';
                const newC = root + (m[2] || '') + bass;
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

    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    
    transpose() { /* Existing transposition logic call */ },
    theme: { init() {} }
};

App.init();
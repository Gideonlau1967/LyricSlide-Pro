/* LyricSlide Pro - generate.js (v21.7 - Two-Column & Selection UI) */

const App = {
    version: "v21.7",

    elements: {
        songTitle: document.getElementById('songTitle'),
        lyricsInput: document.getElementById('lyricsInput'),
        copyrightInfo: document.getElementById('copyrightInfo'),
        generateBtn: document.getElementById('generateBtn'),
        // Ensure you have an element with this ID in your HTML
        selectedTemplateDisplay: document.getElementById('selectedTemplateName'), 
        loadingOverlay: document.getElementById('loadingOverlay'),
        loadingText: document.getElementById('loadingText')
    },

    selectedTemplateFile: null, 

    init() {
        const verEl = document.getElementById('appVersion');
        if (verEl) verEl.textContent = this.version;
        if (this.elements.generateBtn) this.elements.generateBtn.onclick = () => this.generate();
        this.theme.init();
        this.loadDefaultTemplates(); 
        window.LyricApp = this;
    },

    theme: {
        defaults: { '--primary-color': '#334155', '--bg-start': '#f8fafc', '--bg-end': '#f8fafc', '--text-main': '#1e293b', '--card-accent': '#e2e8f0', '--preview-card-bg': '#ffffff', '--preview-chord-color': '#334155', '--preview-lyrics-color': '#1e293b' },
        init() {
            const saved = JSON.parse(localStorage.getItem('lyric_theme') || '{}');
            Object.keys(this.defaults).forEach(key => {
                const val = saved[key] || this.defaults[key];
                document.documentElement.style.setProperty(key, val);
            });
        }
    },

    // --- TEMPLATE LOADING WITH UI FEEDBACK ---
    async loadDefaultTemplates() {
        const gallery = document.getElementById('templateGallery');
        if (!gallery) return;

        try {
            const res = await fetch('./templates.json');
            const names = await res.json();
            gallery.innerHTML = '';

            names.forEach(name => {
                const card = document.createElement('div');
                card.className = 'template-card';
                
                // Construct thumbnail image path (assuming .png exists for every .pptx)
                const thumbSrc = name.replace('.pptx', '.png');
                
                card.innerHTML = `
                    <img class="template-thumb" src="${thumbSrc}" onerror="this.src='data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 60%22><rect width=%22100%22 height=%2260%22 fill=%22%23ccc%22/><text x=%2250%%22 y=%2250%%22 text-anchor=%22middle%22 font-family=%22sans-serif%22 font-size=%228%22 fill=%22%23666%22>No Preview</text></svg>'">
                    <div class="template-card-name" style="font-size: 12px; padding: 5px;">${name.replace('.pptx','')}</div>
                `;

                card.onclick = async () => {
                    // 1. UI Highlight
                    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                    card.classList.add('selected');

                    // 2. Update Selection Text
                    const cleanName = name.replace('.pptx', '');
                    if (this.elements.selectedTemplateDisplay) {
                        this.elements.selectedTemplateDisplay.innerHTML = `Selected: <strong>${cleanName}</strong>`;
                        this.elements.selectedTemplateDisplay.style.color = "#3b82f6";
                    }

                    // 3. Load File Data
                    try {
                        this.showLoading(`Loading ${cleanName}...`);
                        const r = await fetch(`./${encodeURIComponent(name)}`);
                        if (!r.ok) throw new Error("File not found");
                        const blob = await r.blob();
                        this.selectedTemplateFile = blob;
                        this.hideLoading();
                        console.log("Template Ready:", name);
                    } catch (err) {
                        this.hideLoading();
                        alert("Failed to load template file.");
                    }
                };
                gallery.appendChild(card);
            });
        } catch (e) { 
            gallery.innerHTML = '<div style="grid-column: 1/3; text-align: center; padding: 20px;">Could not load template list.</div>'; 
        }
    },

    // --- GENERATION ENGINE (v21.6 LOGIC PRESERVED) ---
    async generate() {
        if (!this.selectedTemplateFile || !this.elements.lyricsInput.value) return alert('Select a template and enter lyrics first.');
        try {
            this.showLoading('Generating PPTX...');
            const zip = await JSZip.loadAsync(this.selectedTemplateFile);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideRels = this.getSlideRels(presRelsXml);
            const slideIds = this.getSlideIds(presXml);
            
            const templateRelPath = slideRels[slideIds[0].rid];
            const templateXml = await zip.file(`ppt/${templateRelPath}`).async('string');
            const templateRelsXml = await zip.file(`ppt/slides/_rels/${templateRelPath.split('/').pop()}.rels`).async('string');
            
            let rawInput = this.elements.lyricsInput.value.trim();
            let sections = rawInput.split(/\r?\n(?=\s*\[(?!(?:TITLE|Copyright Info|LYRICS AND CHORDS)\])[^\]\n]+\])/).filter(s => s.trim() !== '');
            if (sections.length === 0) sections = [rawInput];

            const generated = [];
            const copyrightText = (this.elements.copyrightInfo.value ? this.elements.copyrightInfo.value + " | " : "") + "Generated by " + this.version;

            for (let i = 0; i < sections.length; i++) {
                let slideXml = this.lockInStyleAndReplace(templateXml, '[ TITLE ]', this.elements.songTitle.value);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyrightText);
                slideXml = this.lockInStyleAndReplace(slideXml, '[LYRICS AND CHORDS]', sections[i].trim());

                const name = `song_gen_${i + 1}.xml`;
                zip.file(`ppt/slides/${name}`, slideXml);
                zip.file(`ppt/slides/_rels/${name}.rels`, templateRelsXml);
                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path: `ppt/slides/${name}` });
            }

            this.syncPresentationRegistry(zip, presXml, presRelsXml, generated);
            const outName = (this.elements.songTitle.value || 'Song').replace(/[^a-z0-9]/gi, '_') + '.pptx';
            saveAs(await zip.generateAsync({ type: 'blob' }), outName);
            this.hideLoading();
        } catch (e) { 
            console.error(e);
            this.hideLoading(); 
            alert("Error during generation."); 
        }
    },

    lockInStyleAndReplace(xml, placeholder, replacement) {
        const createFuzzyRegex = (ph) => {
            const chars = ph.split('');
            const fuzzy = chars.map(c => {
                const escaped = this.escRegex(c);
                return escaped === '\\ ' ? '\\s*' : `${escaped}(?:<[^>]+>)*`;
            }).join('(?:<[^>]+>)*');
            return new RegExp(fuzzy, 'gi');
        };

        const phRegex = createFuzzyRegex(placeholder);
        const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;

        return xml.replace(/<(p:sp|p:graphicFrame)>([\s\S]*?)<\/\1>/g, (fullFrame, tagName, innerContent) => {
            phRegex.lastIndex = 0;
            if (phRegex.test(innerContent)) {
                const latinMatch = innerContent.match(/<a:latin typeface="([^"]+)"/);
                const templateFont = latinMatch ? latinMatch[1] : "Arial";
                const sizeMatch = innerContent.match(/sz="(\d+)"/);
                const templateSize = sizeMatch ? sizeMatch[1] : "2400"; 

                if (!/LYRICS/i.test(placeholder)) {
                    const rPrMatch = innerContent.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/);
                    let style = (rPrMatch ? rPrMatch[0] : '<a:rPr lang="en-US"/>');
                    const lines = (replacement || '').split(/\r?\n/);
                    const escapedText = lines.map(l => `<a:r>${style}<a:t xml:space="preserve">${this.escXml(l)}</a:t></a:r>`).join('<a:br/>');
                    return `<${tagName}>${innerContent.replace(phRegex, escapedText)}</${tagName}>`;
                }

                const lines = (replacement || '').split(/\r?\n/);
                let tableRowsXml = '';
                lines.forEach((line) => {
                    let trimmed = line.trim();
                    if (trimmed === '') { 
                        tableRowsXml += this.createTableCellXml(" ", templateFont, templateSize, "ctr", 150000); 
                    } else if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
                        tableRowsXml += this.createTableCellXml(this.escXml(trimmed), templateFont, Math.round(templateSize * 0.8), "ctr", 350000);
                    } else if (line.match(chordRegex)) {
                        const esc = this.escXml(line).replace(/ /g, '&#160;');
                        tableRowsXml += this.createTableCellXml(esc, "Courier New", "1600", "ctr", 350000);
                    } else {
                        tableRowsXml += this.createTableCellXml(this.escXml(line), templateFont, templateSize, "ctr", 400000);
                    }
                });

                const tblStartIdx = innerContent.indexOf('<a:tbl>');
                const gridEndIdx = innerContent.indexOf('</a:tblGrid>') + 12;
                const tblEndIdx = innerContent.lastIndexOf('</a:tbl>');
                
                if (tblStartIdx > -1 && gridEndIdx > 11) {
                    let header = innerContent.substring(tblStartIdx, gridEndIdx);
                    let footer = innerContent.substring(tblEndIdx);
                    header = header.replace(/<a:tableStyleId>[\s\S]*?<\/a:tableStyleId>/, '<a:tableStyleId>{5C22544A-7EE6-4342-B051-7303C2061113}</a:tableStyleId>');
                    if (header.includes('</a:tblPr>')) {
                        header = header.replace('</a:tblPr>', '<a:wholeTbl><a:tcPr><a:noFill/></a:tcPr></a:wholeTbl></a:tblPr>');
                    }
                    return `<${tagName}>${innerContent.substring(0, tblStartIdx)}${header}${tableRowsXml}${footer}</${tagName}>`;
                }
            }
            return fullFrame;
        });
    },

    createTableCellXml(text, font, size, align, height) {
        return `<a:tr h="${height}"><a:tc><a:txBody><a:bodyPr vert="ctr" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/><a:p><a:pPr algn="${align}"/><a:r><a:rPr sz="${size}" lang="en-US"><a:latin typeface="${font}"/><a:cs typeface="${font}"/></a:r><a:t xml:space="preserve">${text}</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB><a:noFill/></a:tcPr></a:tc></a:tr>`;
    },

    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    escXml(s) { return (s||'').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c])); },
    escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); },
    getSlideIds(xml) { let ids = [], m, r = /<p:sldId[^>]+id="([^"]+)"[^>]+r:id="([^"]+)"/g; while (m = r.exec(xml)) ids.push({id: m[1], rid: m[2]}); return ids; },
    getSlideRels(xml) { let rels = {}, m, r = /<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]+slide"[^>]+Target="([^"]+)"/g; while (m = r.exec(xml)) rels[m[1]] = m[2]; return rels; },
    syncPresentationRegistry(zip, pres, rels, gen) {
        const list = '<p:sldIdLst>' + gen.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        zip.file('ppt/presentation.xml', pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, list));
        let doc = new DOMParser().parseFromString(rels, 'application/xml');
        let rs = doc.getElementsByTagName('Relationship');
        for (let j = rs.length - 1; j >= 0; j--) if (rs[j].getAttribute('Type').endsWith('slide')) rs[j].parentNode.removeChild(rs[j]);
        gen.forEach(s => { let el = doc.createElement('Relationship'); el.setAttribute('Id', s.rid); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'); el.setAttribute('Target', `slides/${s.name}`); doc.documentElement.appendChild(el); });
        zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(doc));
        const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>';
        let entries = gen.map(s => `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
        zip.file('[Content_Types].xml', (head + entries + '</Types>').replace('><Override', '><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'));
    }
};

App.init();
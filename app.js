/**
 * LyricSlide Pro - v15.5 (Multi-Box Edition)
 * Targets: {{TITLE}}, {{LYRICS}}, {{COPYRIGHT}}
 */

const App = {
    elements: {
        songTitle: document.getElementById('songTitle'),
        lyricsInput: document.getElementById('lyricsInput'),
        copyrightInfo: document.getElementById('copyrightInfo'),
        generateBtn: document.getElementById('generateBtn'),
        loadingOverlay: document.getElementById('loadingOverlay'),
        loadingText: document.getElementById('loadingText')
    },

    selectedTemplateFile: null,

    init() {
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.loadTemplatesFromDirectory();
        window.LyricApp = this;
    },

    isChordLine(line) {
        const trimmed = line.trim();
        if (!trimmed) return false;
        const chordRegex = /^[A-G][b#]?(m|maj|min|dim|aug|sus|add|2|4|5|6|7|9|11|13|[\+\-\^\(\)])?(\/[A-G][b#]?)?(\s+[A-G][b#]?(m|maj|min|dim|aug|sus|add|2|4|5|6|7|9|11|13)?(\/[A-G][b#]?)?)*$/i;
        return chordRegex.test(trimmed);
    },

    // Helper to create a PowerPoint XML paragraph
    createP(text, align = 'ctr') {
        const safeText = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return `
            <a:p>
                <a:pPr algn="${align}"><a:buNone/></a:pPr>
                <a:r>
                    <a:rPr lang="en-US" dirty="0" smtClean="0" />
                    <a:t>${safeText}</a:t>
                </a:r>
            </a:p>`;
    },

    async generate() {
        if (!this.selectedTemplateFile) {
            alert("Please select a template first.");
            return;
        }

        this.showLoading("Injecting Content...");

        try {
            const zip = await JSZip.loadAsync(this.selectedTemplateFile);
            let slideXml = await zip.file("ppt/slides/slide1.xml").async("string");

            // --- 1. PROCESS LYRICS & CHORDS ---
            const lyricLines = this.elements.lyricsInput.value.split('\n');
            let lyricXmlContent = "";
            lyricLines.forEach(line => {
                if (line.trim().startsWith('[')) return; // Ignore headers like [Chorus]
                const align = this.isChordLine(line) ? 'l' : 'ctr';
                lyricXmlContent += this.createP(line, align);
            });

            // --- 2. PROCESS TITLE & COPYRIGHT ---
            const titleXml = this.createP(this.elements.songTitle.value || "", 'ctr');
            const copyrightXml = this.createP(this.elements.copyrightInfo.value || "", 'ctr');

            // --- 3. XML INJECTION (REPLACING PLACEHOLDERS) ---
            // This regex finds the entire paragraph <a:p>...</a:p> containing the placeholder
            // and replaces it with our new generated XML content.
            
            const findAndReplace = (xml, tag, newContent) => {
                // Regex matches a paragraph containing the tag
                const re = new RegExp(`<a:p>(?:(?!<a:p>).)*?${tag}[\\s\\S]*?<\\/a:p>`, 'g');
                return xml.replace(re, newContent);
            };

            slideXml = findAndReplace(slideXml, '{{TITLE}}', titleXml);
            slideXml = findAndReplace(slideXml, '{{COPYRIGHT}}', copyrightXml);
            slideXml = findAndReplace(slideXml, '{{LYRICS}}', lyricXmlContent);

            // Save back to ZIP
            zip.file("ppt/slides/slide1.xml", slideXml);

            // --- 4. DOWNLOAD ---
            const content = await zip.generateAsync({ type: "blob" });
            const safeName = (this.elements.songTitle.value || "Song").replace(/[^a-z0-9]/gi, '_');
            saveAs(content, `${safeName}.pptx`);

        } catch (error) {
            console.error(error);
            alert("Error: " + error.message);
        } finally {
            this.hideLoading();
        }
    },

    // --- GALLERY LOADING (GitHub Relative Paths) ---
    async loadTemplatesFromDirectory() {
        try {
            const response = await fetch('./templates.json');
            const filenames = await response.json();
            const galleryData = filenames.map(name => {
                const baseName = name.replace(/\.pptx$/i, '');
                return {
                    name: name,
                    thumbUrl: `./${encodeURIComponent(baseName)}.png`, 
                    getFile: async () => {
                        const res = await fetch(`./${encodeURIComponent(name)}`);
                        const blob = await res.blob();
                        return new File([blob], name);
                    }
                };
            });
            this.renderTemplateGallery(galleryData);
        } catch (e) {
            console.warn("Gallery load error:", e);
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
                this.showLoading('Loading...');
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

    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; }
};

App.init();
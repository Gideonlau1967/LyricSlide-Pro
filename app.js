/**
 * LyricSlide Pro - v15.6 (Smart Alignment Edition)
 * Preserves template alignment for Title/Copyright.
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

    // --- DETECTION LOGIC ---
    isChordLine(line) {
        const trimmed = line.trim();
        if (!trimmed) return false;
        // Musical chord detection (A-G with modifiers)
        const chordRegex = /^[A-G][b#]?(m|maj|min|dim|aug|sus|add|2|4|5|6|7|9|11|13|[\+\-\^\(\)])?(\/[A-G][b#]?)?(\s+[A-G][b#]?(m|maj|min|dim|aug|sus|add|2|4|5|6|7|9|11|13)?(\/[A-G][b#]?)?)*$/i;
        return chordRegex.test(trimmed);
    },

    // Helper to create PowerPoint Paragraph XML
    createP(text, align = 'l') {
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

    // Helper to extract alignment from an existing paragraph string
    getExistingAlignment(pXml) {
        const match = pXml.match(/algn="([^"]+)"/);
        return match ? match[1] : 'l'; // Default to left if not found
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

            // --- 1. PREPARE REPLACEMENTS ---

            // A. Title: Find template alignment then replace
            slideXml = this.replaceWithTemplateAlignment(slideXml, '{{TITLE}}', this.elements.songTitle.value);

            // B. Copyright: Find template alignment then replace
            slideXml = this.replaceWithTemplateAlignment(slideXml, '{{COPYRIGHT}}', this.elements.copyrightInfo.value);

            // C. Lyrics: Special Logic (Left chords, Center lyrics)
            const lyricLines = this.elements.lyricsInput.value.split('\n');
            let lyricXmlContent = "";
            lyricLines.forEach(line => {
                if (line.trim().startsWith('[')) return; 
                const align = this.isChordLine(line) ? 'l' : 'ctr';
                lyricXmlContent += this.createP(line, align);
            });
            
            // Replace the lyric block paragraph entirely
            const lyricRe = new RegExp(`<a:p>(?:(?!<a:p>).)*?{{LYRICS}}[\\s\\S]*?<\\/a:p>`, 'g');
            slideXml = slideXml.replace(lyricRe, lyricXmlContent);

            // --- 2. FINALIZE ---
            zip.file("ppt/slides/slide1.xml", slideXml);
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

    // Dynamic replacement that reads the alignment of the paragraph containing the tag
    replaceWithTemplateAlignment(xml, tag, value) {
        // Find the specific paragraph containing the placeholder tag
        const pRegex = new RegExp(`(<a:p>(?:(?!<a:p>).)*?${tag}[\\s\\S]*?<\\/a:p>)`, 'g');
        const match = xml.match(pRegex);
        
        if (match) {
            const originalP = match[0];
            const align = this.getExistingAlignment(originalP);
            const newP = this.createP(value || "", align);
            return xml.replace(originalP, newP);
        }
        return xml;
    },

    // --- GALLERY LOADING ---
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
        } catch (e) { console.warn("Load error", e); }
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
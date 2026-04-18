/**
 * LyricSlide Pro - v15.7 (Perfect Format Edition)
 * Preserves Font Size, Color, and Position from Template.
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

    async generate() {
        if (!this.selectedTemplateFile) {
            alert("Please select a template first.");
            return;
        }

        this.showLoading("Injecting Content...");

        try {
            const zip = await JSZip.loadAsync(this.selectedTemplateFile);
            let slideXml = await zip.file("ppt/slides/slide1.xml").async("string");

            // --- 1. REPLACE TITLE & COPYRIGHT (Keep exact Run Properties) ---
            // This replaces the text inside the existing <a:t> tag without touching the <a:rPr> (styling)
            slideXml = slideXml.replace(/{{TITLE}}/g, (this.elements.songTitle.value || "").replace(/&/g, '&amp;'));
            slideXml = slideXml.replace(/{{COPYRIGHT}}/g, (this.elements.copyrightInfo.value || "").replace(/&/g, '&amp;'));

            // --- 2. PROCESS LYRICS (Clone Style per Line) ---
            
            // Find the paragraph containing {{LYRICS}} to steal its style
            const lyricParaRegex = /(<a:p>[\s\S]*?{{LYRICS}}[\s\S]*?<\/a:p>)/;
            const match = slideXml.match(lyricParaRegex);

            if (match) {
                const templatePara = match[0];
                
                // Extract the specific style (Run Properties <a:rPr>) from the template paragraph
                const rPrMatch = templatePara.match(/<a:rPr[\s\S]*?\/?>/);
                const rPr = rPrMatch ? rPrMatch[0] : '<a:rPr lang="en-US" />';

                const lyricLines = this.elements.lyricsInput.value.split('\n');
                let newLyricXml = "";

                lyricLines.forEach(line => {
                    if (line.trim().startsWith('[')) return; // Skip headers
                    
                    const align = this.isChordLine(line) ? 'l' : 'ctr';
                    const safeText = line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

                    // Construct new paragraph using the template's font properties
                    newLyricXml += `
                        <a:p>
                            <a:pPr algn="${align}"><a:buNone/></a:pPr>
                            <a:r>
                                ${rPr}
                                <a:t>${safeText}</a:t>
                            </a:r>
                        </a:p>`;
                });

                // Swap the single {{LYRICS}} paragraph with the stack of new paragraphs
                slideXml = slideXml.replace(templatePara, newLyricXml);
            }

            // Save back
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

    // --- GALLERY LOADING (Standard) ---
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
        } catch (e) { console.warn(e); }
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
            thumb.onerror = () => { thumb.replaceWith(document.createElement('div')); };
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
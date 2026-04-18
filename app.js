/**
 * LyricSlide Pro - v15.4 (GitHub Edition)
 * Feature: Left-aligned chords, Center-aligned lyrics
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
        console.log("App Init: Starting...");
        this.elements.generateBtn.addEventListener('click', () => this.generate());
        this.loadTemplatesFromDirectory();
        window.LyricApp = this;
    },

    // --- DETECTION LOGIC ---
    isChordLine(line) {
        const trimmed = line.trim();
        if (!trimmed) return false;
        // Regex for musical chords
        const chordRegex = /^[A-G][b#]?(m|maj|min|dim|aug|sus|add|2|4|5|6|7|9|11|13|[\+\-\^\(\)])?(\/[A-G][b#]?)?(\s+[A-G][b#]?(m|maj|min|dim|aug|sus|add|2|4|5|6|7|9|11|13)?(\/[A-G][b#]?)?)*$/i;
        return chordRegex.test(trimmed);
    },

    // --- GENERATION LOGIC ---
    async generate() {
        console.log("Generate Clicked");
        
        if (!this.selectedTemplateFile) {
            alert("Error: Please select a template from the library first!");
            return;
        }

        const rawInput = this.elements.lyricsInput.value.trim();
        if (!rawInput) {
            alert("Error: Please enter some lyrics!");
            return;
        }

        this.showLoading("Processing XML...");

        try {
            console.log("Loading ZIP...");
            const zip = await JSZip.loadAsync(this.selectedTemplateFile);
            
            // 1. Target slide1.xml
            let slideXml = await zip.file("ppt/slides/slide1.xml").async("string");
            console.log("Slide XML loaded");

            // 2. Process Lyrics into Paragraphs
            const lines = rawInput.split('\n');
            let generatedXml = "";

            lines.forEach(line => {
                if (line.trim().startsWith('[')) return; // Skip headers

                const isChord = this.isChordLine(line);
                const alignment = isChord ? 'l' : 'ctr';
                const safeLine = line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

                generatedXml += `
                <a:p>
                    <a:pPr algn="${alignment}"><a:buNone/></a:pPr>
                    <a:r>
                        <a:rPr lang="en-US" dirty="0" smtClean="0" />
                        <a:t>${safeLine}</a:t>
                    </a:r>
                </a:p>`;
            });

            // 3. Robust XML Injection
            // This replaces everything inside the first <p:txBody> tags
            if (slideXml.includes('<p:txBody>')) {
                const head = slideXml.split('<p:txBody>')[0];
                const rest = slideXml.split('</p:txBody>')[1];
                
                // We reconstruct the txBody, keeping the required body properties (bodyPr) and list styles
                slideXml = head + 
                    '<p:txBody><a:bodyPr/><a:lstStyle/>' + 
                    generatedXml + 
                    '</p:txBody>' + 
                    rest;
                console.log("XML Injection successful");
            } else {
                throw new Error("The selected template doesn't have a valid text box on Slide 1.");
            }

            // 4. Placeholder Replacement
            slideXml = slideXml.replace(/\{\{TITLE\}\}/g, this.elements.songTitle.value || "");
            slideXml = slideXml.replace(/\{\{COPYRIGHT\}\}/g, this.elements.copyrightInfo.value || "");

            // 5. Save back to Zip
            zip.file("ppt/slides/slide1.xml", slideXml);

            // 6. Finalize and Download
            console.log("Generating Blob...");
            const content = await zip.generateAsync({ type: "blob" });
            const fileName = (this.elements.songTitle.value || "Song").replace(/[^a-z0-9]/gi, '_') + ".pptx";
            
            console.log("Triggering download...");
            saveAs(content, fileName);

        } catch (error) {
            console.error("Critical Error:", error);
            alert("Something went wrong: " + error.message);
        } finally {
            this.hideLoading();
        }
    },

    // --- TEMPLATE LOADING ---
    async loadTemplatesFromDirectory() {
        try {
            // Using ./ ensures it works on GitHub Pages subfolders
            const response = await fetch('./templates.json');
            if (!response.ok) throw new Error("Could not find templates.json");
            
            const filenames = await response.json();
            const galleryData = filenames.map(name => {
                const baseName = name.replace(/\.pptx$/i, '');
                return {
                    name: name,
                    thumbUrl: `./${encodeURIComponent(baseName)}.png`, 
                    getFile: async () => {
                        const res = await fetch(`./${encodeURIComponent(name)}`);
                        if (!res.ok) throw new Error(`Template file ${name} not found`);
                        const blob = await res.blob();
                        return new File([blob], name);
                    }
                };
            });
            this.renderTemplateGallery(galleryData);
        } catch (e) {
            console.warn("Gallery load failed:", e);
            document.getElementById('templateGallery').innerHTML = `<div class='p-4 text-red-500 text-xs'>Failed to load library. Check console (F12).</div>`;
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
                try {
                    this.showLoading('Downloading Template...');
                    this.selectedTemplateFile = await entry.getFile();
                    console.log("Selected file:", entry.name);
                    document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
                    card.classList.add('selected');
                    document.getElementById('selectedTemplateInfo').classList.remove('hidden');
                    document.getElementById('selectedTemplateName').textContent = entry.name;
                } catch (err) {
                    alert("Error downloading template: " + err.message);
                } finally {
                    this.hideLoading();
                }
            };
            grid.appendChild(card);
        });
        container.appendChild(grid);
    },

    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    setMode(m) {
        const isG = m === 'gen';
        document.getElementById('modeGen').classList.toggle('active', isG);
        document.getElementById('modeTrans').classList.toggle('active', !isG);
        document.getElementById('viewGen').classList.toggle('hidden', !isG);
        document.getElementById('viewTrans').classList.toggle('hidden', isG);
    }
};

App.init();
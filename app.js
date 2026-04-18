/**
 * LyricSlide Pro - v15.3 (Auto-Load Templates)
 * Feature: Left-aligned chords, Center-aligned lyrics
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
        
        this.loadTemplatesFromDirectory();
        
        window.LyricApp = this;
        console.log("LyricSlide Pro v15.3: Ready");
    },

    // --- DETECTION LOGIC ---

    isChordLine(line) {
        // Regex to detect lines that are primarily chords
        // Looks for musical notes (A-G) with common modifiers (maj, min, 7, #, b, /, etc)
        const chordRegex = /^[A-G][b#]?(2|4|5|6|7|9|11|13|maj|min|m|sus|dim|aug|add)?(\/[A-G][b#]?)?(\s+[A-G][b#]?(2|4|5|6|7|9|11|13|maj|min|m|sus|dim|aug|add)?(\/[A-G][b#]?)?)*\s*$/i;
        
        // Remove spaces and check if the line is just chords
        const cleanLine = line.trim();
        if (cleanLine === "") return false;
        
        // A chord line usually has a high ratio of spaces to characters or matches the regex
        return chordRegex.test(cleanLine);
    },

    // --- GENERATION LOGIC ---

    async generate() {
        if (!this.selectedTemplateFile) {
            alert("Please select a template first!");
            return;
        }

        const rawInput = this.elements.lyricsInput.value;
        if (!rawInput.trim()) {
            alert("Please enter some lyrics!");
            return;
        }

        this.showLoading("Creating Presentation...");

        try {
            const zip = await JSZip.loadAsync(this.selectedTemplateFile);
            
            // 1. Process Input into Slides
            // Split by brackets like [Chorus] or [Verse]
            const sections = rawInput.split(/(?=\[.*\])/).filter(s => s.trim() !== "");
            
            // 2. Load Slide 1 as a template base
            const slideTemplateXml = await zip.file("ppt/slides/slide1.xml").async("string");
            
            // 3. For each section, we will create a new slide file
            // Note: For a true robust PPTX generator, you'd update [Content_Types].xml and slide rels.
            // For this version, we will replace the content of Slide 1 with ALL slides 
            // (Standard hack: creating one long slide or modifying the template)
            
            let finalXmlLines = "";
            
            sections.forEach((section) => {
                const lines = section.split('\n');
                lines.forEach(line => {
                    if (line.startsWith('[')) return; // Skip headers for the text body
                    
                    const isChord = this.isChordLine(line);
                    const align = isChord ? 'l' : 'ctr';
                    const escapedLine = line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                    
                    // Generate paragraph XML with specific alignment
                    finalXmlLines += `
                        <a:p>
                            <a:pPr algn="${align}" />
                            <a:r>
                                <a:rPr lang="en-US" dirty="0" smtClean="0" />
                                <a:t>${escapedLine}</a:t>
                            </a:r>
                        </a:p>`;
                });
            });

            // 4. Inject into the XML
            // We search for a placeholder like {{LYRICS}} or just replace the first <a:p> block
            let newXml = slideTemplateXml;
            
            // Replace Title and Copyright placeholders if they exist in the XML
            newXml = newXml.replace(/\{\{TITLE\}\}/g, this.elements.songTitle.value);
            newXml = newXml.replace(/\{\{COPYRIGHT\}\}/g, this.elements.copyrightInfo.value);

            // Replace the text body content
            // This regex finds the first txBody and replaces its paragraph list
            newXml = newXml.replace(/<p:txBody>([\s\S]*?)<\/p:txBody>/, `<p:txBody><a:bodyPr/><a:lstStyle/>${finalXmlLines}</p:txBody>`);

            zip.file("ppt/slides/slide1.xml", newXml);

            // 5. Export
            const content = await zip.generateAsync({ type: "blob" });
            saveAs(content, `${this.elements.songTitle.value || 'Song'}.pptx`);

        } catch (err) {
            console.error(err);
            alert("Error generating PowerPoint: " + err.message);
        } finally {
            this.hideLoading();
        }
    },

    // --- TEMPLATE LOADING (Same as original) ---
    
    async loadTemplatesFromDirectory() {
        try {
            const response = await fetch('./templates.json');
            if (!response.ok) throw new Error("No manifest found");
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
            console.warn("Auto-load failed.", e);
            document.getElementById('templateGallery').innerHTML = `<div class="text-center py-8 text-slate-400 text-[10px]">templates.json not found.</div>`;
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

    changeSemitones(delta) {
        let current = parseInt(this.elements.semitoneDisplay.textContent);
        current += delta;
        this.elements.semitoneDisplay.textContent = current;
    },

    setMode(m) {
        const isG = m === 'gen';
        document.getElementById('modeGen').classList.toggle('active', isG);
        document.getElementById('modeTrans').classList.toggle('active', !isG);
        document.getElementById('viewGen').classList.toggle('hidden', !isG);
        document.getElementById('viewTrans').classList.toggle('hidden', isG);
    },
    showLoading(t) { this.elements.loadingText.textContent = t; this.elements.loadingOverlay.style.display = 'flex'; },
    hideLoading() { this.elements.loadingOverlay.style.display = 'none'; },
    clearTemplate() {
        this.selectedTemplateFile = null;
        document.querySelectorAll('.template-card').forEach(c => c.classList.remove('selected'));
        document.getElementById('selectedTemplateInfo').classList.add('hidden');
    }
};

App.init();
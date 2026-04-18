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
    
    async loadTemplatesFromDirectory() {
        try {
            // 1. Fetch the JSON manifest
            const response = await fetch('./templates.json');
            if (!response.ok) throw new Error("No manifest found");
            
            const filenames = await response.json();
            
            // 2. Map filenames to gallery objects
            const galleryData = filenames.map(name => {
                const baseName = name.replace(/\.pptx$/i, '');
                return {
                    name: name,
                    // Looks for a .png file with the same name as the .pptx
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
            console.warn("Auto-load failed. Waiting for manual folder selection.", e);
            document.getElementById('templateGallery').innerHTML = `
                <div class="text-center py-8 text-slate-400 text-[10px] leading-relaxed">
                    templates.json not found or server restricted.<br>
                    Run on a local server or click the folder icon.
                </div>`;
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
            
            // If the .png doesn't exist, show a placeholder icon
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

    // --- HELPER FUNCTIONS FOR UI ---
    
    changeSemitones(delta) {
        let current = parseInt(this.elements.semitoneDisplay.textContent);
        current += delta;
        this.elements.semitoneDisplay.textContent = current;
    },

    async loadForPreview(file) {
        // Basic feedback when a file is selected for transposition
        this.showLoading("Analyzing file...");
        setTimeout(() => this.hideLoading(), 500);
    },

    // (Rest of your existing functions: generate, transpose, shiftNote, etc. remain the same)
    // ... paste the rest of your original logic here ...
    
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
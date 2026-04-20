/* LyricSlide Pro - transpose.js (v12 Logic Extension) */

// Extend the existing App object
App.originalSlides = [];
App.elements.transFileInput = document.getElementById('transFileInput');
App.elements.transposeBtn = document.getElementById('transposeBtn');
App.elements.semitoneDisplay = document.getElementById('semitoneDisplay');

// Initialize Transpose Events
if (App.elements.transposeBtn) App.elements.transposeBtn.onclick = () => App.transpose();
if (App.elements.transFileInput) {
    App.elements.transFileInput.onchange = (e) => App.loadForPreview(e.target.files[0]);
}

// --- TRANSPOSITION ENGINE ---
App.transpose = async function() {
    const file = this.elements.transFileInput.files[0];
    const semitones = parseInt(this.elements.semitoneDisplay.textContent) || 0;

    if (!file) return alert('Select a PPTX file.');
    
    try {
        this.showLoading('Transposing...');
        const zip = await JSZip.loadAsync(file);
        
        // Slides
        const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
        for (const path of slideFiles) {
            let content = await zip.file(path).async('string');
            content = content.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => {
                const transposed = this.transposeLine(this.unescXml(text), semitones);
                return `<a:t>${this.escXml(transposed)}</a:t>`;
            });
            zip.file(path, content);
        }

        // Notes (v12 feature)
        const notesFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/notesSlides/notesSlide') && k.endsWith('.xml'));
        for (const path of notesFiles) {
            let notesContent = await zip.file(path).async('string');
            notesContent = notesContent.replace(/<a:t>(.*?)<\/a:t>/g, (_, text) => {
                return `<a:t>${this.escXml(this.transposeLine(this.unescXml(text), semitones))}</a:t>`;
            });
            zip.file(path, notesContent);
        }

        saveAs(await zip.generateAsync({ type: 'blob' }), file.name.replace('.pptx', `_transposed.pptx`));
        this.hideLoading();
    } catch (err) { this.hideLoading(); alert("Error: " + err.message); }
};

App.transposeLine = function(text, semitones) {
    if (semitones === 0) return text;
    const lines = text.split('\n');
    const chordRegex = /\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g;
    return lines.map(line => {
        const matches = [...line.matchAll(chordRegex)];
        if (matches.length === 0) return line;

        let result = line, offset = 0;
        for (const m of matches) {
            const original = m[0], pos = m.index + offset;
            const nc = this.shiftNote(m[1], semitones) + (m[2] || '') + (m[3] ? '/' + this.shiftNote(m[3].substring(1), semitones) : '');
            const diff = nc.length - original.length;
            result = result.substring(0, pos) + nc + result.substring(pos + original.length);
            if (diff > 0) {
                let spaceMatch = result.substring(pos + nc.length).match(/^ +/);
                if (spaceMatch && spaceMatch[0].length >= diff) result = result.substring(0, pos + nc.length) + result.substring(pos + nc.length + diff);
                else offset += diff;
            } else if (diff < 0) {
                result = result.substring(0, pos + nc.length) + " ".repeat(Math.abs(diff)) + result.substring(pos + nc.length);
            }
        }
        return result;
    }).join('\n');
};

App.shiftNote = function(note, semitones) {
    let list = note.includes('b') ? this.musical.flats : this.musical.keys;
    let idx = list.indexOf(note);
    if (idx === -1) { list = (list === this.musical.keys ? this.musical.flats : this.musical.keys); idx = list.indexOf(note); }
    if (idx === -1) return note;
    return (semitones >= 0 ? this.musical.keys : this.musical.flats)[(idx + semitones + 12) % 12];
};

// --- PREVIEW ENGINE ---
App.loadForPreview = async function(file) {
    if (!file) return;
    try {
        this.showLoading('Loading preview...');
        const zip = await JSZip.loadAsync(file);
        const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml')).sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
        this.originalSlides = [];
        for (const path of slideFiles) {
            const xml = await zip.file(path).async('string');
            const slideData = [];
            const spRegex = /<p:sp>([\s\S]*?)<\/p:sp>/g;
            let spMatch;
            while ((spMatch = spRegex.exec(xml)) !== null) {
                const spContent = spMatch[1];
                const pRegex = /<a:p>([\s\S]*?)<\/a:p>/g;
                let pMatch;
                while ((pMatch = pRegex.exec(spContent)) !== null) {
                    const pContent = pMatch[1], tagRegex = /<(a:t|a:br)[^>]*>(.*?)<\/\1>|<a:br\/>/g;
                    let pText = '', match;
                    while ((match = tagRegex.exec(pContent)) !== null) {
                        if (match[0].startsWith('<a:br')) pText += '\n';
                        else pText += this.unescXml(match[2] || '');
                    }
                    slideData.push({ text: pText, alignment: pContent.includes('algn="ctr"') ? 'center' : 'left' });
                }
            }
            this.originalSlides.push(slideData);
        }
        this.updatePreview(parseInt(this.elements.semitoneDisplay.textContent) || 0);
        this.hideLoading();
    } catch (err) { this.hideLoading(); }
};

App.updatePreview = function(semitones) {
    const container = document.getElementById('previewContainer');
    if (!container) return;
    container.innerHTML = '';
    this.originalSlides.forEach((slideData, idx) => {
        const card = document.createElement('div'), content = document.createElement('div');
        card.className = 'preview-card'; content.className = 'slide-content';
        slideData.forEach(p => {
            if (p.text.trim()) {
                const l = document.createElement('div');
                l.style.textAlign = p.alignment;
                const trans = this.transposeLine(p.text, semitones);
                l.innerHTML = this.escXml(trans).replace(/\b([A-G][b#]?)(m|maj|dim|aug|sus|2|4|6|7|9|add|11|13)*(\/[A-G][b#]?)?\b/g, '<span class="chord">$&</span>');
                content.appendChild(l);
            }
        });
        card.innerHTML = `<div class="text-[10px] text-slate-400 mb-2">Slide ${idx+1}</div>`;
        card.appendChild(content); container.appendChild(card);
    });
};

// Global UI Function
function changeSemitones(delta) {
    const d = document.getElementById('semitoneDisplay');
    let n = Math.max(-11, Math.min(11, (parseInt(d.textContent) || 0) + delta));
    d.textContent = (n > 0 ? '+' : '') + n;
    if (App.originalSlides.length > 0) App.updatePreview(n);
}
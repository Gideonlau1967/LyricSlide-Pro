/* LyricSlide Pro - Core Logic v15.3 (Fixed Background Preservation & Notes Linking) */

// ... (keep elements, musical, and theme sections exactly as they are) ...

    // --- GENERATION LOGIC (v15.3) ---
    async generate() {
        const file = this.selectedTemplateFile;
        const title = this.elements.songTitle.value || '';
        const lyrics = this.elements.lyricsInput.value || '';
        const copyright = this.elements.copyrightInfo.value || '';

        if (!file) return alert('Please select a template first.');
        if (!lyrics) return alert('Lyrics are required.');

        try {
            this.showLoading('Reading template...');
            const zip = await JSZip.loadAsync(file);
            const presXml = await zip.file('ppt/presentation.xml').async('string');
            const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
            const slideIds = this.getSlideIds(presXml);
            const slideRels = this.getSlideRels(presRelsXml);
            
            const templateRelPath = slideRels[slideIds[0].rid]; // e.g. "slides/slide1.xml"
            const templateSlidePath = `ppt/${templateRelPath}`;
            const templateXml = await zip.file(templateSlidePath).async('string');
            
            // --- FIX: Load the ORIGINAL relationship file to preserve backgrounds ---
            const templateRelFileName = templateRelPath.split('/').pop(); // slide1.xml
            const templateRelsPath = `ppt/slides/_rels/${templateRelFileName}.rels`;
            let originalRelsXml = await zip.file(templateRelsPath).async('string');

            const splitRegex = /\r?\n(?=\s*\[(?!(?:Title|Copyright Info|Lyrics and Chords)\])[^\]\n]+\])/;
            let sections = ("\n" + lyrics).split(splitRegex).filter(s => s.trim() !== '');
            if (sections.length === 0 && lyrics.trim() !== '') sections = [lyrics.trim()];
            
            const newZip = zip;
            const generated = [];

            for (let i = 0; i < sections.length; i++) {
                const sectionContent = sections[i].trim();
                let slideXml = templateXml;
                slideXml = this.lockInStyleAndReplace(slideXml, '[Title]', title);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Copyright Info]', copyright);
                slideXml = this.lockInStyleAndReplace(slideXml, '[Lyrics and Chords]', sectionContent);

                const name = `song_gen_${i + 1}.xml`;
                const path = `ppt/slides/${name}`;
                newZip.file(path, slideXml);

                // Create Presenter Notes
                const noteName = `notesSlideGen${i + 1}.xml`;
                newZip.file(`ppt/notesSlides/${noteName}`, this.createNotesSlideXml(sectionContent));

                // --- FIX: Append Notes relationship to the CLONED original rels ---
                // This ensures backgrounds (rId2, etc) and layouts (rId1) stay intact
                let relsDoc = new DOMParser().parseFromString(originalRelsXml, 'application/xml');
                
                // Add link to the new notes slide
                let noteRel = relsDoc.createElement('Relationship');
                noteRel.setAttribute('Id', 'rIdNotesCustom'); 
                noteRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide');
                noteRel.setAttribute('Target', `../notesSlides/${noteName}`);
                relsDoc.documentElement.appendChild(noteRel);
                
                newZip.file(`ppt/slides/_rels/${name}.rels`, new XMLSerializer().serializeToString(relsDoc));

                // Rel: NotesSlide -> Master (Standard)
                const noteRelXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/relationships">
                    <Relationship Id="rIdNotesMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMaster/notesMaster1.xml"/>
                </Relationships>`;
                newZip.file(`ppt/notesSlides/_rels/${noteName}.rels`, noteRelXml);

                generated.push({ id: 5000 + i, rid: `rIdGen${i + 1}`, name, path, noteName });
            }

            await this.syncPresentationRegistry(newZip, presXml, presRelsXml, generated);

            this.showLoading('Downloading...');
            const finalBlob = await newZip.generateAsync({ type: 'blob' });
            saveAs(finalBlob, `${(title || 'Song').replace(/[^a-z0-9]/gi, '_')}.pptx`);
            this.hideLoading();
        } catch (err) {
            console.error(err);
            alert("Error: " + err.message);
            this.hideLoading();
        }
    },

    // --- REGISTRY SYNC (CRITICAL FOR BACKGROUNDS & NOTES) ---
    async syncPresentationRegistry(newZip, presXml, presRelsXml, generated) {
        // Keep original Slide IDs from template or replace with generated? 
        // To be safe, we replace the slide list entirely with our generated ones.
        const sldIdLst = '<p:sldIdLst>' + generated.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('') + '</p:sldIdLst>';
        newZip.file('ppt/presentation.xml', presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst));

        let relsDoc = new DOMParser().parseFromString(presRelsXml, 'application/xml');
        let relationships = relsDoc.getElementsByTagName('Relationship');
        
        // Remove old slide relationships to prevent double-entries
        for (let j = relationships.length - 1; j >= 0; j--) {
            const type = relationships[j].getAttribute('Type');
            if (type && type.endsWith('slide')) {
                relationships[j].parentNode.removeChild(relationships[j]);
            }
        }
        
        generated.forEach(s => {
            let el = relsDoc.createElement('Relationship');
            el.setAttribute('Id', s.rid);
            el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
            el.setAttribute('Target', `slides/${s.name}`);
            relsDoc.documentElement.appendChild(el);
        });
        newZip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(relsDoc));

        // Update Content Types
        const ctFile = await newZip.file('[Content_Types].xml').async('string');
        let ctDoc = new DOMParser().parseFromString(ctFile, 'application/xml');
        let types = ctDoc.documentElement;
        
        // Clean up any existing generated references
        let overrides = types.getElementsByTagName('Override');
        for (let i = overrides.length - 1; i >= 0; i--) {
            const pn = overrides[i].getAttribute('PartName');
            if (pn.includes('/ppt/slides/song_gen') || pn.includes('/ppt/notesSlides/notesSlideGen')) {
                overrides[i].parentNode.removeChild(overrides[i]);
            }
        }

        generated.forEach(s => {
            let sld = ctDoc.createElement('Override');
            sld.setAttribute('PartName', `/${s.path}`);
            sld.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            types.appendChild(sld);

            let nte = ctDoc.createElement('Override');
            nte.setAttribute('PartName', `/ppt/notesSlides/${s.noteName}`);
            nte.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
            types.appendChild(nte);
        });
        newZip.file('[Content_Types].xml', new XMLSerializer().serializeToString(ctDoc));
    },

// ... (keep transposition, fonts, and logic helpers exactly as they are) ...

    createNotesSlideXml(text) {
        // Use a standard body placeholder index (usually 1 for notes)
        const lines = text.split(/\r?\n/).map(line => 
            `<a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>${this.escXml(line)}</a:t></a:r></a:p>`
        ).join('');
        
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                <p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                    <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${lines}</p:txBody>
                </p:sp>
            </p:spTree></p:cSld>
            <p:clrMapOver r:id="rIdNotesMaster1"/>
        </p:notes>`;
    },

// ... (keep the rest of the utility functions) ...

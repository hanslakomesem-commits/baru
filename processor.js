const JSZip = globalThis.JSZip;
if (!JSZip) throw new Error('JSZip gagal dimuat. Pastikan jszip.min.js ikut di-upload.');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const twipsPerCm = 1440 / 2.54;
const cmToTwips = (cm) => String(Math.round(cm * twipsPerCm));
const twipsToCm = (twips) => Number((Number(twips || '0') / twipsPerCm).toFixed(2));
const parser = new DOMParser();
const serializer = new XMLSerializer();
const FRONT_MATTER_RE = /^(DAFTAR\s+(ISI|TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT|KATA\s+PENGANTAR|MOTTO|PERSEMBAHAN|(?:LEMBAR|HALAMAN)\s+(PENGESAHAN|PERSETUJUAN|PERNYATAAN)|PERNYATAAN\s+(?:KEASLIAN(?:\s+TULISAN)?|ORISINALITAS)|SURAT\s+PERNYATAAN)$/i;
const FRONT_PAGE_START_RE = /^(DAFTAR\s+(ISI|TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT|KATA\s+PENGANTAR|MOTTO|PERSEMBAHAN|(?:LEMBAR|HALAMAN)\s+(PENGESAHAN|PERSETUJUAN|PERNYATAAN)|PERNYATAAN\s+(?:KEASLIAN(?:\s+TULISAN)?|ORISINALITAS)|SURAT\s+PERNYATAAN)$/i;
const END_MATTER_RE = /^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI|LAMPIRAN(?:\s*[-–—]?\s*LAMPIRAN)?|RIWAYAT\s+HIDUP)$/i;
const KNOWN_LEVEL2_RE = /^(KONTEKS|LATAR\s+BELAKANG|RUMUSAN\s+MASALAH|FOKUS\s+PENELITIAN|TUJUAN\s+PENELITIAN|KEGUNAAN\s+PENELITIAN|MANFAAT\s+PENELITIAN|DEFINISI\s+ISTILAH|KAJIAN\s+PENELITIAN\s+TERDAHULU|PENDEKATAN\s+DAN\s+JENIS\s+PENELITIAN|KEHADIRAN\s+PENELITI|LOKASI\s+PENELITIAN|SUMBER\s+DATA|PROSEDUR\s+PENGUMPULAN\s+DATA|TEKNIK\s+PENGUMPULAN\s+DATA|ANALISIS\s+DATA|PENGECEKAN\s+KEABSAHAN\s+DATA|PEGECEKAN\s+KEABSAHAN\s+DATA|SUMBER\s+DATA\s+(PRIMER|SEKUNDER)|OBSERVASI|WAWANCARA|DOKUMENTASI|ETIKA\s+(DEONTOLOGI|TELEOLOGI)|PERPANJANGAN\s+KEIKUTSERTAAN|KETEKUNAN\s+PENGAMATAN|TRIANGULASI|TAHAP\s+PRA\s+PENELITIAN|TAHAP\s+PROSES\s+PENELITIAN|PENYUSUNAN\s+LAPORAN|TAHAP(?:-TAHAP|\s+TAHAP)\s+PENELITIAN|PAPARAN\s+DATA(?:\s+DAN\s+TEMUAN\s+PENELITIAN)?|TEMUAN\s+PENELITIAN|PEMBAHASAN|KESIMPULAN|SARAN)$/i;
const LEVEL3_HINT_RE = /^(PENGERTIAN|DUA\s+TEORI|JENIS|MACAM|PRINSIP|FAKTOR|TUJUAN|CIRI|DASAR|TAHAPAN|PROSES|METODE|AKAD|PEMBIAYAAN)\b/i;
function parseXml(xml) {
    const doc = parser.parseFromString(xml, 'application/xml');
    const error = doc.querySelector('parsererror');
    if (error)
        throw new Error(`XML DOCX tidak valid: ${error.textContent || 'parsererror'}`);
    return doc;
}
function xmlString(doc) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${serializer.serializeToString(doc.documentElement)}`;
}
function wEl(doc, local) {
    return doc.createElementNS(W, `w:${local}`);
}
function wAttr(el, local, value) { el.setAttributeNS(W, `w:${local}`, value); }
function rAttr(el, local, value) { el.setAttributeNS(R, `r:${local}`, value); }
function directChild(parent, local) {
    return Array.from(parent.children).find((c) => c.namespaceURI === W && c.localName === local) || null;
}
function directChildren(parent, local) {
    return Array.from(parent.children).filter((c) => c.namespaceURI === W && c.localName === local);
}
function ensureDirectChild(parent, local, beforeLocal) {
    const existing = directChild(parent, local);
    if (existing)
        return existing;
    const el = wEl(parent.ownerDocument, local);
    if (beforeLocal) {
        const before = directChild(parent, beforeLocal);
        if (before)
            parent.insertBefore(el, before);
        else
            parent.appendChild(el);
    }
    else
        parent.appendChild(el);
    return el;
}
function ensureSectChild(sectPr, local, beforeNames) {
    const existing = directChild(sectPr, local);
    if (existing)
        return existing;
    const el = wEl(sectPr.ownerDocument, local);
    const before = Array.from(sectPr.children).find((c) => beforeNames.includes(c.localName));
    if (before)
        sectPr.insertBefore(el, before);
    else
        sectPr.appendChild(el);
    return el;
}
function normalizeText(s) { return s.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim(); }
function paragraphText(p) {
    return normalizeText(Array.from(p.getElementsByTagNameNS(W, 't')).map((n) => n.textContent || '').join(''));
}
function getBody(doc) {
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!body)
        throw new Error('word/document.xml tidak memiliki w:body.');
    return body;
}
function bodyParagraphs(doc) {
    return Array.from(getBody(doc).children).filter((e) => e.namespaceURI === W && e.localName === 'p');
}
function bodySectionProperties(doc) {
    const body = getBody(doc);
    const out = [];
    for (const child of Array.from(body.children)) {
        if (child.namespaceURI === W && child.localName === 'p') {
            const pPr = directChild(child, 'pPr');
            const sectPr = pPr && directChild(pPr, 'sectPr');
            if (sectPr)
                out.push(sectPr);
        }
    }
    const finalSect = directChild(body, 'sectPr');
    if (finalSect)
        out.push(finalSect);
    return out;
}
function chapterMatch(text) {
    const t = normalizeText(text).toUpperCase();
    return t.match(/^BAB\s*(V|IV|III|II|I)(?=\s|[A-ZÀ-ÖØ-Ý]|$)/);
}
function isChapterText(text) { return Boolean(chapterMatch(text)); }
function chapterRoman(text) { return chapterMatch(text)?.[1] || null; }
function isTableCaption(text) { return /^TABEL\s+(?:[IVXLC]+|\d+)(?:[.\-]\d+)*(?:\s|[.:])/i.test(normalizeText(text)); }
function isFigureCaption(text) { return /^GAMBAR\s+(?:[IVXLC]+|\d+)(?:[.\-]\d+)*(?:\s|[.:])/i.test(normalizeText(text)); }
function getOutlineLevel(p) {
    const pPr = directChild(p, 'pPr');
    const outline = pPr && directChild(pPr, 'outlineLvl');
    const val = Number(outline?.getAttributeNS(W, 'val'));
    if (Number.isFinite(val) && val >= 0 && val <= 2)
        return (val + 1);
    return undefined;
}
function hasExplicitBodyOutline(p) {
    const pPr = directChild(p, 'pPr');
    const outline = pPr && directChild(pPr, 'outlineLvl');
    const val = Number(outline?.getAttributeNS(W, 'val'));
    return Number.isFinite(val) && val >= 3;
}
function paragraphStyle(p) {
    const pPr = directChild(p, 'pPr');
    const style = pPr && directChild(pPr, 'pStyle');
    return style?.getAttributeNS(W, 'val') || '';
}
function styleHeadingLevel(p) {
    const s = paragraphStyle(p);
    const m = s.match(/^Heading([123])$/i) || s.match(/^Judul([123])$/i);
    if (!m)
        return undefined;
    return Number(m[1]);
}
function paragraphBoldRatio(p) {
    let total = 0;
    let bold = 0;
    for (const r of Array.from(p.getElementsByTagNameNS(W, 'r'))) {
        const text = Array.from(r.getElementsByTagNameNS(W, 't')).map((x) => x.textContent || '').join('');
        const n = text.length;
        total += n;
        const rPr = directChild(r, 'rPr');
        const b = rPr && directChild(rPr, 'b');
        if (b && b.getAttributeNS(W, 'val') !== '0' && b.getAttributeNS(W, 'val') !== 'false')
            bold += n;
    }
    return total ? bold / total : 0;
}
function paragraphIndentTwips(p) {
    const pPr = directChild(p, 'pPr');
    const ind = pPr && directChild(pPr, 'ind');
    return {
        left: Number(ind?.getAttributeNS(W, 'left') || ind?.getAttributeNS(W, 'start') || 0),
        first: Number(ind?.getAttributeNS(W, 'firstLine') || 0),
    };
}
function paragraphNumInfo(p, numbering) {
    const pPr = directChild(p, 'pPr');
    const numPr = pPr && directChild(pPr, 'numPr');
    if (!numPr)
        return null;
    const numId = directChild(numPr, 'numId')?.getAttributeNS(W, 'val') || '';
    const ilvl = directChild(numPr, 'ilvl')?.getAttributeNS(W, 'val') || '0';
    const lvl = numbering.get(numId)?.get(ilvl);
    return { numId, ilvl, fmt: lvl?.fmt || '', marker: lvl?.text || '' };
}
function hasComplexContent(p) {
    return p.getElementsByTagNameNS(W, 'drawing').length > 0 || p.getElementsByTagNameNS(W, 'object').length > 0 || p.getElementsByTagNameNS(W, 'pict').length > 0;
}
function isLikelyTocParagraph(p) {
    const style = paragraphStyle(p);
    if (/^TOC\d+$/i.test(style))
        return true;
    const text = paragraphText(p);
    const hasTab = p.getElementsByTagNameNS(W, 'tab').length > 0;
    return (hasTab && /\d+\s*$/.test(text)) || /\.{3,}\s*(?:\d+|[ivxlcdm]+)\s*$/i.test(text);
}
function likelySentence(text) {
    const t = normalizeText(text);
    return t.length > 115 || (/[,;:]\s/.test(t) && t.length > 70) || (/[.!?]$/.test(t) && t.split(/\s+/).length > 10);
}
async function loadNumberingMap(zip) {
    const out = new Map();
    const file = zip.file('word/numbering.xml');
    if (!file)
        return out;
    const doc = parseXml(await file.async('string'));
    const abstracts = new Map();
    for (const abs of Array.from(doc.getElementsByTagNameNS(W, 'abstractNum'))) {
        const id = abs.getAttributeNS(W, 'abstractNumId') || '';
        const levels = new Map();
        for (const lvl of directChildren(abs, 'lvl')) {
            const ilvl = lvl.getAttributeNS(W, 'ilvl') || '0';
            levels.set(ilvl, {
                fmt: directChild(lvl, 'numFmt')?.getAttributeNS(W, 'val') || '',
                text: directChild(lvl, 'lvlText')?.getAttributeNS(W, 'val') || '',
            });
        }
        abstracts.set(id, levels);
    }
    for (const num of Array.from(doc.getElementsByTagNameNS(W, 'num'))) {
        const numId = num.getAttributeNS(W, 'numId') || '';
        const absId = directChild(num, 'abstractNumId')?.getAttributeNS(W, 'val') || '';
        const levels = abstracts.get(absId);
        if (levels)
            out.set(numId, levels);
    }
    return out;
}
function detectHeadingCandidates(doc, numbering) {
    const ps = bodyParagraphs(doc);
    const out = [];
    let inMain = false;
    let currentChapter = '';
    let afterBibliography = false;
    // Daftar isi manual sering berisi teks seperti “BAB I ... 1” yang terlihat seperti heading asli.
    // Semua baris di antara judul DAFTAR ISI dan front-matter berikutnya sengaja diabaikan.
    const tocStart = ps.findIndex((p) => /^DAFTAR\s+ISI$/i.test(paragraphText(p)));
    let tocEnd = -1;
    if (tocStart >= 0) {
        for (let i = tocStart + 1; i < ps.length; i++) {
            if (/^(DAFTAR\s+(TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT|KATA\s+PENGANTAR)$/i.test(paragraphText(ps[i]))) {
                tocEnd = i;
                break;
            }
        }
    }
    for (let i = 0; i < ps.length; i++) {
        const p = ps[i];
        const text = paragraphText(p);
        if (!text || isLikelyTocParagraph(p) || hasComplexContent(p))
            continue;
        if (tocStart >= 0 && tocEnd > tocStart && i > tocStart && i < tocEnd)
            continue;
        if (isTableCaption(text) || isFigureCaption(text))
            continue;
        if (i > 0 && (isTableCaption(paragraphText(ps[i - 1])) || isFigureCaption(paragraphText(ps[i - 1]))) && text.length < 120)
            continue;
        const existing = getOutlineLevel(p) || styleHeadingLevel(p);
        if (!existing && hasExplicitBodyOutline(p))
            continue;
        const chapter = chapterRoman(text);
        const majorFront = FRONT_MATTER_RE.test(text);
        const majorEnd = END_MATTER_RE.test(text);
        if (chapter) {
            inMain = true;
            afterBibliography = false;
            currentChapter = chapter;
        }
        if (/^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI)$/i.test(text))
            afterBibliography = true;
        let level = 0;
        let confidence = 0;
        let reason = '';
        if (existing) {
            level = existing;
            confidence = 1;
            reason = `Sudah memiliki Outline/Heading level ${existing}`;
        }
        else if (chapter) {
            level = 1;
            confidence = 1;
            reason = `Terdeteksi BAB ${chapter}`;
        }
        else if (majorFront || majorEnd) {
            level = 1;
            confidence = 0.98;
            reason = 'Bagian utama/front matter skripsi';
        }
        else if (!inMain || afterBibliography) {
            continue;
        }
        else {
            // Bila BAB ditulis dua paragraf (“BAB I” lalu “PENDAHULUAN”), paragraf kedua tidak
            // dijadikan entri TOC terpisah. Pengguna tetap bisa meninjau BAB melalui Level 1.
            const previousText = i > 0 ? paragraphText(ps[i - 1]) : '';
            if (isChapterText(previousText) && text.length < 100 && text === text.toUpperCase())
                continue;
            const num = paragraphNumInfo(p, numbering);
            const ind = paragraphIndentTwips(p);
            const bold = paragraphBoldRatio(p);
            const short = text.length <= 160;
            const noSentenceEnd = !/[.!?]$/.test(text);
            if (/^[A-Z]\.[\s\t]+\S+/.test(text)) {
                level = 2;
                confidence = 0.97;
                reason = 'Pola Sub BAB huruf kapital (A., B., C.)';
            }
            else if (/^\d+(?:\.\d+){1,3}[\s\t]+\S+/.test(text)) {
                level = 3;
                confidence = 0.97;
                reason = 'Pola penomoran bertingkat (mis. 2.1 / 2.1.1)';
            }
            else if (/^\d+[.)][\s\t]+\S+/.test(text)) {
                level = 3;
                confidence = 0.9;
                reason = 'Pola Sub-sub BAB angka';
            }
            else if (KNOWN_LEVEL2_RE.test(text)) {
                level = 2;
                confidence = 0.96;
                reason = 'Nama Sub BAB akademik yang umum';
            }
            else if (num && num.fmt === 'upperLetter' && bold >= 0.7) {
                level = 2;
                confidence = 0.94;
                reason = 'Judul tebal dengan nomor list Word huruf kapital';
            }
            else if (num && ['lowerLetter', 'lowerRoman'].includes(num.fmt) && bold >= 0.7 && ind.left < 1000) {
                level = 2;
                confidence = 0.9;
                reason = 'Judul tebal dengan penomoran huruf/romawi pada level Sub BAB';
            }
            else if (num && ['lowerLetter', 'lowerRoman'].includes(num.fmt) && ind.left >= 1000 && bold >= 0.7 && short && !likelySentence(text)) {
                level = 3;
                confidence = 0.84;
                reason = 'Judul tebal bernomor huruf/romawi yang menjorok';
            }
            else if (num && num.fmt === 'decimal' && (Number(num.ilvl) >= 1 || ind.left >= 1200) && LEVEL3_HINT_RE.test(text) && short && !likelySentence(text)) {
                level = 3;
                confidence = 0.88;
                reason = 'Judul turunan bernomor/menjorok';
            }
            else if (LEVEL3_HINT_RE.test(text) && short && !likelySentence(text)) {
                level = 3;
                confidence = 0.78;
                reason = 'Judul turunan akademik terdeteksi';
            }
            else if (bold >= 0.8 && short && noSentenceEnd && ind.first === 0 && !/^(TABEL|GAMBAR)\b/i.test(text)) {
                level = 2;
                confidence = 0.82;
                reason = 'Paragraf pendek, tebal, dan bukan kalimat isi';
            }
        }
        if (level > 0)
            out.push({ paragraphIndex: i, text, level, confidence, reason, existingLevel: existing, chapter: currentChapter || undefined });
    }
    return out;
}
function isMainToc(doc) {
    const instr = [
        ...Array.from(doc.getElementsByTagNameNS(W, 'instrText')).map((n) => n.textContent || ''),
        ...Array.from(doc.getElementsByTagNameNS(W, 'fldSimple')).map((n) => n.getAttributeNS(W, 'instr') || ''),
    ];
    return instr.some((x) => /\bTOC\b/i.test(x) && !/\\f\s+[TG](?:\s|$)/i.test(x) && !/\\c\s+/i.test(x));
}
function hasSpecialToc(doc, mark) {
    const instr = [
        ...Array.from(doc.getElementsByTagNameNS(W, 'instrText')).map((n) => n.textContent || ''),
        ...Array.from(doc.getElementsByTagNameNS(W, 'fldSimple')).map((n) => n.getAttributeNS(W, 'instr') || ''),
    ];
    return instr.some((x) => new RegExp(`\\\\f\\s+${mark}(?:\\s|$)`, 'i').test(x));
}
function hasPageField(doc) {
    return Array.from(doc.getElementsByTagNameNS(W, 'instrText')).some((n) => /\bPAGE\b/i.test(n.textContent || '')) ||
        Array.from(doc.getElementsByTagNameNS(W, 'fldSimple')).some((n) => /\bPAGE\b/i.test(n.getAttributeNS(W, 'instr') || ''));
}
function findHeading(doc, regex) {
    return bodyParagraphs(doc).find((p) => regex.test(normalizeText(paragraphText(p)))) || null;
}
function findTocHeading(doc) { return findHeading(doc, /^DAFTAR\s+ISI$/i); }
function collectMargins(doc) {
    return bodySectionProperties(doc).map((sectPr) => {
        const pgMar = directChild(sectPr, 'pgMar');
        return {
            top: twipsToCm(pgMar?.getAttributeNS(W, 'top') || null),
            right: twipsToCm(pgMar?.getAttributeNS(W, 'right') || null),
            bottom: twipsToCm(pgMar?.getAttributeNS(W, 'bottom') || null),
            left: twipsToCm(pgMar?.getAttributeNS(W, 'left') || null),
        };
    });
}
function marginKey(m) { return `${m.top}|${m.right}|${m.bottom}|${m.left}`; }
function firstChapterParagraphIndex(doc) {
    return bodyParagraphs(doc).findIndex((p) => isChapterText(paragraphText(p)) && !isLikelyTocParagraph(p));
}
function paragraphHasHardPageBreak(p) {
    return Array.from(p.getElementsByTagNameNS(W, 'br')).some((b) => (b.getAttributeNS(W, 'type') || '').toLowerCase() === 'page');
}
function paragraphHasSection(p) {
    const pPr = directChild(p, 'pPr');
    return Boolean(pPr && directChild(pPr, 'sectPr'));
}
function safeEmptyParagraph(p) {
    if (paragraphText(p) || hasComplexContent(p))
        return false;
    if (paragraphHasSection(p))
        return false;
    if (p.getElementsByTagNameNS(W, 'br').length || p.getElementsByTagNameNS(W, 'fldChar').length || p.getElementsByTagNameNS(W, 'bookmarkStart').length || p.getElementsByTagNameNS(W, 'bookmarkEnd').length)
        return false;
    return true;
}
function countLikelyBlankParagraphs(doc) {
    const ps = bodyParagraphs(doc);
    const firstChapter = firstChapterParagraphIndex(doc);
    if (firstChapter < 0)
        return 0;
    let run = 0;
    let excess = 0;
    for (let i = firstChapter; i < ps.length; i++) {
        const text = paragraphText(ps[i]);
        if (/^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI)$/i.test(text))
            break;
        const empty = safeEmptyParagraph(ps[i]);
        if (empty)
            run++;
        else {
            if (run >= 3)
                excess += run - 1;
            run = 0;
        }
    }
    if (run >= 3)
        excess += run - 1;
    return excess;
}
function analyzeXml(fileName, doc, numbering, pageFieldOverride) {
    const paragraphs = bodyParagraphs(doc);
    const candidates = detectHeadingCandidates(doc, numbering);
    const chapters = Array.from(new Set(candidates.filter((x) => x.level === 1 && isChapterText(x.text)).map((x) => chapterRoman(x.text)).filter((x) => Boolean(x))));
    const margins = collectMargins(doc);
    const nativeToc = isMainToc(doc);
    const tocHeading = Boolean(findTocHeading(doc));
    const pageFields = pageFieldOverride ?? hasPageField(doc);
    const issues = [];
    const marginVariants = new Set(margins.map(marginKey));
    if (marginVariants.size > 1)
        issues.push({ id: 'margin-inconsistent', title: 'Margin tidak konsisten', detail: `Ditemukan ${marginVariants.size} konfigurasi margin pada ${margins.length} section.`, severity: 'warning', count: marginVariants.size });
    const expected = ['I', 'II', 'III', 'IV', 'V'];
    const missing = expected.filter((r) => !chapters.includes(r));
    if (missing.length)
        issues.push({ id: 'chapter-missing', title: 'BAB belum lengkap terdeteksi', detail: `Belum menemukan: ${missing.map((x) => `BAB ${x}`).join(', ')}.`, severity: 'warning', count: missing.length });
    const chaptersWithoutPageBreak = paragraphs.filter((p) => {
        if (isLikelyTocParagraph(p) || !isChapterText(paragraphText(p)))
            return false;
        const pPr = directChild(p, 'pPr');
        const hasBreak = Boolean(pPr && directChild(pPr, 'pageBreakBefore'));
        const hasBoundary = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === p);
        return !hasBreak && !hasBoundary;
    }).length;
    if (chaptersWithoutPageBreak)
        issues.push({ id: 'chapter-pagebreak', title: 'BAB belum dipastikan mulai halaman baru', detail: `${chaptersWithoutPageBreak} judul BAB belum mempunyai Page Break Before/Section Break tepat sebelum BAB.`, severity: 'warning', count: chaptersWithoutPageBreak });
    const firstChapterIndex = firstChapterParagraphIndex(doc);
    if (firstChapterIndex > 0) {
        const frontStarts = paragraphs.slice(0, firstChapterIndex).filter((p) => FRONT_PAGE_START_RE.test(paragraphText(p)));
        const missingFrontBreak = frontStarts.filter((p) => {
            const pPr = directChild(p, 'pPr');
            return !(pPr && directChild(pPr, 'pageBreakBefore')) && !paragraphHasHardPageBreak(p);
        }).length;
        if (missingFrontBreak)
            issues.push({ id: 'front-pagebreak', title: 'Bagian awal masih bergantung pada Enter', detail: `${missingFrontBreak} judul bagian awal (persetujuan/pengesahan/abstrak/dll.) belum mempunyai Page Break Before sehingga mudah bergeser.`, severity: 'warning', count: missingFrontBreak });
    }
    if (!nativeToc)
        issues.push({ id: 'toc-native', title: 'TOC Word belum ditemukan', detail: tocHeading ? 'Judul “DAFTAR ISI” ada, tetapi field TOC Word belum terdeteksi.' : 'Judul dan field TOC Word belum terdeteksi.', severity: 'warning' });
    if (!pageFields)
        issues.push({ id: 'page-field', title: 'Penomoran PAGE belum terdeteksi', detail: 'Nomor halaman mungkin manual atau belum dibuat sebagai field PAGE.', severity: 'info' });
    const headingGaps = candidates.filter((c) => !c.existingLevel && c.confidence >= 0.76 && c.level > 0).length;
    if (headingGaps)
        issues.push({ id: 'toc-level-gap', title: 'Sub BAB/Sub-sub BAB belum memiliki level TOC', detail: `${headingGaps} kandidat judul terdeteksi tetapi belum memiliki Outline Level. Ini penyebab umum TOC hanya berisi BAB.`, severity: 'warning', count: headingGaps });
    const blankRisk = countLikelyBlankParagraphs(doc);
    if (blankRisk)
        issues.push({ id: 'blank-risk', title: 'Enter/paragraf kosong berlebih', detail: `${blankRisk} paragraf kosong berlebih berpotensi membuat layout bergeser atau halaman kosong.`, severity: 'info', count: blankRisk });
    const frontMatter = candidates.filter((c) => c.level === 1 && FRONT_MATTER_RE.test(c.text)).map((c) => c.text);
    const structure = {
        frontMatter,
        chapters,
        bibliographyFound: paragraphs.some((p) => /^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI)$/i.test(paragraphText(p))),
        appendixFound: paragraphs.some((p) => /^LAMPIRAN/i.test(paragraphText(p))),
        tableCaptions: paragraphs.filter((p) => isTableCaption(paragraphText(p))).length,
        figureCaptions: paragraphs.filter((p) => isFigureCaption(paragraphText(p))).length,
        likelyBlankParagraphs: blankRisk,
    };
    const validationScore = Math.max(0, Math.min(100, 100 - issues.reduce((sum, x) => sum + (x.severity === 'error' ? 14 : x.severity === 'warning' ? 7 : 2), 0)));
    return {
        fileName,
        sections: bodySectionProperties(doc).length,
        paragraphs: paragraphs.length,
        chapters,
        subchapters: candidates.filter((x) => x.level === 2).length,
        subSubchapters: candidates.filter((x) => x.level === 3).length,
        hasNativeToc: nativeToc,
        hasTocHeading: tocHeading,
        hasPageFields: pageFields,
        margins,
        issues,
        headingCandidates: candidates,
        structure,
        validationScore,
    };
}
async function loadDocumentXml(zip) {
    const file = zip.file('word/document.xml');
    if (!file)
        throw new Error('Bukan DOCX yang valid: word/document.xml tidak ditemukan.');
    return parseXml(await file.async('string'));
}
function nextNonEmptyParagraphAfterSectPr(doc, sectPr) {
    const pPr = sectPr.parentElement;
    const p = pPr?.parentElement;
    if (!p || p.localName !== 'p')
        return null;
    let n = p.nextElementSibling;
    while (n) {
        if (n.namespaceURI === W && n.localName === 'p' && paragraphText(n))
            return n;
        n = n.nextElementSibling;
    }
    return null;
}
function firstNonEmptyBodyParagraph(doc) { return bodyParagraphs(doc).find((p) => Boolean(paragraphText(p))) || null; }
function sectionOwnerParagraph(sectPr) {
    const pPr = sectPr.parentElement;
    const p = pPr?.parentElement;
    return p && p.namespaceURI === W && p.localName === 'p' ? p : null;
}
function dedupeSectionBoundariesBeforeChapters(doc, changes) {
    const chapters = bodyParagraphs(doc).filter((p) => isChapterText(paragraphText(p)) && !isLikelyTocParagraph(p));
    let removed = 0;
    for (const chapter of chapters) {
        const boundaries = bodySectionProperties(doc).filter((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === chapter);
        if (boundaries.length <= 1)
            continue;
        // Pertahankan section break yang paling dekat dengan BAB. Hapus hanya duplikat pada paragraf kosong.
        for (const sect of boundaries.slice(0, -1)) {
            const owner = sectionOwnerParagraph(sect);
            if (!owner || paragraphText(owner))
                continue;
            sect.remove();
            removed++;
        }
    }
    if (removed)
        changes.push(`${removed} Section Break duplikat tepat sebelum BAB dibersihkan agar tidak membuat halaman/nomor lompat.`);
}
function removeTrailingEmptyParagraphsBefore(target) {
    let removed = 0;
    let prev = target.previousElementSibling;
    while (prev && prev.namespaceURI === W && prev.localName === 'p' && safeEmptyParagraph(prev)) {
        const next = prev.previousElementSibling;
        prev.remove();
        removed++;
        prev = next;
    }
    return removed;
}
function removeLeadingEmptyParagraphsAfter(target) {
    let removed = 0;
    let next = target.nextElementSibling;
    while (next && next.namespaceURI === W && next.localName === 'p' && safeEmptyParagraph(next)) {
        const after = next.nextElementSibling;
        next.remove();
        removed++;
        next = after;
    }
    return removed;
}
function flattenManualBreaksInParagraph(p) {
    const breaks = Array.from(p.getElementsByTagNameNS(W, 'br'));
    let changed = 0;
    for (const br of breaks) {
        const type = (br.getAttributeNS(W, 'type') || '').toLowerCase();
        if (type && type !== 'textwrapping' && type !== 'page')
            continue;
        const run = br.parentElement;
        if (run) {
            const texts = Array.from(run.getElementsByTagNameNS(W, 't'));
            const last = texts[texts.length - 1];
            if (last && last.textContent && !/\s$/.test(last.textContent))
                last.textContent += ' ';
        }
        br.remove();
        changed++;
    }
    return changed;
}
function ensurePageBreakBefore(p) {
    const pPr = ensureDirectChild(p, 'pPr');
    ensureDirectChild(pPr, 'pageBreakBefore');
}
function applyFrontMatterPagination(doc, changes) {
    const ps = bodyParagraphs(doc);
    const firstChapter = firstChapterParagraphIndex(doc);
    if (firstChapter <= 0)
        return;
    const firstNonEmpty = ps.find((p, i) => i < firstChapter && Boolean(paragraphText(p)));
    const title = firstNonEmpty ? paragraphText(firstNonEmpty) : '';
    let cover2 = null;
    if (title) {
        const firstIndex = ps.indexOf(firstNonEmpty);
        cover2 = ps.find((p, i) => i > firstIndex && i < firstChapter && paragraphText(p) === title) || null;
    }
    const starts = [];
    if (cover2)
        starts.push(cover2);
    for (let i = 0; i < firstChapter; i++) {
        const p = ps[i];
        if (FRONT_PAGE_START_RE.test(paragraphText(p)))
            starts.push(p);
    }
    let pages = 0;
    let removed = 0;
    let flattened = 0;
    const unique = Array.from(new Set(starts));
    for (const p of unique) {
        removed += removeTrailingEmptyParagraphsBefore(p);
        flattened += flattenManualBreaksInParagraph(p);
        ensurePageBreakBefore(p);
        pages++;
    }
    if (pages)
        changes.push(`${pages} awal halaman bagian depan (Cover 2/persetujuan/pengesahan/abstrak/dll.) dipastikan mulai halaman baru tanpa mengandalkan Enter kosong.`);
    if (removed)
        changes.push(`${removed} Enter kosong tepat sebelum awal halaman bagian depan dihapus; ruang vertikal di dalam cover tetap dilindungi.`);
    if (flattened)
        changes.push(`${flattened} manual page/line break pada judul bagian depan dinormalisasi ke Page Break Before.`);
}
function ensureChapterSectionBoundaries(doc, changes) {
    const ps = bodyParagraphs(doc);
    const chapters = ps.filter((p) => isChapterText(paragraphText(p)) && !isLikelyTocParagraph(p));
    const body = getBody(doc);
    const finalSect = directChild(body, 'sectPr');
    if (!finalSect)
        return;
    let added = 0;
    for (const chapter of chapters) {
        const boundaryExists = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === chapter);
        if (boundaryExists)
            continue;
        let prev = chapter.previousElementSibling;
        while (prev && !(prev.namespaceURI === W && prev.localName === 'p'))
            prev = prev.previousElementSibling;
        if (!prev)
            continue;
        const pPr = ensureDirectChild(prev, 'pPr');
        if (directChild(pPr, 'sectPr'))
            continue;
        const clone = finalSect.cloneNode(true);
        directChildren(clone, 'headerReference').forEach((x) => x.remove());
        directChildren(clone, 'footerReference').forEach((x) => x.remove());
        const type = ensureSectChild(clone, 'type', ['pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
        wAttr(type, 'val', 'nextPage');
        pPr.appendChild(clone);
        added++;
    }
    if (added)
        changes.push(`${added} Section Break (Next Page) ditambahkan tepat sebelum BAB agar pembagian BAB dan penomoran stabil.`);
}
function applyMarginsAndPaper(doc, options, changes) {
    if (!options.normalizeMargins)
        return;
    const sizes = options.paper === 'A4' ? { w: '11907', h: '16840', code: '9' } : { w: '11907', h: '18709', code: undefined };
    for (const sectPr of bodySectionProperties(doc)) {
        const pgSz = ensureSectChild(sectPr, 'pgSz', ['pgMar', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
        wAttr(pgSz, 'w', sizes.w);
        wAttr(pgSz, 'h', sizes.h);
        if (sizes.code)
            wAttr(pgSz, 'code', sizes.code);
        else
            pgSz.removeAttributeNS(W, 'code');
        const pgMar = ensureSectChild(sectPr, 'pgMar', ['pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
        wAttr(pgMar, 'top', cmToTwips(options.marginTopCm));
        wAttr(pgMar, 'right', cmToTwips(options.marginRightCm));
        wAttr(pgMar, 'bottom', cmToTwips(options.marginBottomCm));
        wAttr(pgMar, 'left', cmToTwips(options.marginLeftCm));
        if (!pgMar.hasAttributeNS(W, 'header'))
            wAttr(pgMar, 'header', '720');
        if (!pgMar.hasAttributeNS(W, 'footer'))
            wAttr(pgMar, 'footer', '720');
        if (!pgMar.hasAttributeNS(W, 'gutter'))
            wAttr(pgMar, 'gutter', '0');
    }
    changes.push(`Margin: kiri ${options.marginLeftCm} cm, atas ${options.marginTopCm} cm, kanan ${options.marginRightCm} cm, bawah ${options.marginBottomCm} cm; kertas ${options.paper}.`);
}
function setOutlineLevel(p, level) {
    const pPr = ensureDirectChild(p, 'pPr');
    const existing = directChild(pPr, 'outlineLvl');
    const outline = existing || ensureDirectChild(pPr, 'outlineLvl');
    if (level === 0) {
        wAttr(outline, 'val', '9');
        return;
    }
    wAttr(outline, 'val', String(level - 1));
}
function applyHeadingLevels(doc, numbering, options, changes) {
    if (!options.createToc && !options.repairHeadings)
        return;
    const ps = bodyParagraphs(doc);
    const candidates = detectHeadingCandidates(doc, numbering);
    const mainHeadingElements = [];
    let l1 = 0;
    let l2 = 0;
    let l3 = 0;
    let excluded = 0;
    for (const c of candidates) {
        const p = ps[c.paragraphIndex];
        if (!p)
            continue;
        let level = Object.prototype.hasOwnProperty.call(options.headingOverrides, String(c.paragraphIndex))
            ? options.headingOverrides[String(c.paragraphIndex)]
            : c.level;
        if (!options.includeFrontMatterInToc && FRONT_MATTER_RE.test(c.text))
            level = 0;
        setOutlineLevel(p, level);
        if (level === 0) {
            excluded++;
            continue;
        }
        if (options.repairHeadings) {
            const pPr = ensureDirectChild(p, 'pPr');
            ensureDirectChild(pPr, 'keepNext');
            ensureDirectChild(pPr, 'keepLines');
            // Hanya judul di bagian isi yang dibersihkan Enter-nya. Bagian depan tidak disentuh.
            if (c.chapter || isChapterText(c.text))
                mainHeadingElements.push(p);
            else {
                const idx = ps.indexOf(p);
                if (idx >= firstChapterParagraphIndex(doc))
                    mainHeadingElements.push(p);
            }
        }
        if (level === 1)
            l1++;
        else if (level === 2)
            l2++;
        else if (level === 3)
            l3++;
    }
    changes.push(`Level TOC diterapkan: ${l1} Level 1, ${l2} Level 2 (Sub BAB), ${l3} Level 3 (Sub-sub BAB)${excluded ? `, ${excluded} dikeluarkan` : ''}.`);
    if (options.repairHeadings) {
        let removed = 0;
        let flattened = 0;
        for (const p of mainHeadingElements) {
            flattened += flattenManualBreaksInParagraph(p);
            removed += removeTrailingEmptyParagraphsBefore(p);
            removed += removeLeadingEmptyParagraphsAfter(p);
        }
        changes.push('Judul BAB/Sub BAB/Sub-sub BAB diberi Keep with next/Keep lines agar judul tidak tertinggal sendirian di bawah halaman.');
        if (removed)
            changes.push(`${removed} Enter kosong yang menempel tepat sebelum/sesudah BAB, Sub BAB, dan Sub-sub BAB dihapus.`);
        if (flattened)
            changes.push(`${flattened} manual line break di dalam judul BAB/Sub BAB dirapikan agar judul tidak pecah karena Enter/Shift+Enter.`);
    }
}
function applyChapterPagination(doc, options, changes) {
    if (!options.repairChapterPagination)
        return;
    dedupeSectionBoundariesBeforeChapters(doc, changes);
    applyFrontMatterPagination(doc, changes);
    let changed = 0;
    let removed = 0;
    let flattened = 0;
    for (const p of bodyParagraphs(doc)) {
        if (!isChapterText(paragraphText(p)) || isLikelyTocParagraph(p))
            continue;
        removed += removeTrailingEmptyParagraphsBefore(p);
        removed += removeLeadingEmptyParagraphsAfter(p);
        flattened += flattenManualBreaksInParagraph(p);
        const isSectionStart = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === p);
        const pPr = ensureDirectChild(p, 'pPr');
        ensureDirectChild(pPr, 'keepNext');
        ensureDirectChild(pPr, 'keepLines');
        if (!isSectionStart)
            ensureDirectChild(pPr, 'pageBreakBefore');
        changed++;
    }
    if (changed)
        changes.push(`${changed} BAB dipastikan mulai halaman baru tanpa menambah Enter kosong.`);
    if (removed)
        changes.push(`${removed} Enter kosong yang menumpuk tepat di sekitar BAB dibersihkan.`);
    if (flattened)
        changes.push(`${flattened} manual break di judul BAB dinormalisasi.`);
}
function createFieldParagraph(doc, instruction, placeholder, style = 'TOC1') {
    const p = wEl(doc, 'p');
    const pPr = wEl(doc, 'pPr');
    if (style) {
        const pStyle = wEl(doc, 'pStyle');
        wAttr(pStyle, 'val', style);
        pPr.appendChild(pStyle);
    }
    p.appendChild(pPr);
    const beginRun = wEl(doc, 'r');
    const begin = wEl(doc, 'fldChar');
    wAttr(begin, 'fldCharType', 'begin');
    wAttr(begin, 'dirty', 'true');
    beginRun.appendChild(begin);
    p.appendChild(beginRun);
    const instrRun = wEl(doc, 'r');
    const instr = wEl(doc, 'instrText');
    instr.setAttribute('xml:space', 'preserve');
    instr.textContent = ` ${instruction} `;
    instrRun.appendChild(instr);
    p.appendChild(instrRun);
    const sepRun = wEl(doc, 'r');
    const sep = wEl(doc, 'fldChar');
    wAttr(sep, 'fldCharType', 'separate');
    sepRun.appendChild(sep);
    p.appendChild(sepRun);
    const hintRun = wEl(doc, 'r');
    const hint = wEl(doc, 't');
    hint.textContent = placeholder;
    hintRun.appendChild(hint);
    p.appendChild(hintRun);
    const endRun = wEl(doc, 'r');
    const end = wEl(doc, 'fldChar');
    wAttr(end, 'fldCharType', 'end');
    endRun.appendChild(end);
    p.appendChild(endRun);
    return p;
}
function createNativeTocBlock(doc, instruction, placeholder) {
    const sdt = wEl(doc, 'sdt');
    const pr = wEl(doc, 'sdtPr');
    const id = wEl(doc, 'id');
    wAttr(id, 'val', String(-Math.floor(100000000 + Math.random() * 1900000000)));
    pr.appendChild(id);
    const obj = wEl(doc, 'docPartObj');
    const gal = wEl(doc, 'docPartGallery');
    wAttr(gal, 'val', 'Table of Contents');
    obj.appendChild(gal);
    obj.appendChild(wEl(doc, 'docPartUnique'));
    pr.appendChild(obj);
    sdt.appendChild(pr);
    sdt.appendChild(wEl(doc, 'sdtEndPr'));
    const content = wEl(doc, 'sdtContent');
    content.appendChild(createFieldParagraph(doc, instruction, placeholder));
    sdt.appendChild(content);
    return sdt;
}
function createHeadingParagraph(doc, text) {
    const p = wEl(doc, 'p');
    const pPr = wEl(doc, 'pPr');
    const jc = wEl(doc, 'jc');
    wAttr(jc, 'val', 'center');
    pPr.appendChild(jc);
    const br = wEl(doc, 'pageBreakBefore');
    pPr.appendChild(br);
    const outline = wEl(doc, 'outlineLvl');
    wAttr(outline, 'val', '0');
    pPr.appendChild(outline);
    pPr.appendChild(wEl(doc, 'keepNext'));
    p.appendChild(pPr);
    const r = wEl(doc, 'r');
    const rPr = wEl(doc, 'rPr');
    rPr.appendChild(wEl(doc, 'b'));
    r.appendChild(rPr);
    const t = wEl(doc, 't');
    t.textContent = text;
    r.appendChild(t);
    p.appendChild(r);
    return p;
}
function removeLikelyManualTocRegion(doc, heading) {
    let removed = 0;
    let node = heading.nextElementSibling;
    while (node) {
        if (node.namespaceURI === W && node.localName === 'sdt')
            break;
        if (!(node.namespaceURI === W && node.localName === 'p'))
            break;
        const p = node;
        const text = paragraphText(p);
        if (/^(DAFTAR\s+(TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT|KATA\s+PENGANTAR)$/i.test(text))
            break;
        const likelyEntry = isLikelyTocParagraph(p) || /(?:\d+|[ivxlcdm]+)\s*$/i.test(text);
        const empty = !text;
        if (!likelyEntry && !empty)
            break;
        const next = node.nextElementSibling;
        node.remove();
        removed++;
        node = next;
    }
    return removed;
}
function findOrCreateHeadingBeforeFirstChapter(doc, regex, title, changes) {
    let heading = findHeading(doc, regex);
    if (heading)
        return heading;
    const firstChapter = bodyParagraphs(doc).find((p) => isChapterText(paragraphText(p)));
    if (!firstChapter)
        return null;
    heading = createHeadingParagraph(doc, title);
    firstChapter.parentNode.insertBefore(heading, firstChapter);
    changes.push(`Judul “${title}” dibuat karena belum ditemukan.`);
    return heading;
}
function refreshMainTocInstruction(doc) {
    for (const n of Array.from(doc.getElementsByTagNameNS(W, 'instrText'))) {
        const text = n.textContent || '';
        if (/\bTOC\b/i.test(text) && !/\\f\s+[TG](?:\s|$)/i.test(text) && !/\\c\s+/i.test(text)) {
            n.textContent = ' TOC \\o "1-3" \\h \\z \\u ';
            return true;
        }
    }
    return false;
}
function applyToc(doc, options, changes, warnings) {
    if (!options.createToc)
        return;
    if (isMainToc(doc)) {
        refreshMainTocInstruction(doc);
        changes.push('TOC Word yang sudah ada dipertahankan, dipastikan memakai Level 1–3 + Outline Level (\\u), lalu ditandai untuk diperbarui.');
        return;
    }
    const heading = findOrCreateHeadingBeforeFirstChapter(doc, /^DAFTAR\s+ISI$/i, 'DAFTAR ISI', changes);
    if (!heading) {
        warnings.push('BAB I dan posisi “DAFTAR ISI” tidak dapat ditentukan; TOC tidak disisipkan.');
        return;
    }
    if (options.removeLikelyManualToc) {
        const removed = removeLikelyManualTocRegion(doc, heading);
        if (removed)
            changes.push(`${removed} paragraf daftar isi manual yang terdeteksi dihapus sebelum TOC Word dibuat.`);
    }
    const toc = createNativeTocBlock(doc, 'TOC \\o "1-3" \\h \\z \\u', 'Klik kanan → Update Field / Ctrl+A → F9 untuk menghitung TOC.');
    heading.parentNode.insertBefore(toc, heading.nextSibling);
    changes.push('TOC Microsoft Word native dibuat dengan Level 1, Level 2, dan Level 3 — bukan hanya BAB.');
}
function hasTcField(p, mark) {
    return Array.from(p.getElementsByTagNameNS(W, 'fldSimple')).some((x) => new RegExp(`\\bTC\\b.*\\\\f\\s+${mark}`, 'i').test(x.getAttributeNS(W, 'instr') || ''));
}
function addTcField(p, text, mark) {
    if (hasTcField(p, mark))
        return;
    const fld = wEl(p.ownerDocument, 'fldSimple');
    const safeText = text.replace(/"/g, "'");
    wAttr(fld, 'instr', ` TC "${safeText}" \\f ${mark} \\l 1 `);
    p.appendChild(fld);
}
function removeLikelyManualSpecialListRegion(heading, kind) {
    let removed = 0;
    let node = heading.nextElementSibling;
    const entryRe = kind === 'Tabel' ? /^TABEL\s+(?:[IVXLC]+|\d+)/i : /^GAMBAR\s+(?:[IVXLC]+|\d+)/i;
    while (node) {
        if (node.namespaceURI === W && node.localName === 'sdt')
            break;
        if (!(node.namespaceURI === W && node.localName === 'p'))
            break;
        const text = paragraphText(node);
        if (text && (FRONT_PAGE_START_RE.test(text) || isChapterText(text)))
            break;
        const removable = !text || entryRe.test(text) || isLikelyTocParagraph(node);
        if (!removable)
            break;
        const next = node.nextElementSibling;
        node.remove();
        removed++;
        node = next;
    }
    return removed;
}
function applyTableFigureList(doc, options, changes, warnings) {
    const ps = bodyParagraphs(doc);
    if (options.createTableList) {
        const caps = ps.filter((p) => isTableCaption(paragraphText(p)));
        caps.forEach((p) => addTcField(p, paragraphText(p), 'T'));
        if (caps.length) {
            if (!hasSpecialToc(doc, 'T')) {
                const heading = findOrCreateHeadingBeforeFirstChapter(doc, /^DAFTAR\s+TABEL$/i, 'DAFTAR TABEL', changes);
                if (heading) {
                    if (options.removeLikelyManualToc) {
                        const removed = removeLikelyManualSpecialListRegion(heading, 'Tabel');
                        if (removed)
                            changes.push(`${removed} baris Daftar Tabel manual diganti dengan field dinamis Word agar tidak dobel.`);
                    }
                    heading.parentNode.insertBefore(createFieldParagraph(doc, 'TOC \\h \\z \\f T', 'Daftar tabel akan diperbarui oleh Microsoft Word.'), heading.nextSibling);
                }
            }
            changes.push(`${caps.length} caption tabel ditandai dan Daftar Tabel native/dinamis disiapkan.`);
        }
        else
            warnings.push('Tidak menemukan caption yang diawali “Tabel …”; Daftar Tabel tidak dibuat agar tidak ngasal.');
    }
    if (options.createFigureList) {
        const caps = ps.filter((p) => isFigureCaption(paragraphText(p)));
        caps.forEach((p) => addTcField(p, paragraphText(p), 'G'));
        if (caps.length) {
            if (!hasSpecialToc(doc, 'G')) {
                const heading = findOrCreateHeadingBeforeFirstChapter(doc, /^DAFTAR\s+GAMBAR$/i, 'DAFTAR GAMBAR', changes);
                if (heading) {
                    if (options.removeLikelyManualToc) {
                        const removed = removeLikelyManualSpecialListRegion(heading, 'Gambar');
                        if (removed)
                            changes.push(`${removed} baris Daftar Gambar manual diganti dengan field dinamis Word agar tidak dobel.`);
                    }
                    heading.parentNode.insertBefore(createFieldParagraph(doc, 'TOC \\h \\z \\f G', 'Daftar gambar akan diperbarui oleh Microsoft Word.'), heading.nextSibling);
                }
            }
            changes.push(`${caps.length} caption gambar ditandai dan Daftar Gambar native/dinamis disiapkan.`);
        }
        else
            warnings.push('Tidak menemukan caption yang diawali “Gambar …”; Daftar Gambar tidak dibuat agar tidak ngasal.');
    }
}
function applyBodyFormatting(doc, numbering, options, changes) {
    if (!options.normalizeBodyFormatting)
        return;
    const ps = bodyParagraphs(doc);
    const headingIdx = new Set(detectHeadingCandidates(doc, numbering).map((x) => x.paragraphIndex));
    const firstChapter = firstChapterParagraphIndex(doc);
    if (firstChapter < 0)
        return;
    let changed = 0;
    const minLen = options.mode === 'safe' ? 70 : 25;
    for (let i = firstChapter; i < ps.length; i++) {
        const p = ps[i];
        const text = paragraphText(p);
        if (/^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI)$/i.test(text))
            break;
        if (!text || headingIdx.has(i) || isTableCaption(text) || isFigureCaption(text) || END_MATTER_RE.test(text) || text.length < minLen || hasComplexContent(p))
            continue;
        const pPr = ensureDirectChild(p, 'pPr');
        if (directChild(pPr, 'numPr') || paragraphHasSection(p) || paragraphHasHardPageBreak(p) || p.getElementsByTagNameNS(W, 'tab').length)
            continue;
        const oldJc = directChild(pPr, 'jc')?.getAttributeNS(W, 'val') || '';
        const indExisting = directChild(pPr, 'ind');
        const left = Number(indExisting?.getAttributeNS(W, 'left') || indExisting?.getAttributeNS(W, 'start') || 0);
        const hanging = Number(indExisting?.getAttributeNS(W, 'hanging') || 0);
        if (options.mode === 'safe') {
            if (oldJc && !['left', 'both', 'distribute'].includes(oldJc))
                continue;
            if (left > 0 || hanging > 0 || paragraphBoldRatio(p) > 0.55)
                continue;
        }
        const jc = ensureDirectChild(pPr, 'jc');
        wAttr(jc, 'val', 'both');
        const spacing = ensureDirectChild(pPr, 'spacing');
        wAttr(spacing, 'line', String(Math.round(options.lineSpacing * 240)));
        wAttr(spacing, 'lineRule', 'auto');
        wAttr(spacing, 'after', '0');
        const ind = ensureDirectChild(pPr, 'ind');
        if (!ind.hasAttributeNS(W, 'hanging'))
            wAttr(ind, 'firstLine', cmToTwips(options.firstLineIndentCm));
        // Mode Aman tidak mengubah font per-run karena ini paling sering merusak cover/kutipan/format khusus.
        if (options.mode === 'total') {
            for (const r of Array.from(p.getElementsByTagNameNS(W, 'r'))) {
                if (!Array.from(r.getElementsByTagNameNS(W, 't')).length)
                    continue;
                const rPr = directChild(r, 'rPr') || r.insertBefore(wEl(doc, 'rPr'), r.firstChild);
                const fonts = ensureDirectChild(rPr, 'rFonts');
                wAttr(fonts, 'ascii', options.fontFamily);
                wAttr(fonts, 'hAnsi', options.fontFamily);
                wAttr(fonts, 'eastAsia', options.fontFamily);
                const sz = ensureDirectChild(rPr, 'sz');
                wAttr(sz, 'val', String(Math.round(options.fontSizePt * 2)));
                const szCs = ensureDirectChild(rPr, 'szCs');
                wAttr(szCs, 'val', String(Math.round(options.fontSizePt * 2)));
            }
        }
        changed++;
    }
    changes.push(`${changed} paragraf isi dinormalisasi${options.mode === 'safe' ? ' secara konservatif (cover/front matter, kutipan, list, heading, dan format khusus dilindungi)' : `: ${options.fontFamily} ${options.fontSizePt} pt, spasi ${options.lineSpacing}, justify, first-line ${options.firstLineIndentCm} cm`}.`);
}
function cleanBlankParagraphs(doc, options, changes) {
    if (!options.cleanBlankParagraphs)
        return;
    const ps = bodyParagraphs(doc);
    const firstChapter = firstChapterParagraphIndex(doc);
    if (firstChapter < 0)
        return;
    const threshold = options.mode === 'safe' ? 3 : 2;
    const remove = [];
    let run = [];
    const flush = () => { if (run.length >= threshold)
        remove.push(...run.slice(1)); run = []; };
    for (let i = firstChapter; i < ps.length; i++) {
        const p = ps[i];
        const text = paragraphText(p);
        if (/^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI)$/i.test(text)) {
            flush();
            break;
        }
        if (safeEmptyParagraph(p))
            run.push(p);
        else
            flush();
    }
    flush();
    remove.forEach((p) => p.remove());
    if (remove.length)
        changes.push(`${remove.length} paragraf kosong berlebih di bagian isi dihapus. Cover dan seluruh bagian sebelum BAB I sengaja tidak disentuh.`);
}
async function loadRels(zip) {
    const file = zip.file('word/_rels/document.xml.rels');
    if (!file)
        throw new Error('document.xml.rels tidak ditemukan.');
    return parseXml(await file.async('string'));
}
async function loadContentTypes(zip) {
    const file = zip.file('[Content_Types].xml');
    if (!file)
        throw new Error('[Content_Types].xml tidak ditemukan.');
    return parseXml(await file.async('string'));
}
function nextRid(rels) {
    const ids = Array.from(rels.documentElement.children).map((e) => e.getAttribute('Id') || '').map((x) => Number(x.replace(/^rId/i, ''))).filter(Number.isFinite);
    return `rId${Math.max(0, ...ids) + 1}`;
}
function nextPartNumber(zip, prefix) {
    const re = new RegExp(`^word/${prefix}(\\d+)\\.xml$`);
    const nums = Object.keys(zip.files).map((name) => name.match(re)?.[1]).filter(Boolean).map(Number);
    return Math.max(0, ...nums) + 1;
}
function addRel(rels, id, type, target) {
    const rel = rels.createElementNS(RELS, 'Relationship');
    rel.setAttribute('Id', id);
    rel.setAttribute('Type', type);
    rel.setAttribute('Target', target);
    rels.documentElement.appendChild(rel);
}
function addContentType(ct, partName, contentType) {
    if (Array.from(ct.documentElement.children).some((e) => e.getAttribute('PartName') === partName))
        return;
    const ov = ct.createElementNS(CT, 'Override');
    ov.setAttribute('PartName', partName);
    ov.setAttribute('ContentType', contentType);
    ct.documentElement.appendChild(ov);
}
function makeHeaderFooterXml(kind, align, withPage) {
    const rootTag = kind === 'header' ? 'hdr' : 'ftr';
    const pageField = withPage ? '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' : '<w:r><w:t></w:t></w:r>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${rootTag} xmlns:w="${W}" xmlns:r="${R}"><w:p><w:pPr><w:jc w:val="${align}"/></w:pPr>${pageField}</w:p></w:${rootTag}>`;
}
function removeReferences(sectPr, local) { directChildren(sectPr, local).forEach((e) => e.remove()); }
async function addPart(zip, rels, ct, kind, align, withPage) {
    const n = nextPartNumber(zip, kind);
    const fileName = `${kind}${n}.xml`;
    const path = `word/${fileName}`;
    zip.file(path, makeHeaderFooterXml(kind, align, withPage));
    const rid = nextRid(rels);
    addRel(rels, rid, `${R}/${kind}`, fileName);
    addContentType(ct, `/${path}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
    return rid;
}
function appendReference(sectPr, kind, type, rid) {
    const ref = wEl(sectPr.ownerDocument, `${kind}Reference`);
    wAttr(ref, 'type', type);
    rAttr(ref, 'id', rid);
    const firstNonRef = Array.from(sectPr.children).find((e) => !['headerReference', 'footerReference'].includes(e.localName));
    if (firstNonRef)
        sectPr.insertBefore(ref, firstNonRef);
    else
        sectPr.appendChild(ref);
}
async function configureFrontSection(zip, rels, ct, sect, start, hideFirstPage) {
    const pgNum = ensureSectChild(sect, 'pgNumType', ['cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    wAttr(pgNum, 'fmt', 'lowerRoman');
    if (start !== null)
        wAttr(pgNum, 'start', String(start));
    else
        pgNum.removeAttributeNS(W, 'start');
    if (hideFirstPage)
        ensureSectChild(sect, 'titlePg', ['textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    else
        directChild(sect, 'titlePg')?.remove();
    removeReferences(sect, 'headerReference');
    removeReferences(sect, 'footerReference');
    const defaultHeader = await addPart(zip, rels, ct, 'header', 'right', false);
    const defaultFooter = await addPart(zip, rels, ct, 'footer', 'center', true);
    appendReference(sect, 'header', 'default', defaultHeader);
    appendReference(sect, 'footer', 'default', defaultFooter);
    if (hideFirstPage) {
        const firstHeader = await addPart(zip, rels, ct, 'header', 'right', false);
        const firstFooter = await addPart(zip, rels, ct, 'footer', 'center', false);
        appendReference(sect, 'header', 'first', firstHeader);
        appendReference(sect, 'footer', 'first', firstFooter);
    }
}
async function configureMainSection(zip, rels, ct, sect, isChapterStart, restartAtOne) {
    const pgNum = ensureSectChild(sect, 'pgNumType', ['cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    wAttr(pgNum, 'fmt', 'decimal');
    if (restartAtOne)
        wAttr(pgNum, 'start', '1');
    else
        pgNum.removeAttributeNS(W, 'start');
    if (isChapterStart)
        ensureSectChild(sect, 'titlePg', ['textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    else
        directChild(sect, 'titlePg')?.remove();
    removeReferences(sect, 'headerReference');
    removeReferences(sect, 'footerReference');
    const defaultHeader = await addPart(zip, rels, ct, 'header', 'right', true);
    const defaultFooter = await addPart(zip, rels, ct, 'footer', 'center', false);
    appendReference(sect, 'header', 'default', defaultHeader);
    appendReference(sect, 'footer', 'default', defaultFooter);
    if (isChapterStart) {
        const firstHeader = await addPart(zip, rels, ct, 'header', 'right', false);
        const firstFooter = await addPart(zip, rels, ct, 'footer', 'center', true);
        appendReference(sect, 'header', 'first', firstHeader);
        appendReference(sect, 'footer', 'first', firstFooter);
    }
}
async function applyPageNumbering(zip, doc, options, changes, warnings) {
    if (!options.repairPageNumbering)
        return;
    dedupeSectionBoundariesBeforeChapters(doc, changes);
    ensureChapterSectionBoundaries(doc, changes);
    const sects = bodySectionProperties(doc);
    if (sects.length < 2) {
        warnings.push('Gagal membentuk section sebelum BAB I; penomoran otomatis dilewati.');
        return;
    }
    let frontEnd = -1;
    for (let i = 0; i < sects.length; i++) {
        const next = nextNonEmptyParagraphAfterSectPr(doc, sects[i]);
        if (next && chapterRoman(paragraphText(next)) === 'I') {
            frontEnd = i;
            break;
        }
    }
    if (frontEnd < 0) {
        warnings.push('Batas sebelum BAB I belum dapat ditentukan; penomoran otomatis dilewati agar file tidak rusak.');
        return;
    }
    const rels = await loadRels(zip);
    const ct = await loadContentTypes(zip);
    // Semua section sebelum BAB I diperlakukan sebagai bagian awal yang sama. Hanya halaman pertama (cover 1) yang nomor Romawinya disembunyikan.
    for (let i = 0; i <= frontEnd; i++) {
        await configureFrontSection(zip, rels, ct, sects[i], i === 0 ? 1 : null, i === 0);
    }
    for (let i = frontEnd + 1; i < sects.length; i++) {
        const startP = nextNonEmptyParagraphAfterSectPr(doc, sects[i - 1]);
        const startText = startP ? paragraphText(startP) : '';
        const chap = chapterRoman(startText);
        await configureMainSection(zip, rels, ct, sects[i], Boolean(chap), chap === 'I');
    }
    zip.file('word/_rels/document.xml.rels', xmlString(rels));
    zip.file('[Content_Types].xml', xmlString(ct));
    changes.push('Penomoran Word dibuat sebagai field PAGE: cover pertama tanpa nomor tampak, bagian awal Romawi berlanjut, BAB I mulai 1, halaman awal BAB bawah-tengah, halaman berikutnya kanan-atas.');
}
async function ensureUpdateFields(zip, changes) {
    const path = 'word/settings.xml';
    const file = zip.file(path);
    if (!file)
        return;
    const doc = parseXml(await file.async('string'));
    const root = doc.documentElement;
    let update = directChild(root, 'updateFields');
    if (!update) {
        update = wEl(doc, 'updateFields');
        root.appendChild(update);
    }
    wAttr(update, 'val', 'true');
    zip.file(path, xmlString(doc));
    if (!changes.some((x) => x.includes('update field')))
        changes.push('Microsoft Word disetel untuk update field otomatis (TOC/PAGE/Daftar Tabel/Daftar Gambar) saat dibuka.');
}
async function zipHasPageField(zip, doc) {
    if (hasPageField(doc))
        return true;
    const names = Object.keys(zip.files).filter((name) => /^word\/(header|footer)\d+\.xml$/i.test(name));
    for (const name of names) {
        const xml = await zip.file(name)?.async('string');
        if (xml && /(?:w:instr="[^"]*\bPAGE\b|<w:instrText[^>]*>[^<]*\bPAGE\b)/i.test(xml))
            return true;
    }
    return false;
}
export async function analyzeDocx(file) {
    const zip = await JSZip.loadAsync(file);
    const doc = await loadDocumentXml(zip);
    const numbering = await loadNumberingMap(zip);
    return analyzeXml(file.name, doc, numbering, await zipHasPageField(zip, doc));
}
export async function repairDocx(file, options) {
    const zip = await JSZip.loadAsync(file);
    const doc = await loadDocumentXml(zip);
    const numbering = await loadNumberingMap(zip);
    const changes = [];
    const warnings = [];
    // Rapikan section duplikat lalu buat boundary BAB sebelum modul lain bekerja.
    if (options.repairPageNumbering || options.repairChapterPagination)
        dedupeSectionBoundariesBeforeChapters(doc, changes);
    if (options.repairPageNumbering)
        ensureChapterSectionBoundaries(doc, changes);
    applyMarginsAndPaper(doc, options, changes);
    applyHeadingLevels(doc, numbering, options, changes);
    applyChapterPagination(doc, options, changes);
    applyToc(doc, options, changes, warnings);
    applyTableFigureList(doc, options, changes, warnings);
    applyBodyFormatting(doc, numbering, options, changes);
    cleanBlankParagraphs(doc, options, changes);
    await applyPageNumbering(zip, doc, options, changes, warnings);
    zip.file('word/document.xml', xmlString(doc));
    if (options.createToc || options.repairPageNumbering || options.createTableList || options.createFigureList)
        await ensureUpdateFields(zip, changes);
    const output = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    const outputName = file.name.replace(/\.docx$/i, '') + '_RAPI_ZAINNET_V3.docx';
    const reloaded = await JSZip.loadAsync(output);
    const afterDoc = await loadDocumentXml(reloaded);
    const afterNumbering = await loadNumberingMap(reloaded);
    const analysisAfter = analyzeXml(outputName, afterDoc, afterNumbering, await zipHasPageField(reloaded, afterDoc));
    if (options.validateDocument)
        changes.push(`Validasi akhir selesai. Skor struktur: ${analysisAfter.validationScore}/100 dengan ${analysisAfter.issues.length} catatan tersisa.`);
    return { blob: output, outputName, changes, warnings, analysisAfter };
}

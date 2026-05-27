(() => {
    const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

    const MARKER_RE = /\\\*|\*{1,3}/;

    function getTagName(element) {
        return element.tagName || element.nodeName || '';
    }

    function getDirectChild(element, tagName) {
        if (!element) return null;
        for (const child of Array.from(element.children || [])) {
            if (getTagName(child) === tagName) return child;
        }
        return null;
    }

    function getTextFromRun(run) {
        const tNodes = run.getElementsByTagName('a:t');
        let text = '';
        for (let i = 0; i < tNodes.length; i++) {
            text += tNodes[i].textContent || '';
        }
        return text;
    }

    function getElementAttrBool(element, attrName) {
        if (!element || !element.hasAttribute(attrName)) return false;
        const value = String(element.getAttribute(attrName) || '').toLowerCase();
        return value === '1' || value === 'true';
    }

    function escapePlainText(text) {
        return String(text || '')
            .replace(/\\/g, '\\\\')
            .replace(/\*/g, '\\*');
    }

    function wrapRunText(text, bold, italic) {
        const escaped = escapePlainText(text);
        if (!escaped) return '';
        if (bold && italic) return `***${escaped}***`;
        if (bold) return `**${escaped}**`;
        if (italic) return `*${escaped}*`;
        return escaped;
    }

    function hasRichMarkers(text) {
        const value = String(text || '');
        if (!value.includes('*')) return false;
        for (let i = 0; i < value.length; i++) {
            if (value[i] === '\\') {
                i += 1;
                continue;
            }
            if (value[i] === '*') return true;
        }
        return false;
    }

    function stripDelimitersForFallback(text) {
        let out = '';
        const value = String(text || '');
        for (let i = 0; i < value.length; i++) {
            if (value[i] === '\\' && i + 1 < value.length) {
                out += value[i + 1];
                i += 1;
                continue;
            }
            if (value.slice(i, i + 3) === '***') {
                i += 2;
                continue;
            }
            if (value.slice(i, i + 2) === '**') {
                i += 1;
                continue;
            }
            if (value[i] === '*') continue;
            out += value[i];
        }
        return out;
    }

    function encodeParagraph(paragraph) {
        const runs = Array.from(paragraph.getElementsByTagName('a:r'));
        let hadMultipleRuns = runs.length > 1;
        let hadStyledRuns = false;
        const parts = [];

        for (const run of runs) {
            const text = getTextFromRun(run);
            if (!text) continue;

            const rPr = getDirectChild(run, 'a:rPr');
            const bold = getElementAttrBool(rPr, 'b');
            const italic = getElementAttrBool(rPr, 'i');
            if (bold || italic) hadStyledRuns = true;

            parts.push(wrapRunText(text, bold, italic));
        }

        const encoded = parts.join('');
        if (!hadMultipleRuns && !hadStyledRuns) {
            return {
                text: escapePlainText(parts.join('') || getTextFromParagraphPlain(paragraph)),
                hadMultipleRuns: false,
                hadStyledRuns: false
            };
        }

        return {
            text: encoded,
            hadMultipleRuns,
            hadStyledRuns
        };
    }

    function getTextFromParagraphPlain(paragraph) {
        const runs = paragraph.getElementsByTagName('a:r');
        let text = '';
        for (let i = 0; i < runs.length; i++) {
            text += getTextFromRun(runs[i]);
        }
        return text;
    }

    function parseMarkdown(input) {
        const text = String(input || '');
        const warnings = [];
        const segments = [];
        let buffer = '';
        let bold = false;
        let italic = false;

        function flush() {
            if (!buffer) return;
            segments.push({ text: buffer, bold, italic });
            buffer = '';
        }

        let i = 0;
        while (i < text.length) {
            if (text[i] === '\\' && i + 1 < text.length) {
                buffer += text[i + 1];
                i += 2;
                continue;
            }
            if (text.slice(i, i + 3) === '***') {
                flush();
                bold = !bold;
                italic = !italic;
                i += 3;
                continue;
            }
            if (text.slice(i, i + 2) === '**') {
                flush();
                bold = !bold;
                i += 2;
                continue;
            }
            if (text[i] === '*') {
                flush();
                italic = !italic;
                i += 1;
                continue;
            }
            buffer += text[i];
            i += 1;
        }
        flush();

        if (bold || italic) {
            warnings.push({
                type: 'invalid-rich-text',
                message: 'Unclosed markdown emphasis markers; text was inserted without inline formatting.'
            });
            return {
                segments: [{ text: stripDelimitersForFallback(text), bold: false, italic: false }],
                warnings,
                usedFallback: true
            };
        }

        if (segments.length === 0) {
            segments.push({ text: '', bold: false, italic: false });
        }

        return { segments, warnings, usedFallback: false };
    }

    function setRunBooleanAttr(rPr, attrName, value) {
        if (!rPr) return;
        if (value) {
            rPr.setAttribute(attrName, '1');
        } else if (rPr.hasAttribute(attrName)) {
            rPr.setAttribute(attrName, '0');
        }
    }

    function cloneRunFromTemplate(templateRun, doc) {
        const run = doc.createElementNS(A_NS, 'a:r');
        const templateRPr = getDirectChild(templateRun, 'a:rPr');
        if (templateRPr) {
            run.appendChild(templateRPr.cloneNode(true));
        }
        const textNode = doc.createElementNS(A_NS, 'a:t');
        textNode.setAttribute('xml:space', 'preserve');
        run.appendChild(textNode);
        return run;
    }

    function findRunInsertPoint(paragraph) {
        const endParaRPr = getDirectChild(paragraph, 'a:endParaRPr');
        return endParaRPr || null;
    }

    function removeRunElements(paragraph) {
        for (const child of Array.from(paragraph.children || [])) {
            if (getTagName(child) === 'a:r') {
                paragraph.removeChild(child);
            }
        }
    }

    function applyToParagraph(paragraph, segments, options = {}) {
        const templateRun = options.templateRun;
        const styleRun = options.styleRun;
        if (!templateRun || !segments.length) return false;

        const doc = paragraph.ownerDocument;
        const insertBefore = findRunInsertPoint(paragraph);
        removeRunElements(paragraph);

        let firstRun = null;
        for (const segment of segments) {
            if (!segment.text && segments.length > 1) continue;

            const run = cloneRunFromTemplate(templateRun, doc);
            const textNode = getDirectChild(run, 'a:t');
            textNode.textContent = segment.text;
            textNode.setAttribute('xml:space', 'preserve');

            const rPr = getDirectChild(run, 'a:rPr');
            setRunBooleanAttr(rPr, 'b', Boolean(segment.bold));
            setRunBooleanAttr(rPr, 'i', Boolean(segment.italic));

            if (insertBefore) {
                paragraph.insertBefore(run, insertBefore);
            } else {
                paragraph.appendChild(run);
            }

            if (!firstRun) firstRun = run;
            if (typeof styleRun === 'function') {
                styleRun(paragraph, run, segment);
            }
        }

        if (!firstRun) {
            const run = cloneRunFromTemplate(templateRun, doc);
            const textNode = getDirectChild(run, 'a:t');
            textNode.textContent = '';
            textNode.setAttribute('xml:space', 'preserve');
            if (insertBefore) paragraph.insertBefore(run, insertBefore);
            else paragraph.appendChild(run);
            firstRun = run;
            if (typeof styleRun === 'function') styleRun(paragraph, run, { text: '', bold: false, italic: false });
        }

        return true;
    }

    function unescapePlainText(text) {
        let out = '';
        const value = String(text || '');
        for (let i = 0; i < value.length; i++) {
            if (value[i] === '\\' && i + 1 < value.length) {
                out += value[i + 1];
                i += 1;
                continue;
            }
            out += value[i];
        }
        return out;
    }

    function shouldUseRichPath(text, parsed) {
        if (parsed.usedFallback) return false;
        if (hasRichMarkers(text)) return true;
        return parsed.segments.some(segment => segment.bold || segment.italic);
    }

    function plainTextForComparison(text) {
        return stripDelimitersForFallback(text).trim();
    }

    window.PptxRichText = {
        MARKER_RE,
        hasRichMarkers,
        encodeParagraph,
        parseMarkdown,
        applyToParagraph,
        shouldUseRichPath,
        stripDelimitersForFallback,
        unescapePlainText,
        plainTextForComparison
    };
})();

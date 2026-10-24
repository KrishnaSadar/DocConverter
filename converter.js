/**
 * DOCX Question Bank Converter
 * Client-side parser and converter for converting input_format.docx
 * into the strict output_format.docx structure.
 */

(function (window) {
    'use strict';

    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

    const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
    const OLE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject';

    const MIME_MAP = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'emf': 'image/x-emf',
        'wmf': 'image/x-wmf',
        'svg': 'image/svg+xml',
        'tiff': 'image/tiff',
        'tif': 'image/tiff',
        'bin': 'application/vnd.openxmlformats-officedocument.oleObject'
    };

    const LETTER_TO_NUM = {
        'a': '1', 'A': '1',
        'b': '2', 'B': '2',
        'c': '3', 'C': '3',
        'd': '4', 'D': '4'
    };

    /**
     * Get concatenated text from all <w:t> elements within a node
     */
    function getNodeText(node) {
        if (!node) return '';
        const tElements = node.getElementsByTagNameNS ?
            node.getElementsByTagNameNS(W_NS, 't') :
            node.getElementsByTagName('w:t');
        if (tElements && tElements.length > 0) {
            let res = '';
            for (let i = 0; i < tElements.length; i++) {
                res += tElements[i].textContent || '';
            }
            return res;
        }
        return node.textContent || '';
    }

    /**
     * Helper to get direct children or descendants by local name
     */
    function getElementsByLocalName(parent, localName) {
        const result = [];
        if (!parent) return result;
        const all = parent.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName === localName || el.nodeName.endsWith(':' + localName)) {
                result.push(el);
            }
        }
        return result;
    }

    /**
     * Helper to get direct child elements by local name
     */
    function getDirectChildElements(parent, localName) {
        const result = [];
        if (!parent || !parent.childNodes) return result;
        for (let i = 0; i < parent.childNodes.length; i++) {
            const child = parent.childNodes[i];
            if (child.nodeType === 1) { // ELEMENT_NODE
                if (child.localName === localName || child.nodeName.endsWith(':' + localName)) {
                    result.push(child);
                }
            }
        }
        return result;
    }

    /**
     * Convert Base64 string to Uint8Array
     */
    function base64ToUint8Array(base64) {
        const binaryString = window.atob(base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes;
    }

    /**
     * Main DocxConverter Class
     */
    class DocxConverter {
        constructor(options = {}) {
            this.positiveMarks = options.positiveMarks || '4';
            this.negativeMarks = options.negativeMarks || '1';
            this.onProgress = options.onProgress || function () {};
        }

        setProgress(percent, message) {
            this.onProgress({ percent, message });
        }

        /**
         * Main conversion pipeline
         * @param {File|Blob|ArrayBuffer} inputFile - Input .docx file
         * @param {File|Blob|ArrayBuffer|string} [templateFile] - Optional template .docx (or base64 string)
         * @returns {Promise<{blob: Blob, stats: Object, preview: Array}>}
         */
        async convert(inputFile, templateFile) {
            this.setProgress(5, 'Loading input document archive...');

            if (typeof JSZip === 'undefined') {
                throw new Error('JSZip library is not loaded. Please ensure jszip.min.js is included.');
            }

            // 1. Load Input ZIP
            const inputZip = await JSZip.loadAsync(inputFile);

            this.setProgress(15, 'Reading input document XML and relationships...');
            const docXmlStr = await inputZip.file('word/document.xml').async('text');
            const parser = new DOMParser();
            const docXml = parser.parseFromString(docXmlStr, 'application/xml');

            // Read input relationships
            let relsXml = null;
            const relsFile = inputZip.file('word/_rels/document.xml.rels');
            if (relsFile) {
                const relsXmlStr = await relsFile.async('text');
                relsXml = parser.parseFromString(relsXmlStr, 'application/xml');
            }

            // Map all input relationships of type image & OLE
            const inputMediaRels = [];
            if (relsXml) {
                const relElements = getElementsByLocalName(relsXml, 'Relationship');
                for (let i = 0; i < relElements.length; i++) {
                    const r = relElements[i];
                    const rtype = r.getAttribute('Type') || '';
                    const rid = r.getAttribute('Id') || '';
                    const target = r.getAttribute('Target') || '';

                    if (rtype === IMAGE_REL || rtype === OLE_REL) {
                        const cleanTarget = target.replace(/^\.\.\//, '').replace(/^\//, '');
                        const zipPath = cleanTarget.startsWith('word/') ? cleanTarget : 'word/' + cleanTarget;
                        inputMediaRels.push({
                            id: rid,
                            type: rtype,
                            target: cleanTarget,
                            zipPath: zipPath,
                            ext: cleanTarget.split('.').pop().toLowerCase()
                        });
                    }
                }
            }

            this.setProgress(25, 'Parsing tables, questions, answer key, and solutions...');

            // Find all tables in document.xml
            const tables = getElementsByLocalName(docXml, 'tbl');
            if (tables.length === 0) {
                throw new Error('No tables found in input document. Please verify the input file format.');
            }

            // Parse Question Table (typically table 0, with rows containing question numbers)
            const qData = this.parseQuestions(tables[0], parser);

            // Parse Answer Key
            const answerMap = this.parseAnswerKey(tables, docXml);

            // Parse Solutions
            const solutionMap = this.parseSolutions(tables);

            this.setProgress(45, 'Loading destination template...');

            // 2. Load Template ZIP
            let templateZip;
            if (templateFile) {
                if (typeof templateFile === 'string') {
                    // Base64 string
                    const templateBytes = base64ToUint8Array(templateFile);
                    templateZip = await JSZip.loadAsync(templateBytes);
                } else {
                    templateZip = await JSZip.loadAsync(templateFile);
                }
            } else if (typeof DEFAULT_TEMPLATE_B64 !== 'undefined' && DEFAULT_TEMPLATE_B64) {
                const templateBytes = base64ToUint8Array(DEFAULT_TEMPLATE_B64);
                templateZip = await JSZip.loadAsync(templateBytes);
            } else {
                throw new Error('No template document provided and DEFAULT_TEMPLATE_B64 is unavailable.');
            }

            this.setProgress(55, 'Preparing template and transferring embedded media...');

            // Read template document.xml, rels, and [Content_Types].xml
            const tDocXmlStr = await templateZip.file('word/document.xml').async('text');
            const tDocXml = parser.parseFromString(tDocXmlStr, 'application/xml');

            const tRelsXmlStr = await templateZip.file('word/_rels/document.xml.rels').async('text');
            const tRelsXml = parser.parseFromString(tRelsXmlStr, 'application/xml');

            const ctXmlStr = await templateZip.file('[Content_Types].xml').async('text');
            const ctXml = parser.parseFromString(ctXmlStr, 'application/xml');

            // Find maximum existing relationship ID in template
            const tRelElements = getElementsByLocalName(tRelsXml, 'Relationship');
            let maxRidNum = 0;
            for (let i = 0; i < tRelElements.length; i++) {
                const id = tRelElements[i].getAttribute('Id') || '';
                const m = id.match(/rId(\d+)/i);
                if (m) {
                    const n = parseInt(m[1], 10);
                    if (n > maxRidNum) maxRidNum = n;
                }
            }

            // Transfer media files and build oldRid -> newRid map
            const oldRidToNewRid = {};
            const typesNode = ctXml.documentElement;
            const relsRoot = tRelsXml.documentElement;

            for (const mediaRel of inputMediaRels) {
                const file = inputZip.file(mediaRel.zipPath);
                if (!file) {
                    console.warn(`Referenced file not found in input zip: ${mediaRel.zipPath}`);
                    continue;
                }

                const fileBytes = await file.async('uint8array');
                maxRidNum++;
                const newRid = 'rId' + maxRidNum;
                oldRidToNewRid[mediaRel.id] = newRid;

                let targetInRels = '';
                let destZipPath = '';

                if (mediaRel.type === IMAGE_REL) {
                    const ext = mediaRel.ext || 'png';
                    const fileName = `input_${mediaRel.id}.${ext}`;
                    targetInRels = `media/${fileName}`;
                    destZipPath = `word/media/${fileName}`;

                    // Add image to zip
                    templateZip.file(destZipPath, fileBytes);

                    // Add content type if needed
                    const mime = MIME_MAP[ext] || 'image/png';
                    this.ensureContentTypeDefault(ctXml, ext, mime);
                } else {
                    // OLE Object
                    const fileName = `input_${mediaRel.id}.bin`;
                    targetInRels = `embeddings/${fileName}`;
                    destZipPath = `word/embeddings/${fileName}`;

                    // Add ole to zip
                    templateZip.file(destZipPath, fileBytes);

                    // Add content type override
                    this.ensureContentTypeOverride(
                        ctXml,
                        `/word/embeddings/${fileName}`,
                        'application/vnd.openxmlformats-officedocument.oleObject'
                    );
                }

                // Add relationship element
                const newRelElem = tRelsXml.createElementNS(PKG_REL_NS, 'Relationship');
                newRelElem.setAttribute('Id', newRid);
                newRelElem.setAttribute('Type', mediaRel.type);
                newRelElem.setAttribute('Target', targetInRels);
                relsRoot.appendChild(newRelElem);
            }

            this.setProgress(70, 'Building standard question tables with remapped relationships...');

            // Find prototype table in template
            const tTables = getElementsByLocalName(tDocXml, 'tbl');
            if (tTables.length === 0) {
                throw new Error('Template document does not contain a standard table structure.');
            }
            const protoTable = tTables[0];

            // Locate body element
            const bodyList = getElementsByLocalName(tDocXml, 'body');
            if (bodyList.length === 0) {
                throw new Error('Invalid template: <w:body> element not found.');
            }
            const body = bodyList[0];

            // Preserve <w:sectPr> if present
            let sectPrNode = null;
            const sectPrList = getElementsByLocalName(body, 'sectPr');
            if (sectPrList.length > 0) {
                sectPrNode = sectPrList[sectPrList.length - 1].cloneNode(true);
            }

            // Clear body children
            while (body.firstChild) {
                body.removeChild(body.firstChild);
            }

            // Build output tables
            const sortedQNums = Object.keys(qData).map(n => parseInt(n, 10)).sort((a, b) => a - b);
            const previewData = [];

            for (let i = 0; i < sortedQNums.length; i++) {
                const qnum = sortedQNums[i];
                const q = qData[qnum];
                const ansLetter = answerMap[qnum] || 'a';
                const ansValue = LETTER_TO_NUM[ansLetter] || (ansLetter.match(/^[1-4]$/) ? ansLetter : '1');
                const solCell = solutionMap[qnum] || null;

                const qTable = this.createQuestionTable(
                    protoTable,
                    tDocXml,
                    qnum,
                    q,
                    ansValue,
                    solCell,
                    oldRidToNewRid
                );

                body.appendChild(qTable);

                // Add separator paragraph
                const sepP = tDocXml.createElementNS(W_NS, 'w:p');
                body.appendChild(sepP);

                // Store preview item
                previewData.push({
                    qnum: qnum,
                    questionText: getNodeText(q.qCell).trim(),
                    options: q.options.map(opt => opt ? getNodeText(opt).trim() : ''),
                    answerKey: ansValue,
                    answerLetter: ansLetter,
                    solutionText: solCell ? getNodeText(solCell).trim() : ''
                });

                if (i % 10 === 0 || i === sortedQNums.length - 1) {
                    const pct = 70 + Math.floor((i / sortedQNums.length) * 20);
                    this.setProgress(pct, `Generating Question Table ${i + 1} of ${sortedQNums.length}...`);
                }
            }

            // Restore <w:sectPr> at the end of body
            if (sectPrNode) {
                body.appendChild(sectPrNode);
            }

            this.setProgress(92, 'Serializing XML and compressing final package...');

            const serializer = new XMLSerializer();
            const finalDocXmlStr = serializer.serializeToString(tDocXml);
            const finalRelsXmlStr = serializer.serializeToString(tRelsXml);
            const finalCtXmlStr = serializer.serializeToString(ctXml);

            templateZip.file('word/document.xml', finalDocXmlStr);
            templateZip.file('word/_rels/document.xml.rels', finalRelsXmlStr);
            templateZip.file('[Content_Types].xml', finalCtXmlStr);

            // Generate docx blob
            const outBlob = await templateZip.generateAsync({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                compression: 'DEFLATE',
                compressionOptions: { level: 6 }
            });

            this.setProgress(100, 'Conversion completed successfully!');

            return {
                blob: outBlob,
                stats: {
                    totalQuestions: sortedQNums.length,
                    totalAnswersFound: Object.keys(answerMap).length,
                    totalSolutionsFound: Object.keys(solutionMap).length,
                    mediaRemapped: Object.keys(oldRidToNewRid).length,
                    fileSizeBytes: outBlob.size
                },
                preview: previewData
            };
        }

        /**
         * Parse question table rows into structured question objects
         */
        parseQuestions(table, parser) {
            const rows = getElementsByLocalName(table, 'tr');
            const qRowIndices = {};

            for (let ri = 0; ri < rows.length; ri++) {
                const cells = getElementsByLocalName(rows[ri], 'tc');
                if (cells.length > 0) {
                    const txt = getNodeText(cells[0]).trim();
                    const m = txt.match(/^(\d+)\.?$/);
                    if (m) {
                        const qnum = parseInt(m[1], 10);
                        qRowIndices[qnum] = ri;
                    }
                }
            }

            const sortedQNums = Object.keys(qRowIndices).map(n => parseInt(n, 10)).sort((a, b) => a - b);
            const questions = {};

            for (let i = 0; i < sortedQNums.length; i++) {
                const qnum = sortedQNums[i];
                const ri = qRowIndices[qnum];
                const rowCells = getElementsByLocalName(rows[ri], 'tc');
                const qCell = rowCells.length > 1 ? rowCells[1] : rowCells[0];

                const nextRi = (i + 1 < sortedQNums.length) ? qRowIndices[sortedQNums[i + 1]] : rows.length;
                const optRows = [];
                for (let r = ri + 1; r < nextRi; r++) {
                    optRows.push(rows[r]);
                }

                let opt_a = null, opt_b = null, opt_c = null, opt_d = null;

                if (optRows.length >= 1) {
                    const c1 = getElementsByLocalName(optRows[0], 'tc');
                    if (c1.length > 2) opt_a = c1[2];
                    if (c1.length > 4) opt_b = c1[4];
                }
                if (optRows.length >= 2) {
                    const c2 = getElementsByLocalName(optRows[1], 'tc');
                    if (c2.length > 2) opt_c = c2[2];
                    if (c2.length > 4) opt_d = c2[4];
                }

                // Fallback for inline options (e.g. "a) 4 units  b) 1 unit  c) 2 units  d) 3 units")
                if (!opt_a && !opt_b && !opt_c && !opt_d) {
                    const qText = getNodeText(qCell);
                    const inlineMatch = qText.match(/a\)\s*([\s\S]+?)\s*b\)\s*([\s\S]+?)\s*c\)\s*([\s\S]+?)\s*d\)\s*([\s\S]+?)$/i);
                    if (inlineMatch) {
                        opt_a = this.createSyntheticTc(parser, inlineMatch[1].trim());
                        opt_b = this.createSyntheticTc(parser, inlineMatch[2].trim());
                        opt_c = this.createSyntheticTc(parser, inlineMatch[3].trim());
                        opt_d = this.createSyntheticTc(parser, inlineMatch[4].trim());
                    }
                }

                questions[qnum] = {
                    qCell: qCell,
                    options: [opt_a, opt_b, opt_c, opt_d]
                };
            }

            return questions;
        }

        /**
         * Parse Answer Key table or document text for answers
         */
        parseAnswerKey(tables, docXml) {
            const answerMap = {};

            // 1. Look through all tables for answer pattern "1)b 2)b 3)b..."
            for (let ti = 1; ti < tables.length; ti++) {
                const tblText = getNodeText(tables[ti]);
                const matches = tblText.matchAll(/(\d+)\)\s*([a-dA-D1-4])/g);
                let foundAny = false;
                for (const m of matches) {
                    const qnum = parseInt(m[1], 10);
                    answerMap[qnum] = m[2].toLowerCase();
                    foundAny = true;
                }
                if (foundAny) break;
            }

            // 2. If answerMap is still empty, search entire document
            if (Object.keys(answerMap).length === 0) {
                const fullText = getNodeText(docXml);
                const matches = fullText.matchAll(/(\d+)\)\s*([a-dA-D1-4])/g);
                for (const m of matches) {
                    const qnum = parseInt(m[1], 10);
                    answerMap[qnum] = m[2].toLowerCase();
                }
            }

            return answerMap;
        }

        /**
         * Parse Solutions table
         */
        parseSolutions(tables) {
            const solutions = {};

            // Find solutions table (usually table 4, or table following 'HINTS AND SOLUTIONS')
            let solTable = null;
            for (let ti = tables.length - 1; ti >= 0; ti--) {
                const rows = getElementsByLocalName(tables[ti], 'tr');
                if (rows.length > 5) {
                    // Check if row 1 cells[0] has number 1
                    const c = getElementsByLocalName(rows[1], 'tc');
                    if (c.length > 0 && getNodeText(c[0]).trim().startsWith('1')) {
                        solTable = tables[ti];
                        break;
                    }
                }
            }

            if (solTable) {
                const rows = getElementsByLocalName(solTable, 'tr');
                for (let ri = 1; ri < rows.length; ri++) {
                    const cells = getElementsByLocalName(rows[ri], 'tc');
                    if (cells.length > 1) {
                        const txt = getNodeText(cells[0]).trim().replace(/\.$/, '');
                        const qnum = parseInt(txt, 10);
                        if (!isNaN(qnum)) {
                            solutions[qnum] = cells[1];
                        }
                    }
                }
            }

            return solutions;
        }

        /**
         * Remap relationship IDs in element attributes
         */
        remapRelationshipIds(elem, oldRidToNewRid) {
            if (!elem || !elem.attributes) return;

            // Check attributes of current element
            for (let i = 0; i < elem.attributes.length; i++) {
                const attr = elem.attributes[i];
                const val = attr.value;
                if (val && oldRidToNewRid[val]) {
                    attr.value = oldRidToNewRid[val];
                }
            }

            // Check children recursively
            if (elem.children) {
                for (let i = 0; i < elem.children.length; i++) {
                    this.remapRelationshipIds(elem.children[i], oldRidToNewRid);
                }
            }
        }

        /**
         * Copy all paragraphs from a source cell to destination cell
         */
        copyCellParagraphs(srcCell, dstTc, docXml, oldRidToNewRid) {
            // Remove existing paragraphs in dstTc
            const existingPs = getElementsByLocalName(dstTc, 'p');
            for (let i = existingPs.length - 1; i >= 0; i--) {
                const p = existingPs[i];
                if (p.parentNode === dstTc) {
                    dstTc.removeChild(p);
                }
            }

            if (!srcCell) {
                const emptyP = docXml.createElementNS(W_NS, 'w:p');
                dstTc.appendChild(emptyP);
                return;
            }

            // Find all direct <w:p> in srcCell
            const srcPs = getElementsByLocalName(srcCell, 'p');
            let addedCount = 0;

            for (let i = 0; i < srcPs.length; i++) {
                const srcP = srcPs[i];
                if (srcP.parentNode === srcCell || srcPs.length <= 5) {
                    const newP = docXml.importNode ? docXml.importNode(srcP, true) : srcP.cloneNode(true);
                    this.remapRelationshipIds(newP, oldRidToNewRid);
                    dstTc.appendChild(newP);
                    addedCount++;
                }
            }

            if (addedCount === 0) {
                const emptyP = docXml.createElementNS(W_NS, 'w:p');
                dstTc.appendChild(emptyP);
            }
        }

        /**
         * Set simple text inside a destination cell
         */
        setCellText(dstTc, docXml, text) {
            const existingPs = getElementsByLocalName(dstTc, 'p');
            for (let i = existingPs.length - 1; i >= 0; i--) {
                const p = existingPs[i];
                if (p.parentNode === dstTc) {
                    dstTc.removeChild(p);
                }
            }

            const p = docXml.createElementNS(W_NS, 'w:p');
            const r = docXml.createElementNS(W_NS, 'w:r');
            const t = docXml.createElementNS(W_NS, 'w:t');
            t.textContent = String(text);
            r.appendChild(t);
            p.appendChild(r);
            dstTc.appendChild(p);
        }

        /**
         * Create a synthetic cell node for inlined options
         */
        createSyntheticTc(parser, text) {
            const xml = `<w:tc xmlns:w="${W_NS}"><w:p><w:r><w:t>${text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</w:t></w:r></w:p></w:tc>`;
            const doc = parser.parseFromString(xml, 'application/xml');
            return doc.documentElement;
        }

        /**
         * Ensure ContentType Default entry exists
         */
        ensureContentTypeDefault(ctXml, ext, contentType) {
            const root = ctXml.documentElement;
            const defaults = getElementsByLocalName(ctXml, 'Default');
            for (let i = 0; i < defaults.length; i++) {
                if (defaults[i].getAttribute('Extension') === ext) {
                    return;
                }
            }
            const el = ctXml.createElementNS(CT_NS, 'Default');
            el.setAttribute('Extension', ext);
            el.setAttribute('ContentType', contentType);
            root.appendChild(el);
        }

        /**
         * Ensure ContentType Override entry exists
         */
        ensureContentTypeOverride(ctXml, partName, contentType) {
            const root = ctXml.documentElement;
            const overrides = getElementsByLocalName(ctXml, 'Override');
            for (let i = 0; i < overrides.length; i++) {
                if (overrides[i].getAttribute('PartName') === partName) {
                    return;
                }
            }
            const el = ctXml.createElementNS(CT_NS, 'Override');
            el.setAttribute('PartName', partName);
            el.setAttribute('ContentType', contentType);
            root.appendChild(el);
        }

        /**
         * Create a complete 10-row question table matching output_format.docx
         */
        createQuestionTable(protoTable, docXml, qnum, qData, answerValue, solCell, oldRidToNewRid) {
            const tbl = docXml.importNode ? docXml.importNode(protoTable, true) : protoTable.cloneNode(true);
            const rows = getElementsByLocalName(tbl, 'tr');

            const getTcs = (tr) => getElementsByLocalName(tr, 'tc');

            // Row 0 - Question
            if (rows.length > 0) {
                const tcs = getTcs(rows[0]);
                if (tcs.length > 1) {
                    this.copyCellParagraphs(qData.qCell, tcs[1], docXml, oldRidToNewRid);
                }
            }

            // Row 1 - Type
            if (rows.length > 1) {
                const tcs = getTcs(rows[1]);
                if (tcs.length > 1) {
                    this.setCellText(tcs[1], docXml, 'multiple_choice');
                }
            }

            // Rows 2-5 - Options (a, b, c, d)
            for (let i = 0; i < 4; i++) {
                const rowIdx = 2 + i;
                if (rows.length > rowIdx) {
                    const tcs = getTcs(rows[rowIdx]);
                    if (tcs.length > 1) {
                        const optCell = qData.options[i];
                        if (optCell) {
                            this.copyCellParagraphs(optCell, tcs[1], docXml, oldRidToNewRid);
                        } else {
                            this.setCellText(tcs[1], docXml, '');
                        }
                    }
                }
            }

            // Row 6 - Answer
            if (rows.length > 6) {
                const tcs = getTcs(rows[6]);
                if (tcs.length > 1) {
                    this.setCellText(tcs[1], docXml, answerValue || '1');
                }
            }

            // Row 7 - Solution
            if (rows.length > 7) {
                const tcs = getTcs(rows[7]);
                if (tcs.length > 1) {
                    if (solCell) {
                        this.copyCellParagraphs(solCell, tcs[1], docXml, oldRidToNewRid);
                    } else {
                        this.setCellText(tcs[1], docXml, '');
                    }
                }
            }

            // Row 8 - Positive Marks
            if (rows.length > 8) {
                const tcs = getTcs(rows[8]);
                if (tcs.length > 1) {
                    this.setCellText(tcs[1], docXml, this.positiveMarks);
                }
            }

            // Row 9 - Negative Marks
            if (rows.length > 9) {
                const tcs = getTcs(rows[9]);
                if (tcs.length > 1) {
                    this.setCellText(tcs[1], docXml, this.negativeMarks);
                }
            }

            return tbl;
        }
    }

    window.DocxConverter = DocxConverter;

})(typeof window !== 'undefined' ? window : this);

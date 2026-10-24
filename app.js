/**
 * Application Controller for DOCX Converter Web App
 */

document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfoPill = document.getElementById('fileInfoPill');
    const fileNameSpan = document.getElementById('fileName');
    const fileSizeSpan = document.getElementById('fileSize');
    const removeFileBtn = document.getElementById('removeFileBtn');

    const templateInput = document.getElementById('templateInput');
    const templateStatus = document.getElementById('templateStatus');
    const positiveMarksInput = document.getElementById('positiveMarks');
    const negativeMarksInput = document.getElementById('negativeMarks');

    const convertBtn = document.getElementById('convertBtn');
    const loadSampleBtn = document.getElementById('loadSampleBtn');

    const progressCard = document.getElementById('progressCard');
    const progressStatus = document.getElementById('progressStatus');
    const progressPercent = document.getElementById('progressPercent');
    const progressBarFill = document.getElementById('progressBarFill');

    const resultsCard = document.getElementById('resultsCard');
    const downloadBtn = document.getElementById('downloadBtn');

    const statQuestions = document.getElementById('statQuestions');
    const statAnswers = document.getElementById('statAnswers');
    const statSolutions = document.getElementById('statSolutions');
    const statMedia = document.getElementById('statMedia');
    const statFileSize = document.getElementById('statFileSize');

    const previewCard = document.getElementById('previewCard');
    const questionsList = document.getElementById('questionsList');
    const searchQuestionsInput = document.getElementById('searchQuestions');

    // App State
    let selectedFile = null;
    let customTemplateFile = null;
    let convertedBlob = null;
    let previewDataCache = [];

    // Helper: format bytes
    function formatBytes(bytes) {
        if (!bytes || bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // Set selected file
    function setFile(file) {
        if (!file) return;
        if (!file.name.toLowerCase().endsWith('.docx')) {
            alert('Please select a valid .docx document.');
            return;
        }
        selectedFile = file;
        fileNameSpan.textContent = file.name;
        fileSizeSpan.textContent = formatBytes(file.size);
        fileInfoPill.style.display = 'flex';
        convertBtn.disabled = false;

        // Reset previous conversion results if any
        resultsCard.style.display = 'none';
        previewCard.style.display = 'none';
        progressCard.style.display = 'none';
    }

    // Remove file
    function removeFile() {
        selectedFile = null;
        fileInput.value = '';
        fileInfoPill.style.display = 'none';
        convertBtn.disabled = true;
        resultsCard.style.display = 'none';
        previewCard.style.display = 'none';
    }

    // Drag & Drop handlers
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files && files.length > 0) {
            setFile(files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) {
            setFile(e.target.files[0]);
        }
    });

    removeFileBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile();
    });

    // Custom Template handler
    if (templateInput) {
        templateInput.addEventListener('change', (e) => {
            if (e.target.files && e.target.files.length > 0) {
                customTemplateFile = e.target.files[0];
                if (templateStatus) {
                    templateStatus.textContent = `Custom: ${customTemplateFile.name}`;
                    templateStatus.style.color = '#a5b4fc';
                }
            }
        });
    }

    // Load Sample Button Handler
    if (loadSampleBtn) {
        loadSampleBtn.addEventListener('click', async () => {
            try {
                loadSampleBtn.disabled = true;
                loadSampleBtn.textContent = 'Loading sample...';

                // Try fetching input_format.docx
                const res = await fetch('input_format.docx');
                if (!res.ok) throw new Error('Sample input_format.docx not found at root.');
                const blob = await res.blob();
                const file = new File([blob], 'input_format.docx', {
                    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                });
                setFile(file);
                loadSampleBtn.innerHTML = `
                    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" />
                    </svg>
                    Sample Loaded
                `;
            } catch (err) {
                console.warn('Sample auto-load notice:', err);
                alert('To use the sample, you can select input_format.docx from your folder.');
                loadSampleBtn.innerHTML = `
                    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                    </svg>
                    Try Sample Document
                `;
            } finally {
                loadSampleBtn.disabled = false;
            }
        });
    }

    // Render Preview
    function renderPreview(items, query = '') {
        questionsList.innerHTML = '';
        const filter = query.toLowerCase().trim();

        const filtered = items.filter(it => {
            if (!filter) return true;
            return (
                it.qnum.toString().includes(filter) ||
                it.questionText.toLowerCase().includes(filter) ||
                it.solutionText.toLowerCase().includes(filter) ||
                it.options.some(opt => opt.toLowerCase().includes(filter))
            );
        });

        if (filtered.length === 0) {
            questionsList.innerHTML = `
                <div style="text-align: center; padding: 24px; color: var(--text-muted);">
                    No questions match your filter.
                </div>
            `;
            return;
        }

        const fragment = document.createDocumentFragment();

        filtered.forEach(it => {
            const card = document.createElement('div');
            card.className = 'q-item-card';

            const optLetters = ['a', 'b', 'c', 'd'];
            const optRowsHtml = it.options.map((optText, idx) => {
                const optNum = (idx + 1).toString();
                const isCorrect = (it.answerKey === optNum || it.answerLetter === optLetters[idx]);
                return `
                    <div class="q-opt ${isCorrect ? 'is-correct' : ''}">
                        <strong>${optLetters[idx]})</strong> ${escapeHtml(optText || '(empty)')}
                        ${isCorrect ? ' <span style="font-size:0.75rem; float:right;">✓ Correct</span>' : ''}
                    </div>
                `;
            }).join('');

            card.innerHTML = `
                <div class="q-item-top">
                    <span class="q-num-pill">Question ${it.qnum}</span>
                    <span class="q-ans-pill">Answer: Option ${it.answerKey} (${it.answerLetter.toUpperCase()})</span>
                </div>
                <div class="q-text">${escapeHtml(it.questionText)}</div>
                <div class="q-options-grid">${optRowsHtml}</div>
                ${it.solutionText ? `
                    <div class="q-sol-preview">
                        <strong>Solution:</strong> ${escapeHtml(it.solutionText.substring(0, 160))}${it.solutionText.length > 160 ? '...' : ''}
                    </div>
                ` : ''}
            `;

            fragment.appendChild(card);
        });

        questionsList.appendChild(fragment);
    }

    function escapeHtml(str) {
        if (!str) return '';
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    // Search filter event
    searchQuestionsInput.addEventListener('input', (e) => {
        renderPreview(previewDataCache, e.target.value);
    });

    // Convert Button Handler
    convertBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        try {
            convertBtn.disabled = true;
            progressCard.style.display = 'block';
            resultsCard.style.display = 'none';
            previewCard.style.display = 'none';

            const posMarks = positiveMarksInput.value.trim() || '4';
            const negMarks = negativeMarksInput.value.trim() || '1';

            const converter = new DocxConverter({
                positiveMarks: posMarks,
                negativeMarks: negMarks,
                onProgress: ({ percent, message }) => {
                    progressPercent.textContent = `${percent}%`;
                    progressStatus.textContent = message;
                    progressBarFill.style.width = `${percent}%`;
                }
            });

            // Run conversion
            const result = await converter.convert(selectedFile, customTemplateFile);
            convertedBlob = result.blob;
            previewDataCache = result.preview;

            // Fill Stats
            statQuestions.textContent = result.stats.totalQuestions;
            statAnswers.textContent = result.stats.totalAnswersFound;
            statSolutions.textContent = result.stats.totalSolutionsFound;
            statMedia.textContent = result.stats.mediaRemapped;
            statFileSize.textContent = formatBytes(result.stats.fileSizeBytes);

            // Render Preview
            renderPreview(previewDataCache);

            // Show Cards
            resultsCard.style.display = 'block';
            previewCard.style.display = 'block';

            // Scroll to results
            resultsCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

        } catch (error) {
            console.error('Conversion error:', error);
            alert(`Conversion failed: ${error.message || error}`);
            progressStatus.textContent = 'Error during conversion.';
            progressStatus.style.color = '#f43f5e';
        } finally {
            convertBtn.disabled = false;
        }
    });

    // Download Button Handler
    downloadBtn.addEventListener('click', () => {
        if (!convertedBlob) return;
        const url = URL.createObjectURL(convertedBlob);
        const a = document.createElement('a');
        const originalName = selectedFile ? selectedFile.name.replace(/\.[^/.]+$/, '') : 'document';
        a.href = url;
        a.download = `${originalName}_converted_output.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    });
});

// PDF.js Worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// State
const state = {
    pdfDoc: null,
    pageNum: 1,
    pageRendering: false,
    pageNumPending: null,
    scale: 1.0,
    annotations: [], // { id, type, pageNumber, x, y, width, height, color, size, category, text/comment, stampType, font }
    undoStack: [],
    redoStack: [],
    selectedAnnotationId: null,
    mode: 'point', // point, text, highlight, stamp
    dragStart: null, // {x, y} for highlight dragging
    currentSettings: {
        color: 'blue',
        size: 12, // Default Medium
        stampType: 'ok',
        font: 'Noto Sans JP' // Default Font
    },
    isGraphMode: false,
    selectedDotIds: new Set(),
    pdfBuffer: null,
    // Drag & Drop State
    isDragging: false,
    draggedAnnotationId: null,
    dragOffset: { x: 0, y: 0 },
    editingId: null // Track which annotation is being edited
};

// DOM Elements
const canvas = document.getElementById('the-canvas');
const ctx = canvas.getContext('2d');
const annotationLayer = document.getElementById('annotation-layer');
const pdfWrapper = document.getElementById('pdf-wrapper');
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const pageNumSpan = document.getElementById('page-num');
const pageCountSpan = document.getElementById('page-count');
const prevPageBtn = document.getElementById('prev-page');
const nextPageBtn = document.getElementById('next-page');

const printBtn = document.getElementById('print-btn');
const commentList = document.getElementById('comment-list');
const undoBtn = document.getElementById('undo-btn');
const redoBtn = document.getElementById('redo-btn');
const exportCsvBtn = document.getElementById('export-csv');
const stampControls = document.getElementById('stamp-controls');
const styleControls = document.getElementById('style-controls');
const imageControls = document.getElementById('image-controls');
const stampTypeSelect = document.getElementById('stamp-type');
const graphModeCheckbox = document.getElementById('graph-mode-chk');
const pointSizeControl = document.getElementById('point-size-control');
const textSizeControl = document.getElementById('text-size-control');
const fontSizeInput = document.getElementById('font-size-input');
const fontSelect = document.getElementById('font-select');
const imageInput = document.getElementById('image-input');

// --- Initialization ---

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file && file.type === 'application/pdf') loadPdf(file);
});

// Drag & Drop Handling (Global)
document.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (!state.pdfDoc) {
        dropZone.style.backgroundColor = 'rgba(230, 240, 255, 0.9)';
    }
});

document.addEventListener('dragleave', (e) => {
    e.preventDefault();
    if (!state.pdfDoc) {
        dropZone.style.backgroundColor = 'rgba(255, 255, 255, 0.95)';
    }
});

document.addEventListener('drop', (e) => {
    e.preventDefault();
    if (!state.pdfDoc) {
        dropZone.style.backgroundColor = 'rgba(255, 255, 255, 0.95)';
    }

    const file = e.dataTransfer.files[0];
    if (file) {
        if (file.type === 'application/pdf') {
            loadPdf(file);
        } else if (file.type === 'image/png' || file.type === 'image/jpeg' || file.type === 'image/jpg') {
            // If PDF is loaded, add image
            if (state.pdfDoc) {
                loadImage(file, (dataUrl, width, height) => {
                    // Add to mouse position
                    const rect = annotationLayer.getBoundingClientRect();
                    const x = e.clientX - rect.left;
                    const y = e.clientY - rect.top;

                    // Default size: maintain aspect ratio, max 200px width
                    const scale = Math.min(200 / width, 200 / height, 1);
                    const w = width * scale;
                    const h = height * scale;

                    addAnnotation({
                        type: 'image',
                        x: x - w / 2,
                        y: y - h / 2,
                        width: w,
                        height: h,
                        imageData: dataUrl,
                        imageType: file.type
                    });
                });
            } else {
                alert('先にPDFを開いてください。');
            }
        } else {
            // Only alert if it's not a supported file type and we are trying to do something
            // But to be safe and not annoying, maybe just log or ignore if it's not what we want?
            // The user specifically asked for this fix, so let's be explicit.
            // However, if the user drops something else, we should prevent default anyway (done above).
        }
    }
});

prevPageBtn.addEventListener('click', () => changePage(-1));
nextPageBtn.addEventListener('click', () => changePage(1));

printBtn.addEventListener('click', () => window.print());
exportCsvBtn.addEventListener('click', exportCSV);
undoBtn.addEventListener('click', undo);
redoBtn.addEventListener('click', redo);

// Mode Switching
document.querySelectorAll('input[name="mode"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        state.mode = e.target.value;
        // Force disable graph mode when switching to other modes to prevent confusion
        if (state.mode !== 'point') { // Graph mode only makes sense with points
            state.isGraphMode = false;
            graphModeCheckbox.checked = false;
        }
        updateUIForMode();
        state.selectedAnnotationId = null;
        renderAnnotations();
    });
});

function updateUIForMode() {
    stampControls.style.display = 'none';
    styleControls.style.display = 'none';
    imageControls.style.display = 'none';

    if (state.mode === 'stamp') {
        stampControls.style.display = 'flex';
    } else if (state.mode === 'image') {
        imageControls.style.display = 'flex';
    } else if (state.mode === 'highlight' || state.mode === 'whiteout') {
        // No controls
    } else {
        styleControls.style.display = 'flex';

        // Toggle between Point (S/M/L) and Text (Numeric) size controls
        if (state.mode === 'text') {
            pointSizeControl.style.display = 'none';
            textSizeControl.style.display = 'inline';
            fontSelect.style.display = 'inline-block'; // Show font select
        } else {
            pointSizeControl.style.display = 'inline';
            textSizeControl.style.display = 'none';
            fontSelect.style.display = 'none'; // Hide font select
        }
    }
}

// Settings
// Point Size (S/M/L)
document.querySelectorAll('input[name="size"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        const newSize = parseInt(e.target.value);
        state.currentSettings.size = newSize;

        // Only apply to Point annotations if selected
        if (state.selectedAnnotationId) {
            const ann = state.annotations.find(a => a.id === state.selectedAnnotationId);
            if (ann && ann.type === 'point') {
                ann.size = newSize;
                renderAnnotations();
            }
        }
    });
});

// Text Font Size (Numeric)
fontSizeInput.addEventListener('input', (e) => {
    const newSize = parseInt(e.target.value);
    if (newSize > 0) {
        state.currentSettings.size = newSize;

        // If editing, update the input box immediately
        if (state.editingId) {
            const input = document.querySelector('.annotation-input');
            if (input) {
                input.style.fontSize = `${newSize}px`;
            }
            const ann = state.annotations.find(a => a.id === state.editingId);
            if (ann) ann.size = newSize;
        }

        // If selected (Text), update
        if (state.selectedAnnotationId) {
            const ann = state.annotations.find(a => a.id === state.selectedAnnotationId);
            if (ann && ann.type === 'text') {
                ann.size = newSize;
                renderAnnotations();
            }
        }
    }
});

// Font Selection
fontSelect.addEventListener('change', (e) => {
    const newFont = e.target.value;
    state.currentSettings.font = newFont;

    // If editing, update the input box immediately
    if (state.editingId) {
        const input = document.querySelector('.annotation-input');
        if (input) {
            input.style.fontFamily = newFont;
            // Also update the state object so it persists if we save
            const ann = state.annotations.find(a => a.id === state.editingId);
            if (ann) ann.font = newFont;
        }
    }

    // If selected (Text), update
    if (state.selectedAnnotationId) {
        const ann = state.annotations.find(a => a.id === state.selectedAnnotationId);
        if (ann && ann.type === 'text') {
            ann.font = newFont;
            renderAnnotations();
        }
    }
});

document.querySelectorAll('input[name="color"]').forEach(r => r.addEventListener('change', e => {
    state.currentSettings.color = e.target.value;

    // If editing, update the input box immediately
    if (state.editingId) {
        const input = document.querySelector('.annotation-input');
        if (input) {
            const colorName = state.currentSettings.color;
            if (colorName === 'blue') input.style.color = 'var(--blue-color)';
            else if (colorName === 'red') input.style.color = 'var(--red-color)';
            else input.style.color = '#333';

            // Also update the state object
            const ann = state.annotations.find(a => a.id === state.editingId);
            if (ann) ann.color = colorName;
        }
    }

    if (state.selectedAnnotationId) {
        pushState();
        const ann = state.annotations.find(a => a.id === state.selectedAnnotationId);
        if (ann) {
            ann.color = state.currentSettings.color;
            renderAnnotations();
        }
    }
}));

stampTypeSelect.addEventListener('change', e => state.currentSettings.stampType = e.target.value);

imageInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        loadImage(file, (dataUrl, width, height) => {
            // Add to center of screen
            const rect = annotationLayer.getBoundingClientRect();
            const w = Math.min(200, width);
            const h = w * (height / width);

            addAnnotation({
                type: 'image',
                x: rect.width / 2 - w / 2,
                y: rect.height / 2 - h / 2,
                width: w,
                height: h,
                imageData: dataUrl,
                imageType: file.type
            });
        });
        imageInput.value = ''; // Reset
    }
});

function loadImage(file, callback) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
            callback(e.target.result, img.width, img.height);
        };
        img.src = e.target.result;
    };
    reader.readAsDataURL(file);
}

graphModeCheckbox.addEventListener('change', (e) => {
    state.isGraphMode = e.target.checked;
    renderAnnotations();
});

// Keyboard Shortcuts
document.addEventListener('keydown', (e) => {
    if (e.key === 'Delete' || e.key === 'Backspace') {
        if (state.selectedAnnotationId) {
            deleteAnnotation(state.selectedAnnotationId);
        }
    }
    if ((e.metaKey || e.ctrlKey) && e.key === 'z') {
        e.preventDefault();
        if (e.shiftKey) redo();
        else undo();
    }
});

// --- PDF Logic ---

function loadPdf(file) {
    const fileReader = new FileReader();
    fileReader.onload = function () {
        const typedarray = new Uint8Array(this.result);
        state.pdfBuffer = this.result.slice(0);
        pdfjsLib.getDocument(typedarray).promise.then((pdfDoc_) => {
            state.pdfDoc = pdfDoc_;
            pageCountSpan.textContent = state.pdfDoc.numPages;
            state.pageNum = 1;
            state.annotations = [];
            state.undoStack = [];
            state.redoStack = [];
            updateUndoRedoButtons();

            dropZone.classList.add('hidden');
            pdfWrapper.classList.remove('hidden');
            renderPage(state.pageNum);
            renderSidebar();
        }).catch(err => {
            console.error(err);
            alert('PDF読み込みエラー');
        });
    };
    fileReader.readAsArrayBuffer(file);
}

function renderPage(num) {
    state.pageRendering = true;
    state.pdfDoc.getPage(num).then((page) => {
        const containerWidth = document.getElementById('viewer-container').clientWidth - 40;
        const unscaledViewport = page.getViewport({ scale: 1 });
        const scale = containerWidth / unscaledViewport.width;
        state.scale = scale < 1 ? 1 : scale;
        const viewport = page.getViewport({ scale: state.scale });

        canvas.height = viewport.height;
        canvas.width = viewport.width;
        annotationLayer.style.width = `${viewport.width}px`;
        annotationLayer.style.height = `${viewport.height}px`;

        const renderContext = { canvasContext: ctx, viewport: viewport };
        page.render(renderContext).promise.then(() => {
            state.pageRendering = false;
            pageNumSpan.textContent = num;
            renderAnnotations();
            if (state.pageNumPending !== null) {
                renderPage(state.pageNumPending);
                state.pageNumPending = null;
            }
        });
    });
    prevPageBtn.disabled = num <= 1;
    nextPageBtn.disabled = num >= state.pdfDoc.numPages;
}

function changePage(offset) {
    const newPage = state.pageNum + offset;
    if (newPage >= 1 && newPage <= state.pdfDoc.numPages) {
        state.pageNum = newPage;
        renderPage(state.pageNum);
    }
}

// --- Annotation Interaction ---

annotationLayer.addEventListener('mousedown', (e) => {
    // Handle Resize Handles
    if (e.target.classList.contains('resize-handle')) {
        const handle = e.target;
        const parent = handle.closest('.annotation');
        const id = parent.dataset.id;

        state.isResizing = true;
        state.resizingId = id;
        state.resizeHandle = handle.dataset.handle; // nw, ne, sw, se
        state.resizeStart = { x: e.clientX, y: e.clientY };

        // Prevent drag
        e.stopPropagation();
        return;
    }

    if (e.target !== annotationLayer && !e.target.classList.contains('highlight-annotation') && !e.target.classList.contains('annotation') && !e.target.classList.contains('image-annotation')) return;

    // Handle Dragging Start
    if (e.target.closest('.annotation')) {
        const el = e.target.closest('.annotation');
        const id = el.dataset.id;
        const ann = state.annotations.find(a => a.id === id);

        // Don't drag if in graph mode and clicking a dot (selection takes precedence)
        if (state.isGraphMode && ann.type === 'point') return;

        // Allow clicking "through" whiteout/highlight to add text/points on top
        if ((state.mode === 'text' || state.mode === 'point' || state.mode === 'stamp') &&
            (ann.type === 'whiteout' || ann.type === 'highlight')) {
            // Don't start dragging, let the click event handle addition
            return;
        }

        state.isDragging = true;
        state.draggedAnnotationId = id;
        const rect = annotationLayer.getBoundingClientRect();

        // Calculate offset from the annotation's top-left
        // ann.xPct is left, ann.yPct is top
        const annX = ann.xPct * rect.width;
        const annY = ann.yPct * rect.height;

        state.dragOffset = {
            x: (e.clientX - rect.left) - annX,
            y: (e.clientY - rect.top) - annY
        };

        // Select it too
        selectAnnotation(id);
        return;
    }

    if (state.mode === 'highlight' || state.mode === 'whiteout') {
        const rect = annotationLayer.getBoundingClientRect();
        state.dragStart = {
            x: e.clientX - rect.left,
            y: e.clientY - rect.top
        };
    }
});

annotationLayer.addEventListener('mousemove', (e) => {
    const rect = annotationLayer.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    if (state.isResizing && state.resizingId) {
        const ann = state.annotations.find(a => a.id === state.resizingId);
        if (ann) {
            const dx = e.clientX - state.resizeStart.x;
            const dy = e.clientY - state.resizeStart.y;

            // Convert percentages to pixels for calculation
            let curX = ann.xPct * rect.width;
            let curY = ann.yPct * rect.height;
            let curW = ann.wPct * rect.width;
            let curH = ann.hPct * rect.height;

            if (state.resizeHandle.includes('e')) curW += dx;
            if (state.resizeHandle.includes('s')) curH += dy;
            if (state.resizeHandle.includes('w')) { curX += dx; curW -= dx; }
            if (state.resizeHandle.includes('n')) { curY += dy; curH -= dy; }

            // Minimum size
            if (curW < 20) curW = 20;
            if (curH < 20) curH = 20;

            // Update state
            ann.xPct = curX / rect.width;
            ann.yPct = curY / rect.height;
            ann.wPct = curW / rect.width;
            ann.hPct = curH / rect.height;

            state.resizeStart = { x: e.clientX, y: e.clientY };
            renderAnnotations();
        }
        return;
    }

    if (state.isDragging && state.draggedAnnotationId) {
        const ann = state.annotations.find(a => a.id === state.draggedAnnotationId);
        if (ann) {
            // Update position
            // New Left = MouseX - OffsetX
            // Convert back to percentage
            const newX = x - state.dragOffset.x;
            const newY = y - state.dragOffset.y;

            ann.xPct = newX / rect.width;
            ann.yPct = newY / rect.height;

            // Re-render immediately for smooth drag
            renderAnnotations();
        }
        return;
    }

    if ((state.mode === 'highlight' || state.mode === 'whiteout') && state.dragStart) {
        // Visual feedback for dragging
        const rect = annotationLayer.getBoundingClientRect();
        const currentX = e.clientX - rect.left;
        const currentY = e.clientY - rect.top;

        const x = Math.min(state.dragStart.x, currentX);
        const y = Math.min(state.dragStart.y, currentY);
        const width = Math.abs(currentX - state.dragStart.x);
        const height = Math.abs(currentY - state.dragStart.y);

        let selectionBox = document.getElementById('drag-selection-box');
        if (!selectionBox) {
            selectionBox = document.createElement('div');
            selectionBox.id = 'drag-selection-box';
            selectionBox.className = 'drag-selection';
            annotationLayer.appendChild(selectionBox);
        }

        selectionBox.style.left = x + 'px';
        selectionBox.style.top = y + 'px';
        selectionBox.style.width = width + 'px';
        selectionBox.style.height = height + 'px';
    }
});

annotationLayer.addEventListener('mouseup', (e) => {
    if (state.isResizing) {
        state.isResizing = false;
        state.resizingId = null;
        pushState();
        return;
    }

    if (state.isDragging) {
        state.isDragging = false;
        state.draggedAnnotationId = null;
        pushState(); // Save state after drag
        renderSidebar(); // Update sidebar if order changed (though we sort by Y usually)
        return;
    }

    if ((state.mode === 'highlight' || state.mode === 'whiteout') && state.dragStart) {
        const rect = annotationLayer.getBoundingClientRect();
        const currentX = e.clientX - rect.left;
        const currentY = e.clientY - rect.top;

        const x = Math.min(state.dragStart.x, currentX);
        const y = Math.min(state.dragStart.y, currentY);
        const width = Math.abs(currentX - state.dragStart.x);
        const height = Math.abs(currentY - state.dragStart.y);

        if (width > 5 && height > 5) {
            addAnnotation({
                type: state.mode, // 'highlight' or 'whiteout'
                x, y, width, height,
                color: state.mode === 'highlight' ? 'yellow' : 'white',
                comment: ''
            });
        }
        state.dragStart = null;

        // Remove selection box
        const selectionBox = document.getElementById('drag-selection-box');
        if (selectionBox) selectionBox.remove();
    }
});

annotationLayer.addEventListener('click', (e) => {
    // Debug Info
    console.log('Click on annotationLayer', e.target, state.mode, state.isGraphMode);

    // If in Graph Mode, handle dot selection
    if (state.isGraphMode) {
        if (e.target.classList.contains('dot')) {
            const id = e.target.dataset.id;
            if (state.selectedDotIds.has(id)) {
                state.selectedDotIds.delete(id);
            } else {
                state.selectedDotIds.add(id);
            }
            renderAnnotations();
        }
        return;
    }

    // If clicked on an existing annotation, select it (handled by bubble up or separate listener)
    if (e.target.closest('.annotation')) {
        const el = e.target.closest('.annotation');
        const id = el.dataset.id;
        const ann = state.annotations.find(a => a.id === id);

        // Allow clicking "through" whiteout/highlight to add text/points on top
        if ((state.mode === 'text' || state.mode === 'point' || state.mode === 'stamp') &&
            (ann.type === 'whiteout' || ann.type === 'highlight')) {
            // Pass through -> Don't return, so we can add annotation below
        } else {
            // Prevent adding new annotation if clicking existing (normal behavior)
            return;
        }
    }

    // If clicked on background
    if (state.mode === 'point' || state.mode === 'text' || state.mode === 'stamp') {
        const rect = annotationLayer.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;

        if (state.mode === 'point') {
            addAnnotation({
                type: 'point',
                x, y,
                color: state.currentSettings.color,
                size: state.currentSettings.size,
                comment: ''
            });
        } else if (state.mode === 'text') {
            // Create input immediately
            showTextInput(x, y);
        } else if (state.mode === 'stamp') {
            addAnnotation({
                type: 'stamp',
                x, y,
                stampType: state.currentSettings.stampType,
                comment: '' // Optional
            });
        }
    } else {
        // Deselect if clicking empty space
        if (state.selectedAnnotationId) {
            state.selectedAnnotationId = null;
            renderAnnotations();
        }
    }
});

function addAnnotation(data) {
    pushState();
    const id = Date.now().toString();
    // Use annotationLayer dimensions for percentage calculation because canvas dimensions might be different (high DPI or responsive)
    const layerWidth = annotationLayer.clientWidth || canvas.width;
    const layerHeight = annotationLayer.clientHeight || canvas.height;

    if (!layerWidth || !layerHeight) {
        console.error('Annotation layer dimensions are invalid:', layerWidth, layerHeight);
        alert('エラー: 描画領域のサイズが取得できませんでした。ページを再読み込みしてください。');
        return;
    }

    const annotation = {
        id,
        pageNumber: state.pageNum,
        xPct: data.x / layerWidth,
        yPct: data.y / layerHeight,
        wPct: data.width ? data.width / layerWidth : 0,
        hPct: data.height ? data.height / layerHeight : 0,
        ...data
    };
    if (isNaN(annotation.xPct) || isNaN(annotation.yPct)) {
        console.error('Invalid annotation coordinates:', annotation);
        return;
    }

    state.annotations.push(annotation);
    try {
        renderAnnotations();
    } catch (e) {
        console.error('Error rendering annotations:', e);
    }
    renderSidebar();
}

function showTextInput(x, y, existingId = null) {
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'annotation-input';
    input.style.left = x + 'px';
    input.style.top = y + 'px';

    // Apply current size to input
    const fontSize = existingId
        ? (state.annotations.find(a => a.id === existingId).size || 14)
        : state.currentSettings.size;
    input.style.fontSize = `${fontSize}px`;

    // Apply current font to input
    const font = existingId
        ? (state.annotations.find(a => a.id === existingId).font || 'Noto Sans JP')
        : state.currentSettings.font;
    input.style.fontFamily = font;

    // Apply current color to input
    const colorName = existingId
        ? (state.annotations.find(a => a.id === existingId).color || 'black')
        : state.currentSettings.color;

    if (colorName === 'blue') input.style.color = 'var(--blue-color)';
    else if (colorName === 'red') input.style.color = 'var(--red-color)';
    else input.style.color = '#333';

    if (existingId) {
        state.editingId = existingId; // Set editing state
        const ann = state.annotations.find(a => a.id === existingId);
        input.value = ann.text || '';
        renderAnnotations(); // Re-render to hide the underlying text
    }

    let isComposing = false;
    let isSaved = false; // Flag to prevent double save (Enter + Blur)

    input.addEventListener('compositionstart', () => isComposing = true);
    input.addEventListener('compositionend', () => isComposing = false);

    const saveAndClose = () => {
        if (isSaved) return;
        isSaved = true;

        const text = input.value.trim();
        if (text) {
            if (existingId) {
                pushState();
                const ann = state.annotations.find(a => a.id === existingId);
                if (ann) {
                    ann.text = text;
                    ann.font = state.currentSettings.font; // Update font if changed
                }
            } else {
                addAnnotation({
                    type: 'text',
                    x, y,
                    text: text,
                    color: state.currentSettings.color,
                    size: state.currentSettings.size,
                    font: state.currentSettings.font
                });
            }
        }

        input.remove();
        state.editingId = null; // Clear editing state
        renderAnnotations();
        renderSidebar();
    };

    input.onkeydown = (e) => {
        if (e.key === 'Enter' && !isComposing) {
            saveAndClose();
        } else if (e.key === 'Escape') {
            isSaved = true; // Don't save on blur if escaped
            input.remove();
            state.editingId = null; // Clear editing state
            renderAnnotations();
        }
    };

    input.onblur = () => {
        // Delay slightly to allow click events to process (e.g. if clicking another annotation)
        setTimeout(() => {
            saveAndClose();
        }, 100);
    };

    annotationLayer.appendChild(input);
    input.focus();
}

function deleteAnnotation(id) {
    pushState();
    state.annotations = state.annotations.filter(a => a.id !== id);
    state.selectedAnnotationId = null;
    renderAnnotations();
    renderSidebar();
}

function selectAnnotation(id) {
    state.selectedAnnotationId = id;

    // Update UI based on selection type
    const ann = state.annotations.find(a => a.id === id);
    if (ann) {
        if (ann.type === 'text') {
            pointSizeControl.style.display = 'none';
            textSizeControl.style.display = 'inline';
            fontSelect.style.display = 'inline-block';
            fontSizeInput.value = ann.size || 14;
            fontSelect.value = ann.font || 'Noto Sans JP';

            // Also switch mode UI to show style controls if not already
            styleControls.style.display = 'flex';
            stampControls.style.display = 'none';
        } else if (ann.type === 'point') {
            pointSizeControl.style.display = 'inline';
            textSizeControl.style.display = 'none';
            fontSelect.style.display = 'none';
            // Update radio button
            const radio = document.querySelector(`input[name="size"][value="${ann.size}"]`);
            if (radio) radio.checked = true;

            styleControls.style.display = 'flex';
            stampControls.style.display = 'none';
        } else if (ann.type === 'stamp') {
            styleControls.style.display = 'none';
            stampControls.style.display = 'flex';
            stampTypeSelect.value = ann.stampType;
        } else if (ann.type === 'image') {
            styleControls.style.display = 'none';
            stampControls.style.display = 'none';
            imageControls.style.display = 'flex';
        } else {
            // Highlight/Whiteout
            styleControls.style.display = 'none';
            stampControls.style.display = 'none';
        }
    }

    renderAnnotations();
    // Highlight in sidebar
    renderSidebar();
}

function renderAnnotations() {
    try {
        annotationLayer.innerHTML = '';
        const pageAnnotations = state.annotations.filter(a => a.pageNumber === state.pageNum);

        // Filter out the one being edited so it doesn't show up twice (once as text, once as input)
        const visibleAnnotations = state.editingId
            ? pageAnnotations.filter(a => a.id !== state.editingId)
            : pageAnnotations;

        // Draw Graph Lines
        if (state.selectedDotIds.size > 1) {
            const svgNS = "http://www.w3.org/2000/svg";
            const svg = document.createElementNS(svgNS, "svg");
            svg.style.position = "absolute";
            svg.style.top = "0";
            svg.style.left = "0";
            svg.style.width = "100%";
            svg.style.height = "100%";
            svg.style.pointerEvents = "none";
            svg.style.zIndex = "5"; // Below dots (z-index 10)
            svg.setAttribute("viewBox", `0 0 ${canvas.width} ${canvas.height}`);

            const selectedPoints = pageAnnotations.filter(a => state.selectedDotIds.has(a.id) && a.type === 'point');
            const bluePoints = selectedPoints.filter(a => a.color === 'blue').sort((a, b) => a.xPct - b.xPct);
            const redPoints = selectedPoints.filter(a => a.color === 'red').sort((a, b) => a.xPct - b.xPct);
            const blackPoints = selectedPoints.filter(a => a.color === 'black').sort((a, b) => a.xPct - b.xPct);

            [bluePoints, redPoints, blackPoints].forEach(points => {
                if (points.length > 1) {
                    const polyline = document.createElementNS(svgNS, "polyline");
                    const pointsStr = points.map(p => `${p.xPct * canvas.width},${p.yPct * canvas.height}`).join(' ');
                    polyline.setAttribute("points", pointsStr);
                    polyline.setAttribute("fill", "none");
                    let strokeColor = "var(--blue-color)";
                    if (points[0].color === 'red') strokeColor = "var(--red-color)";
                    else if (points[0].color === 'black') strokeColor = "#333";

                    polyline.setAttribute("stroke", strokeColor);
                    polyline.setAttribute("stroke-width", "2");
                    polyline.style.pointerEvents = "none"; // Ensure lines don't block clicks
                    svg.appendChild(polyline);
                }
            });

            annotationLayer.appendChild(svg);
        }

        visibleAnnotations.forEach(ann => {
            const el = document.createElement('div');
            el.classList.add('annotation');
            el.dataset.id = ann.id;

            // Use percentages for positioning
            el.style.left = `${ann.xPct * 100}%`;
            el.style.top = `${ann.yPct * 100}%`;

            if (state.selectedAnnotationId === ann.id) {
                el.classList.add('selected');
            }

            if (ann.type === 'point') {
                el.classList.add('dot', ann.color);
                // Use numeric size
                const size = ann.size || 14;
                el.style.width = `${size}px`;
                el.style.height = `${size}px`;

                if (state.isGraphMode && state.selectedDotIds.has(ann.id)) {
                    el.style.boxShadow = '0 0 0 3px rgba(0,0,0,0.3)';
                }
            } else if (ann.type === 'text') {
                el.classList.add('text-annotation', ann.color);
                // Use numeric size
                const size = ann.size || 14;
                el.style.fontSize = `${size}px`;
                el.style.fontFamily = ann.font || 'Noto Sans JP';

                el.textContent = ann.text;
                el.ondblclick = (e) => {
                    e.stopPropagation();
                    // For editing, we need pixel coordinates for the input
                    const rect = annotationLayer.getBoundingClientRect();
                    const x = ann.xPct * rect.width;
                    const y = ann.yPct * rect.height;
                    showTextInput(x, y, ann.id);
                };
            } else if (ann.type === 'highlight') {
                el.classList.add('highlight-annotation');
                el.style.width = `${ann.wPct * 100}%`;
                el.style.height = `${ann.hPct * 100}%`;
                el.style.transform = 'none';
            } else if (ann.type === 'whiteout') {
                el.classList.add('whiteout-annotation');
                el.style.width = `${ann.wPct * 100}%`;
                el.style.height = `${ann.hPct * 100}%`;
                el.style.transform = 'none';
            } else if (ann.type === 'stamp') {
                el.classList.add('stamp-annotation', ann.stampType);
                let text = ann.stampType.toUpperCase();
                if (ann.stampType === 'review') text = '要確認';
                if (ann.stampType === 'fix') text = '修正';
                el.textContent = text;
            } else if (ann.type === 'image') {
                el.classList.add('image-annotation');
                el.style.width = `${ann.wPct * 100}%`;
                el.style.height = `${ann.hPct * 100}%`;
                el.style.backgroundImage = `url(${ann.imageData})`;
                el.style.backgroundSize = 'contain';
                el.style.backgroundRepeat = 'no-repeat';
                el.style.backgroundPosition = 'center';

                // Add resize handles if selected
                if (state.selectedAnnotationId === ann.id) {
                    ['nw', 'ne', 'sw', 'se'].forEach(pos => {
                        const handle = document.createElement('div');
                        handle.className = `resize-handle ${pos}`;
                        handle.dataset.handle = pos;
                        handle.style.position = 'absolute';
                        handle.style.width = '10px';
                        handle.style.height = '10px';
                        handle.style.backgroundColor = 'var(--primary-color)';
                        handle.style.border = '1px solid white';
                        if (pos.includes('n')) handle.style.top = '-5px';
                        if (pos.includes('s')) handle.style.bottom = '-5px';
                        if (pos.includes('w')) handle.style.left = '-5px';
                        if (pos.includes('e')) handle.style.right = '-5px';
                        handle.style.cursor = `${pos}-resize`;
                        handle.style.zIndex = '25';
                        el.appendChild(handle);
                    });
                }
            }

            annotationLayer.appendChild(el);
        });
    } catch (e) {
        console.error('Fatal error in renderAnnotations:', e);
    }
}

function renderSidebar() {
    commentList.innerHTML = '';
    if (state.annotations.length === 0) {
        commentList.innerHTML = '<li class="sidebar-item-empty">注釈はまだありません</li>';
        return;
    }

    state.annotations.sort((a, b) => a.pageNumber - b.pageNumber || a.yPct - b.yPct).forEach(ann => {
        const li = document.createElement('li');
        li.className = 'sidebar-item';
        if (state.selectedAnnotationId === ann.id) li.classList.add('selected');

        let content = '';
        let icon = '';
        let details = '';

        if (ann.type === 'point') {
            content = '点';
            icon = '●';
            // Convert size to label
            let sizeLabel = ann.size + 'px';
            if (ann.size === 8) sizeLabel = '小';
            if (ann.size === 12) sizeLabel = '中';
            if (ann.size === 16) sizeLabel = '大';
            details = `サイズ: ${sizeLabel}`;
        } else if (ann.type === 'text') {
            content = ann.text || '(空のテキスト)';
            icon = 'T';
            let sizeLabel = ann.size + 'px';
            if (ann.size === 8) sizeLabel = '小';
            if (ann.size === 12) sizeLabel = '中';
            if (ann.size === 16) sizeLabel = '大';
            details = `サイズ: ${sizeLabel}`;
        } else if (ann.type === 'highlight') {
            content = ann.comment || 'ハイライト';
            icon = 'H';
        } else if (ann.type === 'whiteout') {
            content = ann.comment || 'ホワイトアウト';
            icon = 'W';
        } else if (ann.type === 'stamp') {
            let text = ann.stampType.toUpperCase();
            if (ann.stampType === 'review') text = '要確認';
            if (ann.stampType === 'fix') text = '修正';
            content = text;
            icon = 'S';
        } else if (ann.type === 'image') {
            content = '画像';
            icon = 'I';
        }

        li.innerHTML = `
            <div class="sidebar-item-header">
                <span>P.${ann.pageNumber} [${icon}]</span>
                <button class="sidebar-delete-btn" onclick="event.stopPropagation(); deleteAnnotation('${ann.id}')">×</button>
            </div>
            <div class="sidebar-item-content">${content}</div>
        `;

        li.onclick = () => {
            if (state.pageNum !== ann.pageNumber) {
                state.pageNum = ann.pageNumber;
                renderPage(state.pageNum);
                // Wait for render then highlight
                setTimeout(() => highlightAnnotation(ann.id), 100);
            } else {
                highlightAnnotation(ann.id);
            }
            selectAnnotation(ann.id);
        };

        commentList.appendChild(li);
    });
}

function highlightAnnotation(id) {
    const el = document.querySelector(`.annotation[data-id="${id}"]`);
    if (el) {
        el.classList.remove('highlighted');
        void el.offsetWidth; // trigger reflow
        el.classList.add('highlighted');
    }
}

// --- Undo / Redo ---

function pushState() {
    const snapshot = JSON.stringify(state.annotations);
    state.undoStack.push(snapshot);
    state.redoStack = []; // Clear redo stack on new action
    updateUndoRedoButtons();
}

function undo() {
    if (state.undoStack.length === 0) return;
    const current = JSON.stringify(state.annotations);
    state.redoStack.push(current);
    const prev = state.undoStack.pop();
    state.annotations = JSON.parse(prev);
    renderAnnotations();
    renderSidebar();
    updateUndoRedoButtons();
}

function redo() {
    if (state.redoStack.length === 0) return;
    const current = JSON.stringify(state.annotations);
    state.undoStack.push(current);
    const next = state.redoStack.pop();
    state.annotations = JSON.parse(next);
    renderAnnotations();
    renderSidebar();
    updateUndoRedoButtons();
}

function updateUndoRedoButtons() {
    undoBtn.disabled = state.undoStack.length === 0;
    redoBtn.disabled = state.redoStack.length === 0;
}

// --- Export ---

function exportCSV() {
    const headers = ['id', 'type', 'pageNumber', 'xPct', 'yPct', 'widthPct', 'heightPct', 'color', 'size', 'content', 'stampType', 'font'];
    const rows = state.annotations.map(a => [
        a.id,
        a.type,
        a.pageNumber,
        a.xPct.toFixed(4),
        a.yPct.toFixed(4),
        a.wPct ? a.wPct.toFixed(4) : '',
        a.hPct ? a.hPct.toFixed(4) : '',
        a.color || '',
        a.size || '',
        (a.text || a.comment || '').replace(/"/g, '""'), // Escape quotes
        a.stampType || '',
        a.font || ''
    ]);

    const csvContent = [
        headers.join(','),
        ...rows.map(row => row.map(cell => `"${cell}"`).join(','))
    ].join('\n');

    const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvContent], { type: 'text/csv;charset=utf-8;' }); // BOM for Excel
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `annotations_${Date.now()}.csv`;
    link.click();
}

// --- PDF Download (Merge) ---

// --- PDF Download (Print) ---



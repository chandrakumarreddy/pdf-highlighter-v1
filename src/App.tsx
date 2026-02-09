import { useState, useEffect, useRef, useCallback, useMemo, memo } from 'react';
import {
  Upload,
  Undo,
  Download,
  FileText,
  Table,
  Sparkles,
  AlertCircle,
  Loader2,
  Trash2,
  FileSearch,
  PlusCircle,
  FileSignature,
} from 'lucide-react';

// External Scripts
const PDF_JS_URL =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
const PDF_WORKER_URL =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
const XLSX_URL =
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const JSZIP_URL =
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
const DOCX_JS_URL =
  'https://cdn.jsdelivr.net/npm/docx-preview@0.1.15/dist/docx-preview.min.js';
const TFJS_URL =
  'https://cdn.jsdelivr.net/npm/@tensorflow/tfjs@4.15.0/dist/tf.min.js';

const SIMILARITY_THRESHOLD = 0.985;
const X_MATCH_TOLERANCE = 15;
const LINE_Y_TOLERANCE = 3; // Pixels to group nodes into the same visual line

// --- VIRTUALIZED PDF PAGE COMPONENT ---

const VirtualPage = memo(
  ({ pageNum, pdfDoc, selections, excludedNodes, onSelect, onExclude }) => {
    const containerRef = useRef(null);
    const [isLoaded, setIsLoaded] = useState(false);
    const [isVisible, setIsVisible] = useState(false);
    const renderTaskRef = useRef(null);

    useEffect(() => {
      const observer = new IntersectionObserver(
        ([entry]) => setIsVisible(entry.isIntersecting),
        { rootMargin: '600px 0px', threshold: 0.01 },
      );
      if (containerRef.current) observer.observe(containerRef.current);
      return () => observer.disconnect();
    }, []);

    const renderPage = useCallback(async () => {
      if (!pdfDoc || !isVisible || isLoaded) return;

      try {
        const page = await pdfDoc.getPage(pageNum);
        const scale = window.devicePixelRatio || 1;
        const viewport = page.getViewport({ scale: 1.5 * scale });
        const container = containerRef.current;
        if (!container) return;

        container.style.width = `${viewport.width / scale}px`;
        container.style.height = `${viewport.height}px`;

        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.width = '100%';
        canvas.style.height = '100%';
        container.appendChild(canvas);

        const textLayer = document.createElement('div');
        textLayer.className =
          'absolute inset-0 z-10 pointer-events-none text-layer-container';
        container.appendChild(textLayer);

        const ctx = canvas.getContext('2d');
        renderTaskRef.current = page.render({ canvasContext: ctx, viewport });
        await renderTaskRef.current.promise;

        const textContent = await page.getTextContent();
        const fragment = document.createDocumentFragment();

        textContent.items.forEach((item, idx) => {
          const tx = window.pdfjsLib.Util.transform(
            viewport.transform,
            item.transform,
          );
          // Normalize Y coordinate for grouping
          const pdfY =
            Math.round(item.transform[5] / LINE_Y_TOLERANCE) * LINE_Y_TOLERANCE;

          const span = document.createElement('span');
          span.className =
            'text-node absolute pointer-events-auto text-transparent cursor-pointer hover:bg-indigo-500/5 transition-colors group/node';
          span.textContent = item.str;
          span.style.left = `${tx[4] / scale}px`;
          span.style.top = `${(tx[5] - item.height * 1.5 * scale) / scale}px`;
          span.style.width = `${(item.width * 1.5 * scale) / scale}px`;
          span.style.height = `${(item.height * 1.5 * scale) / scale}px`;
          span.style.fontSize = `${(item.height * 1.5 * scale) / scale}px`;

          const nodeId = `pdf-${pageNum}-${idx}`;
          span.setAttribute('data-node-id', nodeId);
          span.setAttribute('data-pdf-y', pdfY.toString());

          const removeBtn = document.createElement('div');
          removeBtn.className =
            'absolute -top-2 -right-2 hidden group-hover/node:flex bg-rose-500 text-white rounded-full p-0.5 z-50 pointer-events-auto shadow-sm transition-transform hover:scale-110 flex items-center justify-center';
          removeBtn.innerHTML =
            '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>';
          removeBtn.onclick = (e) => {
            e.stopPropagation();
            onExclude(nodeId);
          };
          span.appendChild(removeBtn);

          fragment.appendChild(span);
        });

        textLayer.appendChild(fragment);
        setIsLoaded(true);
        page.cleanup();
      } catch (err) {
        console.error('PDF Render Error:', err);
      }
    }, [pdfDoc, pageNum, isVisible, isLoaded, onExclude]);

    useEffect(() => {
      const container = containerRef.current;
      if (isVisible) renderPage();
      return () => {
        if (container) {
          if (renderTaskRef.current) renderTaskRef.current.cancel();
          container.innerHTML = '';
          setIsLoaded(false);
        }
      };
    }, [isVisible, renderPage]);

    // Apply Line-Aware Highlights
    useEffect(() => {
      if (!isLoaded || !containerRef.current) return;
      const nodes = containerRef.current.querySelectorAll('.text-node');
      nodes.forEach((n) => {
        const nodeId = n.getAttribute('data-node-id');
        const isExcluded = excludedNodes.includes(nodeId);
        const isSelected = selections.nodeIds.has(nodeId);

        if (isSelected && !isExcluded) {
          n.classList.add(
            'bg-yellow-400/40',
            'outline',
            'outline-1',
            'outline-amber-600',
          );
        } else {
          n.classList.remove(
            'bg-yellow-400/40',
            'outline',
            'outline-1',
            'outline-amber-600',
          );
        }
      });
    }, [selections, isLoaded, excludedNodes]);

    return (
      <div
        ref={containerRef}
        className="relative mx-auto mb-8 min-h-[500px] overflow-hidden rounded-sm border border-slate-300 bg-white shadow-md"
        onClick={(e) => {
          const target = e.target.closest('.text-node');
          if (target) onSelect(target.getAttribute('data-node-id'));
        }}
      />
    );
  },
);

// --- DOCX VIEWER COMPONENT ---

const DocxViewer = memo(
  ({ buffer, selections, excludedNodes, onSelect, onExclude }) => {
    const containerRef = useRef(null);
    const [error, setError] = useState(null);
    const [isRendered, setIsRendered] = useState(false);
    const isMountedRef = useRef(false);

    useEffect(() => {
      isMountedRef.current = true;
      return () => {
        isMountedRef.current = false;
      };
    }, []);

    // eslint-disable-next-line @eslint-react/hooks-extra/no-direct-set-state-in-use-effect
    useEffect(() => {
      if (containerRef.current && buffer) {
        if (!window.docx) {
          if (isMountedRef.current) setError('Word viewer engine not ready');
          return;
        }
        if (isMountedRef.current) setIsRendered(false);
        containerRef.current.innerHTML = '';
        window.docx
          .renderAsync(buffer, containerRef.current)
          .then(() => {
            const containerRect = containerRef.current.getBoundingClientRect();
            const blocks = containerRef.current.querySelectorAll(
              'p, td, h1, h2, h3, h4, h5, h6',
            );
            blocks.forEach((el, idx) => {
              if (el.innerText && el.innerText.trim().length > 0) {
                const rect = el.getBoundingClientRect();
                const relX = Math.round(rect.left - containerRect.left);
                el.setAttribute('data-docx-x', relX.toString());
                el.setAttribute('data-node-id', `docx-node-${idx}`);
                el.classList.add(
                  'docx-selectable',
                  'cursor-pointer',
                  'transition-all',
                  'rounded-sm',
                );
                el.style.position = 'relative';
                el.querySelectorAll('span').forEach(
                  (s) => (s.style.pointerEvents = 'none'),
                );
              }
            });
            if (isMountedRef.current) setIsRendered(true);
          })
          .catch((err) => {
            console.error('Docx render error:', err);
            if (isMountedRef.current) setError('Failed to render Word document');
          });
      }
    }, [buffer]);

    useEffect(() => {
      if (!isRendered || !containerRef.current) return;
      const elements =
        containerRef.current.querySelectorAll('.docx-selectable');
      elements.forEach((el) => {
        const nodeId = el.getAttribute('data-node-id');
        const isExcluded = excludedNodes.includes(nodeId);
        const isSelected = selections.nodeIds.has(nodeId);

        const existingBtn = el.querySelector('.docx-remove-btn');
        if (existingBtn) existingBtn.remove();

        if (isSelected && !isExcluded) {
          el.style.backgroundColor = 'rgba(234, 179, 8, 0.25)';
          el.style.boxShadow = 'inset 0 0 0 1px #ca8a04';
          const btn = document.createElement('div');
          btn.className =
            'docx-remove-btn absolute -top-1 -right-1 hidden bg-rose-500 text-white rounded-full p-0.5 z-50 pointer-events-auto shadow-sm animate-in zoom-in duration-200';
          btn.style.width = '16px';
          btn.style.height = '16px';
          btn.innerHTML =
            '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="4"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>';
          btn.onclick = (e) => {
            e.stopPropagation();
            onExclude(nodeId);
          };
          el.appendChild(btn);
        } else {
          el.style.backgroundColor = '';
          el.style.boxShadow = '';
        }
      });
    }, [isRendered, selections, excludedNodes, onExclude]);

    const handleDocxClick = (e) => {
      const target = e.target.closest('.docx-selectable');
      if (target) onSelect(target.getAttribute('data-node-id'));
    };

    if (error)
      return (
        <div className="flex items-center justify-center p-20 font-bold text-rose-500">
          <AlertCircle className="mr-2" /> {error}
        </div>
      );

    return (
      <div>
        <style>{`
        .docx-container .docx-selectable:hover { background-color: rgba(99, 102, 241, 0.05); }
        .docx-container .docx-selectable:hover .docx-remove-btn { display: flex !important; align-items: center; justify-content: center; }
        .docx-container section { background: white !important; margin: 0 auto 2.5rem auto !important; box-shadow: 0 10px 30px -10px rgba(0,0,0,0.15) !important; padding: 4rem !important; max-width: 850px !important; }
      `}</style>
        <div
          ref={containerRef}
          onClick={handleDocxClick}
          className="docx-container mx-auto"
        />
      </div>
    );
  },
);

// --- MAIN APPLICATION ---

const App = () => {
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [activeFileId, setActiveFileId] = useState(null);
  const [fileType, setFileType] = useState(null);
  const [pdfDoc, setPdfDoc] = useState(null);
  const [numPages, setNumPages] = useState(0);
  const [excelSheets, setExcelSheets] = useState([]);
  const [docxBuffer, setDocxBuffer] = useState(null);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [loading, setLoading] = useState(false);
  const [selections, setSelections] = useState(() => ({
    type: null,
    nodeIds: new Set(),
    indices: {},
  }));
  const [excludedNodes, setExcludedNodes] = useState([]);

  const structuralTensorsRef = useRef(null);
  const nodeLookupRef = useRef([]);
  const fileInputRef = useRef(null);

  useEffect(() => {
    const loadScript = (src, id) =>
      new Promise((resolve) => {
        if (document.getElementById(id)) return resolve();
        const script = document.createElement('script');
        script.src = src;
        script.id = id;
        script.async = false;
        script.onload = resolve;
        document.head.appendChild(script);
      });

    const initScripts = async () => {
      await loadScript(PDF_JS_URL, 'pdf-js');
      await loadScript(XLSX_URL, 'xlsx-js');
      await loadScript(JSZIP_URL, 'jszip-js');
      await loadScript(DOCX_JS_URL, 'docx-js');
      await loadScript(TFJS_URL, 'tfjs');
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
    };

    initScripts();
  }, []);

  // --- TFJS STRUCTURAL PATTERN ENGINE ---
  const buildStructuralMap = async (pdfInstance) => {
    if (!window.tf || !pdfInstance) return;
    const nodes = [];

    for (let p = 1; p <= pdfInstance.numPages; p++) {
      const page = await pdfInstance.getPage(p);
      const textContent = await page.getTextContent();
      textContent.items.forEach((item, idx) => {
        if (item.str.trim().length === 0) return;
        nodes.push({
          id: `pdf-${p}-${idx}`,
          x: item.transform[4],
          y:
            Math.round(item.transform[5] / LINE_Y_TOLERANCE) * LINE_Y_TOLERANCE,
          fontSize: item.height,
          text: item.str,
          page: p,
        });
      });
    }

    if (nodes.length > 0) {
      nodeLookupRef.current = nodes;
      // Feature Encoding
      const features = nodes.map((n) => [
        n.x / 1000,
        n.fontSize / 100,
        /^\d/.test(n.text.trim()) ? 1 : 0,
        Math.min(n.text.length / 500, 1),
      ]);
      if (structuralTensorsRef.current) structuralTensorsRef.current.dispose();
      structuralTensorsRef.current = window.tf.tensor2d(features);
    }
  };

  const switchToFile = async (fileEntry) => {
    setLoading(true);
    setActiveFileId(fileEntry.id);
    setFileType(fileEntry.type);
    setSelections({ type: null, nodeIds: new Set(), indices: {} });
    setExcludedNodes([]);
    setPdfDoc(null);
    setExcelSheets([]);
    setDocxBuffer(null);

    try {
      // SAFE BUFFER HANDLING: slice(0) avoids detachment
      const safeBuffer = fileEntry.buffer.slice(0);

      if (fileEntry.type === 'PDF') {
        const pdf = await window.pdfjsLib.getDocument({
          data: new Uint8Array(safeBuffer),
        }).promise;
        setPdfDoc(pdf);
        setNumPages(pdf.numPages);
        await buildStructuralMap(pdf);
      } else if (fileEntry.type === 'XLSX') {
        const workbook = window.XLSX.read(safeBuffer, { type: 'array' });
        setExcelSheets(
          workbook.SheetNames.map((name) => ({
            name,
            data: window.XLSX.utils.sheet_to_json(workbook.Sheets[name], {
              header: 1,
              defval: '',
            }),
          })),
        );
      } else if (fileEntry.type === 'DOCX') {
        setDocxBuffer(safeBuffer);
      }
    } catch (err) {
      console.error(err);
    }
    setLoading(false);
  };

  const performPatternMatch = (seedNodeId) => {
    if (fileType === 'PDF') {
      const seedIdx = nodeLookupRef.current.findIndex(
        (n) => n.id === seedNodeId,
      );
      if (seedIdx === -1) return;

      window.tf.tidy(() => {
        const seedFeature = structuralTensorsRef.current.slice(
          [seedIdx, 0],
          [1, -1],
        );
        const dotProduct = structuralTensorsRef.current.matMul(
          seedFeature.transpose(),
        );
        const similarity = dotProduct.dataSync();

        const newIds = new Set(selections.nodeIds);
        const activeLines = new Set();

        similarity.forEach((score, idx) => {
          if (score > SIMILARITY_THRESHOLD) {
            const node = nodeLookupRef.current[idx];
            activeLines.add(`${node.page}-${node.y}`);
          }
        });

        // FULL LINE SELECTION: Match everything sharing the same Y-coordinate as a similar node
        nodeLookupRef.current.forEach((n) => {
          if (activeLines.has(`${n.page}-${n.y}`)) newIds.add(n.id);
        });

        setSelections((prev) => ({ ...prev, type: 'PDF', nodeIds: newIds }));
      });
    }

    if (fileType === 'DOCX') {
      const container = document.querySelector('.docx-container');
      const targetEl = container?.querySelector(
        `[data-node-id="${seedNodeId}"]`,
      );
      if (!targetEl) return;

      const targetX = parseInt(targetEl.getAttribute('data-docx-x'));
      const allSelectable = container.querySelectorAll('.docx-selectable');
      const newIds = new Set(selections.nodeIds);

      allSelectable.forEach((el) => {
        const x = parseInt(el.getAttribute('data-docx-x'));
        if (Math.abs(x - targetX) < X_MATCH_TOLERANCE) {
          newIds.add(el.getAttribute('data-node-id'));
        }
      });
      setSelections((prev) => ({ ...prev, type: 'DOCX', nodeIds: newIds }));
    }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const ext = file.name.split('.').pop().toLowerCase();
    const type = ext === 'xlsx' ? 'XLSX' : ext.includes('doc') ? 'DOCX' : 'PDF';
    const buffer = await file.arrayBuffer();
    const id = Math.random().toString(36).substr(2, 9);
    const newFile = { id, name: file.name, type, buffer };
    setUploadedFiles((prev) => [...prev, newFile]);
    switchToFile(newFile);
    e.target.value = null;
  };

  const activeFileNameStr = useMemo(() => {
    const found = uploadedFiles.find((f) => f.id === activeFileId);
    return found ? String(found.name) : 'No file selected';
  }, [uploadedFiles, activeFileId]);

  const totalSelectedCount = useMemo(() => {
    if (!selections.type) return 0;
    if (selections.type === 'XLSX') {
      return Object.values(selections.indices || {}).reduce(
        (acc, curr) => acc + (curr?.length || 0),
        0,
      );
    }
    return selections.nodeIds.size;
  }, [selections]);

  return (
    <div className="flex h-screen overflow-hidden bg-[#F0F2F5] font-sans text-slate-800">
      <aside className="z-50 flex w-[300px] flex-col border-r border-slate-200 bg-white shadow-xl">
        <div className="p-6">
          <div className="mb-2 flex items-center gap-4">
            <div className="relative overflow-hidden rounded-xl bg-slate-900 p-2.5 shadow-xl">
              <FileSearch className="relative z-10 text-white" size={24} />
              <div className="absolute top-0 right-0 h-full w-full bg-gradient-to-br from-indigo-500/40 to-transparent"></div>
            </div>
            <h1 className="text-xl font-extrabold tracking-tight text-slate-900">
              RFP Agent
            </h1>
          </div>
          <div className="mt-4 h-px w-full bg-slate-100"></div>
        </div>

        <div className="scrollbar-hide flex-1 space-y-6 overflow-y-auto px-6 pb-6">
          <div
            onClick={() => fileInputRef.current?.click()}
            className="group flex cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed border-slate-200 p-6 transition-all hover:border-indigo-400 hover:bg-indigo-50/30"
          >
            <Upload
              size={24}
              className="mb-2 text-slate-400 transition-transform group-hover:scale-110"
            />
            <p className="text-xs font-bold text-slate-700">
              Upload Project File
            </p>
            <input
              type="file"
              ref={fileInputRef}
              className="hidden"
              accept=".pdf,.xlsx,.docx,.doc"
              onChange={handleFileUpload}
            />
          </div>

          <div className="space-y-2">
            <h3 className="mb-4 text-[10px] font-bold tracking-widest text-slate-400 uppercase">
              Project Library
            </h3>
            {uploadedFiles.map((f) => (
              <div
                key={f.id}
                onClick={() => switchToFile(f)}
                className={`group flex cursor-pointer items-center gap-3 rounded-xl border p-3 transition-all ${activeFileId === f.id ? 'border-indigo-100 bg-indigo-50 shadow-sm ring-1 ring-indigo-200' : 'border-transparent bg-white hover:bg-slate-50'}`}
              >
                <div
                  className={`rounded-lg p-1.5 ${activeFileId === f.id ? 'bg-indigo-600 text-white shadow-sm' : 'bg-slate-100 text-slate-500'}`}
                >
                  {f.type === 'PDF' && <FileText size={14} />}
                  {f.type === 'XLSX' && <Table size={14} />}
                  {f.type === 'DOCX' && <FileSignature size={14} />}
                </div>
                <div className="min-w-0 flex-1">
                  <p
                    className={`truncate text-xs font-bold ${activeFileId === f.id ? 'text-indigo-900' : 'text-slate-700'}`}
                  >
                    {String(f.name)}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setUploadedFiles((prev) =>
                      prev.filter((x) => x.id !== f.id),
                    );
                    if (activeFileId === f.id) setActiveFileId(null);
                  }}
                  className="p-1 text-slate-400 opacity-0 transition-all group-hover:opacity-100 hover:text-rose-500"
                >
                  <Trash2 size={14} />
                </button>
              </div>
            ))}
          </div>

          <div className="space-y-3 border-t border-slate-100 pt-4">
            <button
              type="button"
              onClick={() => {
                setSelections({ type: null, nodeIds: new Set(), indices: {} });
                setExcludedNodes([]);
              }}
              className="flex w-full items-center justify-center gap-2 rounded-xl border border-slate-200 py-3 text-sm font-bold text-slate-500 transition-colors hover:bg-slate-50"
            >
              <Undo size={16} /> Clear Patterns
            </button>
            <button
              type="button"
              className="flex w-full items-center justify-center gap-2 rounded-xl bg-indigo-600 py-4 text-sm font-bold text-white shadow-lg shadow-indigo-200 transition-all hover:bg-indigo-700 active:scale-95"
            >
              <Download size={18} /> Export Clean Data
            </button>
          </div>
        </div>
      </aside>

      <main className="relative flex flex-1 flex-col overflow-hidden">
        <header className="sticky top-0 z-40 flex h-14 items-center border-b border-slate-200 bg-white px-6 shadow-sm">
          <div className="flex items-center gap-3 overflow-hidden text-slate-500">
            {fileType === 'PDF' && (
              <FileText size={18} className="text-indigo-500" />
            )}
            {fileType === 'XLSX' && (
              <Table size={18} className="text-emerald-500" />
            )}
            {fileType === 'DOCX' && (
              <FileSignature size={18} className="text-blue-500" />
            )}
            <span className="max-w-[400px] truncate text-sm font-bold">
              {activeFileNameStr}
            </span>
          </div>
          {totalSelectedCount > 0 && (
            <div className="ml-auto flex items-center gap-2 rounded-full border border-amber-200 bg-amber-50 px-3 py-1 text-[10px] font-bold text-amber-700 uppercase">
              <Sparkles size={12} /> AI Analysis Active
            </div>
          )}
        </header>

        <div className="relative flex-1 overflow-auto scroll-smooth bg-[#E8EBF2]">
          {!activeFileId && !loading && (
            <div className="flex h-full flex-col items-center justify-center text-slate-400">
              <FileSearch size={64} className="mb-4 opacity-10" />
              <p className="text-center text-lg font-bold">
                Structural Node Highlighter
                <br />
                <span className="text-sm font-normal">
                  Select a line once, capture the entire hierarchy.
                </span>
              </p>
            </div>
          )}
          {loading && (
            <div className="flex h-full flex-col items-center justify-center">
              <Loader2
                className="mb-4 animate-spin text-indigo-600"
                size={40}
              />
              <p className="text-sm font-bold tracking-widest text-slate-500 uppercase">
                Running TF.js Inference...
              </p>
            </div>
          )}
          {!loading && fileType === 'PDF' && pdfDoc && (
            <div className="p-8">
              {Array.from({ length: numPages }).map((_, i) => (
                <VirtualPage
                  key={`page-${i + 1}`}
                  pageNum={i + 1}
                  pdfDoc={pdfDoc}
                  selections={selections}
                  excludedNodes={excludedNodes}
                  onSelect={(id) => performPatternMatch(id)}
                  onExclude={(nodeId) =>
                    setExcludedNodes((prev) => [...prev, nodeId])
                  }
                />
              ))}
            </div>
          )}

          {!loading && fileType === 'XLSX' && (
            <div className="flex h-full flex-col overflow-hidden bg-white">
              <div className="relative flex-1 overflow-auto">
                <table className="w-max min-w-full table-fixed border-collapse">
                  <thead className="sticky top-0 z-30">
                    <tr className="bg-[#f8fafc]">
                      <th className="sticky left-0 z-50 h-7 w-10 border border-slate-300 bg-[#f8fafc]"></th>
                      {excelSheets[activeSheetIndex]?.data[0] &&
                        Array.from({
                          length: Object.keys(
                            excelSheets[activeSheetIndex].data[0],
                          ).length,
                        }).map((_, i) => {
                          const isSelected =
                            selections.type === 'XLSX' &&
                            selections.indices?.[activeSheetIndex]?.includes(i);
                          return (
                            <th
                              key={`col-${i}`}
                              onClick={() => {
                                setSelections((prev) => {
                                  const cur =
                                    prev.indices?.[activeSheetIndex] || [];
                                  const next = cur.includes(i)
                                    ? cur.filter((x) => x !== i)
                                    : [...cur, i];
                                  return {
                                    ...prev,
                                    type: 'XLSX',
                                    indices: {
                                      ...prev.indices,
                                      [activeSheetIndex]: next,
                                    },
                                  };
                                });
                              }}
                              className={`group h-7 min-w-[180px] cursor-pointer border border-slate-300 transition-colors ${isSelected ? 'bg-indigo-600' : 'bg-[#f8fafc] hover:bg-slate-200'}`}
                            >
                              <PlusCircle
                                size={14}
                                className={`mx-auto ${isSelected ? 'text-white' : 'text-slate-300 group-hover:text-indigo-600'}`}
                              />
                            </th>
                          );
                        })}
                    </tr>
                  </thead>
                  <tbody>
                    {(excelSheets[activeSheetIndex]?.data || []).map(
                      (row, rIdx) => (
                        <tr key={`row-${rIdx}`} className="group/row h-8">
                          <td className="sticky left-0 z-20 flex h-8 w-10 items-center justify-center border border-slate-300 bg-[#f8fafc] text-center text-[10px]">
                            <span>{rIdx + 1}</span>
                          </td>
                          {Object.values(row).map((cell, cIdx) => {
                            const isSelectedCol =
                              selections.type === 'XLSX' &&
                              selections.indices?.[activeSheetIndex]?.includes(
                                cIdx,
                              );
                            const cellStr = cell ? String(cell).trim() : '';
                            return (
                              <td
                                key={`row-${rIdx}-col-${cIdx}`}
                                className={`truncate border border-slate-200 px-3 py-1 text-[12px] transition-all ${isSelectedCol && cellStr ? 'bg-amber-50/50 font-medium ring-1 ring-amber-400 ring-inset' : 'bg-white'}`}
                              >
                                {cellStr}
                              </td>
                            );
                          })}
                        </tr>
                      ),
                    )}
                  </tbody>
                </table>
              </div>
              <div className="flex h-10 shrink-0 items-center gap-2 overflow-x-auto border-t border-slate-300 bg-[#f1f3f4] px-4">
                {excelSheets.map((s, idx) => (
                  <div
                    key={`sheet-${s.name}`}
                    onClick={() => setActiveSheetIndex(idx)}
                    className={`flex h-full cursor-pointer items-center border-x border-slate-200 px-5 text-xs font-bold whitespace-nowrap transition-all ${activeSheetIndex === idx ? 'border-t-2 border-t-indigo-600 bg-white text-indigo-600 shadow-sm' : 'text-slate-500 hover:bg-slate-200'}`}
                  >
                    {String(s.name)}
                  </div>
                ))}
              </div>
            </div>
          )}

          {!loading && fileType === 'DOCX' && docxBuffer && (
            <DocxViewer
              buffer={docxBuffer}
              selections={selections}
              excludedNodes={excludedNodes}
              onSelect={(id) => performPatternMatch(id)}
              onExclude={(nodeId) =>
                setExcludedNodes((prev) => [...prev, nodeId])
              }
            />
          )}
        </div>
      </main>
    </div>
  );
};

export default App;

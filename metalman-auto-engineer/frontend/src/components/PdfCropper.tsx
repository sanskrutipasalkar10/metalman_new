import React, { useState, useRef, useEffect, useCallback } from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import axios from 'axios';
import { API_BASE } from '@/lib/api';

// @ts-ignore
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.js`;

const PdfCropper = () => {
    const [pdf, setPdf] = useState<any>(null);
    const [pageNumber, setPageNumber] = useState(1);
    const [numPages, setNumPages] = useState(0);
    const [scale, setScale] = useState(1.0);
    const [isSelecting, setIsSelecting] = useState(false);
    const [selection, setSelection] = useState<any>(null);
    const [savedCount, setSavedCount] = useState(0);
    const [isProcessing, setIsProcessing] = useState(false);
    const [pdfFile, setPdfFile] = useState<File | null>(null);
    const [isSettingRefImage, setIsSettingRefImage] = useState(false);
    
    const canvasRef = useRef<HTMLCanvasElement>(null);
    const overlayRef = useRef<HTMLCanvasElement>(null);
    const viewerRef = useRef<HTMLDivElement>(null);
    const scrollContainerRef = useRef<HTMLDivElement>(null);
    const renderTimeoutRef = useRef<any>(null);

    const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (file) {
            setPdfFile(file);
            const loadingTask = pdfjsLib.getDocument(URL.createObjectURL(file));
            const pdfDoc = await loadingTask.promise;
            setPdf(pdfDoc);
            setNumPages(pdfDoc.numPages);
            renderPage(1, pdfDoc, scale);
        }
    };

    const setReferenceImage = async () => {
        if (!pdfFile) return;
        setIsSettingRefImage(true);
        try {
            const formData = new FormData();
            formData.append('drawing_pdf', pdfFile);
            const response = await axios.post(`${API_BASE}/api/welding/set-reference-image`, formData);
            if (response.data.status === 'success') {
                alert('Reference image set on the check sheet.');
            } else {
                alert('Error: ' + response.data.message);
            }
        } catch (error: any) {
            const detail = error?.response?.data?.detail;
            alert(detail || 'Could not find a reference isometric view in this drawing.');
        } finally {
            setIsSettingRefImage(false);
        }
    };

    const renderPage = useCallback(async (num: number, pdfDoc = pdf, currentScale = scale) => {
        if (!pdfDoc) return;
        const page = await pdfDoc.getPage(num);
        const viewport = page.getViewport({ scale: currentScale });
        const canvas = canvasRef.current;
        if (!canvas) return;
        
        const context = canvas.getContext('2d');
        if (!context) return;
        
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        const renderContext = {
            canvasContext: context,
            viewport: viewport
        };
        await page.render(renderContext).promise;
        
        if (overlayRef.current) {
            overlayRef.current.width = canvas.width;
            overlayRef.current.height = canvas.height;
        }
    }, [pdf, scale]);

    useEffect(() => {
        if (pdf) {
            renderPage(pageNumber);
        }
    }, [scale, pageNumber, renderPage]);

    useEffect(() => {
        const handleWheel = (e: WheelEvent) => {
            if (e.ctrlKey) {
                e.preventDefault();
                
                const container = scrollContainerRef.current;
                if (!container) return;

                const rect = container.getBoundingClientRect();
                const mouseX = e.clientX - rect.left;
                const mouseY = e.clientY - rect.top;

                const pointX = (container.scrollLeft + mouseX) / scale;
                const pointY = (container.scrollTop + mouseY) / scale;

                const zoomFactor = 1.05; 
                let newScale = e.deltaY > 0 ? scale / zoomFactor : scale * zoomFactor;
                newScale = Math.min(10, Math.max(0.1, newScale));

                if (Math.abs(newScale - scale) < 0.001) return;

                clearTimeout(renderTimeoutRef.current);
                renderTimeoutRef.current = setTimeout(() => {
                    setScale(newScale);
                    setTimeout(() => {
                        if (scrollContainerRef.current) {
                            scrollContainerRef.current.scrollLeft = pointX * newScale - mouseX;
                            scrollContainerRef.current.scrollTop = pointY * newScale - mouseY;
                        }
                    }, 0);
                }, 16);
            }
        };

        const container = scrollContainerRef.current;
        if (container) {
            container.addEventListener('wheel', handleWheel, { passive: false });
        }
        return () => {
            if (container) {
                container.removeEventListener('wheel', handleWheel);
            }
        };
    }, [scale]);

    const handleMouseDown = (e: React.MouseEvent) => {
        if (!overlayRef.current) return;
        const rect = overlayRef.current.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
        setSelection({ startX: x, startY: y, endX: x, endY: y });
        setIsSelecting(true);
    };

    const handleMouseMove = (e: React.MouseEvent) => {
        if (!isSelecting || !overlayRef.current) return;
        const rect = overlayRef.current.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
        setSelection((prev: any) => ({ ...prev, endX: x, endY: y }));
        drawSelection(x, y);
    };

    const handleMouseUp = () => {
        setIsSelecting(false);
    };

    const drawSelection = (currentX: number, currentY: number) => {
        if (!overlayRef.current || !selection) return;
        const canvas = overlayRef.current;
        const ctx = canvas.getContext('2d');
        if (!ctx) return;
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.strokeStyle = '#ff0000';
        ctx.lineWidth = 2;
        ctx.strokeRect(
            selection.startX, 
            selection.startY, 
            currentX - selection.startX, 
            currentY - selection.startY
        );
    };

    const captureCrop = async () => {
        if (!selection || !pdf || !overlayRef.current) return;
        setIsProcessing(true);
        
        try {
            const { startX, startY, endX, endY } = selection;
            const uiWidth = Math.abs(endX - startX);
            const uiHeight = Math.abs(endY - startY);
            const uiLeft = Math.min(startX, endX);
            const uiTop = Math.min(startY, endY);

            if (uiWidth === 0 || uiHeight === 0) return;

            const highResScale = 4.0;
            const page = await pdf.getPage(pageNumber);
            const viewport = page.getViewport({ scale: highResScale });

            const ratio = highResScale / scale;
            const cropLeft = uiLeft * ratio;
            const cropTop = uiTop * ratio;
            const cropWidth = uiWidth * ratio;
            const cropHeight = uiHeight * ratio;

            const highResCanvas = document.createElement('canvas');
            highResCanvas.width = viewport.width;
            highResCanvas.height = viewport.height;
            const highResCtx = highResCanvas.getContext('2d');
            if (!highResCtx) return;
            
            await page.render({
                canvasContext: highResCtx,
                viewport: viewport
            }).promise;

            const finalCanvas = document.createElement('canvas');
            finalCanvas.width = cropWidth;
            finalCanvas.height = cropHeight;
            const finalCtx = finalCanvas.getContext('2d');
            if (!finalCtx) return;
            
            finalCtx.drawImage(
                highResCanvas,
                cropLeft, cropTop, cropWidth, cropHeight,
                0, 0, cropWidth, cropHeight
            );

            const imageData = finalCanvas.toDataURL('image/png', 1.0);
            
            const formData = new FormData();
            formData.append('image_data', imageData);

            const response = await axios.post(`${API_BASE}/api/save-crop`, formData);
            if (response.data.status === 'success') {
                setSavedCount(response.data.sl_no);
                const ctx = overlayRef.current.getContext('2d');
                if (ctx) ctx.clearRect(0, 0, overlayRef.current.width, overlayRef.current.height);
                setSelection(null);
            } else {
                alert('Error: ' + response.data.message);
            }
        } catch (error) {
            console.error(error);
            alert('Error saving crop');
        } finally {
            setIsProcessing(false);
        }
    };

    const downloadExcel = () => {
        window.open(`${API_BASE}/api/download-excel`, '_blank');
    };

    const resetExcel = async () => {
        if (confirm('Are you sure?')) {
            await axios.post(`${API_BASE}/api/reset-excel`);
            setSavedCount(0);
        }
    };

    const toggleFullScreen = () => {
        if (!viewerRef.current) return;
        if (!document.fullscreenElement) {
            viewerRef.current.requestFullscreen();
        } else {
            document.exitFullscreen();
        }
    };

    return (
        <div ref={viewerRef} style={{ display: 'flex', flexDirection: 'column', height: '100%', width: '100%', backgroundColor: 'hsl(var(--secondary))', fontFamily: 'var(--font-sans)', overflow: 'hidden' }}>
            <div style={{ height: '64px', minHeight: '64px', padding: '0 24px', backgroundColor: '#ffffff', borderBottom: '1px solid hsl(var(--border))', display: 'flex', alignItems: 'center', gap: '20px', zIndex: 100, boxSizing: 'border-box' }}>
                <div style={{ fontWeight: '700', color: 'hsl(var(--primary))', fontSize: '1.25rem', fontFamily: 'var(--font-display)', letterSpacing: '-0.02em' }}>METALMAN <span style={{ color: 'hsl(var(--foreground))', fontWeight: '400' }}>CROPPER</span></div>
                
                <input type="file" accept="application/pdf" onChange={onFileChange} style={{ fontSize: '0.85rem', cursor: 'pointer' }} />
                
                <div style={{ display: 'flex', alignItems: 'center', border: '1px solid hsl(var(--border))', borderRadius: 'var(--radius)', overflow: 'hidden', backgroundColor: 'hsl(var(--muted))' }}>
                    <button onClick={() => setScale(s => s / 1.1)} style={{ padding: '8px 16px', background: 'transparent', border: 'none', borderRight: '1px solid hsl(var(--border))', cursor: 'pointer', fontWeight: '600' }}>-</button>
                    <span style={{ padding: '0 16px', minWidth: '70px', textAlign: 'center', fontSize: '0.9rem', fontWeight: '500' }}>{Math.round(scale * 100)}%</span>
                    <button onClick={() => setScale(s => s * 1.1)} style={{ padding: '8px 16px', background: 'transparent', border: 'none', borderLeft: '1px solid hsl(var(--border))', cursor: 'pointer', fontWeight: '600' }}>+</button>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'hsl(var(--muted))', padding: '4px', borderRadius: 'var(--radius)', border: '1px solid hsl(var(--border))' }}>
                    <button disabled={pageNumber <= 1} onClick={() => setPageNumber(p => p - 1)} style={{ padding: '6px 12px', background: 'white', border: '1px solid hsl(var(--border))', borderRadius: 'calc(var(--radius) - 2px)', cursor: 'pointer', opacity: pageNumber <= 1 ? 0.5 : 1 }}>←</button>
                    <span style={{ fontSize: '0.9rem', fontWeight: '500', minWidth: '60px', textAlign: 'center' }}>{pageNumber} / {numPages}</span>
                    <button disabled={pageNumber >= numPages} onClick={() => setPageNumber(p => p + 1)} style={{ padding: '6px 12px', background: 'white', border: '1px solid hsl(var(--border))', borderRadius: 'calc(var(--radius) - 2px)', cursor: 'pointer', opacity: pageNumber >= numPages ? 0.5 : 1 }}>→</button>
                </div>

                <button onClick={toggleFullScreen} style={{ padding: '8px 16px', borderRadius: 'var(--radius)', border: '1px solid hsl(var(--border))', background: '#fff', cursor: 'pointer', fontSize: '0.875rem', fontWeight: '500', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{ fontSize: '1.1rem' }}>⛶</span> Fullscreen
                </button>

                <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: '20px' }}>
                    <span style={{ fontSize: '0.9rem', color: 'hsl(var(--muted-foreground))' }}>Items: <b style={{ color: 'hsl(var(--primary))' }}>{savedCount}</b></span>
                    <button
                        onClick={setReferenceImage}
                        disabled={!pdfFile || isSettingRefImage}
                        title="Locate this drawing's Reference Isometric View and place it in the check sheet's photo row"
                        style={{
                            background: '#fff',
                            color: 'hsl(var(--primary))',
                            border: '1px solid hsl(var(--primary))',
                            padding: '10px 20px',
                            borderRadius: 'var(--radius)',
                            cursor: pdfFile ? 'pointer' : 'not-allowed',
                            fontWeight: '600',
                            whiteSpace: 'nowrap',
                            flexShrink: 0,
                            opacity: !pdfFile || isSettingRefImage ? 0.6 : 1,
                        }}
                    >
                        {isSettingRefImage ? 'Locating…' : 'Set Reference Image'}
                    </button>
                    <button
                        onClick={captureCrop}
                        disabled={!selection || isProcessing} 
                        style={{ 
                            background: 'hsl(var(--accent))', 
                            color: 'white', 
                            border: 'none', 
                            padding: '10px 24px', 
                            borderRadius: 'var(--radius)', 
                            cursor: 'pointer', 
                            fontWeight: '600',
                            letterSpacing: '0.01em',
                            boxShadow: 'var(--shadow-card)'
                        }}
                    >
                        {isProcessing ? 'Processing...' : 'SAVE CROP'}
                    </button>
                    <button 
                        onClick={downloadExcel} 
                        disabled={savedCount === 0} 
                        style={{ 
                            background: 'hsl(var(--primary))', 
                            color: 'white', 
                            border: 'none', 
                            padding: '10px 24px', 
                            borderRadius: 'var(--radius)', 
                            cursor: 'pointer',
                            fontWeight: '600',
                            opacity: savedCount === 0 ? 0.6 : 1
                        }}
                    >
                        DOWNLOAD
                    </button>
                    <button onClick={resetExcel} style={{ color: 'hsl(var(--destructive, 0 75% 50%))', background: 'none', border: 'none', cursor: 'pointer', fontSize: '0.85rem', fontWeight: '500' }}>RESET</button>
                </div>
            </div>

            <div 
                ref={scrollContainerRef}
                style={{ 
                    flex: 1, 
                    overflow: 'auto', 
                    padding: '80px', 
                    display: 'flex', 
                    justifyContent: 'flex-start', 
                    alignItems: 'flex-start', 
                    backgroundColor: 'hsl(215 15% 25%)'  /* Darker gray for canvas contrast */
                }}
            >
                <div style={{ 
                    position: 'relative', 
                    background: '#fff', 
                    boxShadow: '0 20px 50px rgba(0,0,0,0.3)',
                    width: 'max-content',
                    height: 'max-content',
                    margin: '0 auto',
                    borderRadius: '2px'
                }}>
                    <canvas ref={canvasRef} style={{ display: 'block' }} />
                    <canvas 
                        ref={overlayRef} 
                        onMouseDown={handleMouseDown}
                        onMouseMove={handleMouseMove}
                        onMouseUp={handleMouseUp}
                        style={{ 
                            position: 'absolute', 
                            top: 0, 
                            left: 0, 
                            cursor: 'crosshair',
                            zIndex: 10
                        }}
                    />
                </div>
            </div>
            
            <div style={{ background: '#ffffff', padding: '8px 24px', fontSize: '0.8rem', color: 'hsl(var(--muted-foreground))', borderTop: '1px solid hsl(var(--border))', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                 <span>Hold <kbd style={{ background: 'hsl(var(--muted))', padding: '2px 6px', borderRadius: '4px', border: '1px solid hsl(var(--border))' }}>Ctrl</kbd> + <kbd style={{ background: 'hsl(var(--muted))', padding: '2px 6px', borderRadius: '4px', border: '1px solid hsl(var(--border))' }}>Scroll</kbd> to zoom • Target pointer for precise navigation</span>
                 <span style={{ fontWeight: '500' }}>Metalman Cropping System v2.0</span>
            </div>
        </div>
    );
};

export default PdfCropper;

/* eslint-disable @typescript-eslint/no-explicit-any */
import { Injectable, signal, PLATFORM_ID, inject } from '@angular/core';
import { isPlatformBrowser } from '@angular/common';
import { PdfDb } from './pdf-db';
import { MarkdownRenderer } from './markdown-renderer';
import { DocxExporter } from './docx/docx-exporter';
import { PDFDocument } from 'pdf-lib';

export interface PdfPageData {
  pageNum: number;
  items: any[];
  pageImageUrl: string; // PNG Data URL, scale 1.4 (lazy rendered)
  extractedImages?: any[];
}

@Injectable({
  providedIn: 'root'
})
export class PdfProcessor {
  private platformId = inject(PLATFORM_ID);
  private db = new PdfDb(this.platformId);
  private pdfjsLib: any = null;
  private currentPdfDoc: any = null;
  private currentFileName = '';
  private loadedPdfDocs = new Map<string, any>();
  private pageToDocMap = new Map<number, { doc: any; localPageNum: number }>();
  
  isScriptLoaded = signal(false);

  async loadPdfEngine(updateStatus: (msg: string) => void, setError: (msg: string) => void): Promise<void> {
    if (!isPlatformBrowser(this.platformId)) return;
    if (this.pdfjsLib) {
      this.isScriptLoaded.set(true);
      return;
    }

    try {
      if ((window as any).pdfjsLib) {
        this.pdfjsLib = (window as any).pdfjsLib;
        this.isScriptLoaded.set(true);
        return;
      }

      updateStatus('Đang nạp trình kết xuất PDF...');
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      script.async = true;

      await new Promise<void>((resolve, reject) => {
        script.onload = () => {
          this.pdfjsLib = (window as any).pdfjsLib;
          if (this.pdfjsLib) {
            this.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            this.isScriptLoaded.set(true);
            resolve();
          } else {
            reject(new Error('Không tìm thấy pdfjsLib sau khi nạp script.'));
          }
        };
        script.onerror = () => reject(new Error('Lỗi khi tải thư viện PDF.js'));
        document.head.appendChild(script);
      });
    } catch (e: any) {
      setError('Lỗi nạp thư viện PDF.js: ' + (e.message || e));
    }
  }

  isLoaded(): boolean {
    return this.isScriptLoaded();
  }

  getPdfjsLib(): any {
    return this.pdfjsLib || (typeof window !== 'undefined' ? (window as any).pdfjsLib : null);
  }

  async loadRestoredPdfDocuments(files: File[], isMultiFileMode: boolean, chunks: any[]): Promise<void> {
    if (!this.pdfjsLib) {
      await this.loadPdfEngine(() => undefined, () => undefined);
    }
    const lib = this.getPdfjsLib();
    if (!lib) return;

    this.pageToDocMap.clear();
    this.loadedPdfDocs.clear();

    if (!isMultiFileMode && files.length === 1) {
      const file = files[0];
      try {
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = lib.getDocument({ data: arrayBuffer });
        this.currentPdfDoc = await loadingTask.promise;
        this.currentFileName = file.name;
        this.loadedPdfDocs.set(file.name, this.currentPdfDoc);
        
        let maxPage = 0;
        if (chunks && chunks.length > 0) {
          maxPage = chunks[chunks.length - 1].endPageNum;
        }
        for (let i = 1; i <= maxPage; i++) {
          this.pageToDocMap.set(i, { doc: this.currentPdfDoc, localPageNum: i });
        }
      } catch (e) {
        console.warn('Could not restore pdfjsDoc for single file:', e);
      }
    } else if (isMultiFileMode && files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        try {
          const arrayBuffer = await file.arrayBuffer();
          const loadingTask = lib.getDocument({ data: arrayBuffer });
          const filePdfDoc = await loadingTask.promise;
          this.loadedPdfDocs.set(file.name, filePdfDoc);
          
          if (i === 0) {
            this.currentPdfDoc = filePdfDoc;
            this.currentFileName = file.name;
          }
          
          const chunk = chunks[i];
          if (chunk) {
            let localPageNum = 1;
            for (let globalPageNum = chunk.startPageNum; globalPageNum <= chunk.endPageNum; globalPageNum++) {
              this.pageToDocMap.set(globalPageNum, { doc: filePdfDoc, localPageNum: localPageNum++ });
            }
          }
        } catch (e) {
          console.warn(`Could not restore pdfjsDoc for ${file.name}:`, e);
        }
      }
    }
  }

  async loadPdfDocument(file: File): Promise<any> {
    if (!this.pdfjsLib) {
      await this.loadPdfEngine(() => undefined, () => undefined);
    }
    const lib = this.getPdfjsLib();
    if (!lib) {
      throw new Error('Thư viện PDF.js chưa sẵn sàng.');
    }
    const arrayBuffer = await file.arrayBuffer();
    const loadingTask = lib.getDocument({ data: arrayBuffer });
    this.currentPdfDoc = await loadingTask.promise;
    this.currentFileName = file.name;
    return this.currentPdfDoc;
  }

  private async renderPageFromDoc(doc: any, localPageNum: number): Promise<string> {
    try {
      const page = await doc.getPage(localPageNum);
      const dpr = typeof window !== 'undefined' ? Math.min(window.devicePixelRatio || 1, 1.5) : 1;
      const scale = Math.min(1.3 * dpr, 2.0); // Optimized scale to prevent huge canvas memory spikes
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) return '';
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);

      const renderContext = {
        canvasContext: ctx,
        viewport: viewport
      };
      await page.render(renderContext).promise;
      return canvas.toDataURL('image/jpeg', 0.85); // JPEG is 10x smaller than PNG in RAM and renders faster
    } catch (err) {
      console.error(`Error rendering local page ${localPageNum} to image:`, err);
      return '';
    }
  }

  async renderPageToPng(pageNum: number): Promise<string> {
    const mapping = this.pageToDocMap.get(pageNum);
    if (mapping && mapping.doc) {
      return this.renderPageFromDoc(mapping.doc, mapping.localPageNum);
    }
    if (this.currentPdfDoc) {
      return this.renderPageFromDoc(this.currentPdfDoc, pageNum);
    }
    return '';
  }

  async extractPdfChunks(file: File, onProgress: (msg: string) => void): Promise<{ pages: PdfPageData[], chunks: any[] }> {
    this.pageToDocMap.clear();
    this.loadedPdfDocs.clear();

    onProgress('Đọc tài liệu PDF...');
    let numPages = 0;
    try {
      const doc = await this.loadPdfDocument(file);
      numPages = doc.numPages;
      if (this.currentPdfDoc) {
        this.loadedPdfDocs.set(file.name, this.currentPdfDoc);
      }
    } catch {
      const arrayBuffer = await file.arrayBuffer();
      const srcDoc = await PDFDocument.load(arrayBuffer);
      numPages = srcDoc.getPageCount();
    }

    if (numPages > 500) {
      throw new Error(`Tài liệu có ${numPages} trang, vượt quá giới hạn cho phép (tối đa 500 trang). Vui lòng chia nhỏ tài liệu trước khi xử lý.`);
    }

    const itemsExtracted: PdfPageData[] = [];
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      itemsExtracted.push({
        pageNum,
        items: [],
        pageImageUrl: '',
        extractedImages: []
      });
      if (this.currentPdfDoc) {
        this.pageToDocMap.set(pageNum, { doc: this.currentPdfDoc, localPageNum: pageNum });
      }
    }

    // Optimize chunk splitting: target ~50MB per chunk (max 12 pages)
    // For a 100-110MB file, this divides into 2 chunks (just like a 50MB file) for maximum speed.
    const MAX_CHUNK_PAGES = 12;
    const TARGET_MAX_CHUNK_MB = 55;
    const fileSizeMb = file.size / (1024 * 1024);
    const avgMbPerPage = fileSizeMb / Math.max(1, numPages);

    let targetPagesPerChunk = MAX_CHUNK_PAGES;
    if (avgMbPerPage > 0) {
      const sizeBasedLimit = Math.max(1, Math.floor(TARGET_MAX_CHUNK_MB / avgMbPerPage));
      targetPagesPerChunk = Math.min(MAX_CHUNK_PAGES, sizeBasedLimit);
    }

    const createChunks = (pages: PdfPageData[]): any[] => {
      const chunks: any[] = [];
      const divide = (p: PdfPageData[]) => {
        if (p.length <= targetPagesPerChunk) {
          if (p.length > 0) {
            chunks.push({
              id: '',
              originalFileName: '',
              index: chunks.length,
              startPageNum: p[0].pageNum,
              endPageNum: p[p.length - 1].pageNum,
              pages: p,
              status: 'pending',
              errorMessage: '',
              markdownContent: '',
              reflowHtml: ''
            });
          }
          return;
        }
        const mid = Math.ceil(p.length / 2);
        divide(p.slice(0, mid));
        divide(p.slice(mid));
      };
      divide(pages);
      return chunks;
    };

    const generatedChunks = createChunks(itemsExtracted);

    let chunkCounter = 1;
    for (const chunk of generatedChunks) {
      chunk.id = `Phần ${chunkCounter}`;
      chunkCounter++;
    }

    // On-demand rendering: Render immediately only the first 1-2 pages of chunk 1
    if (generatedChunks.length > 0) {
      onProgress('Đang chuẩn bị Bản gốc...');
      const firstChunk = generatedChunks[0];
      const initialPages = firstChunk.pages.slice(0, 2);
      for (const page of initialPages) {
        if (!page.pageImageUrl) {
          page.pageImageUrl = await this.renderPageToPng(page.pageNum);
        }
      }
    }

    return { pages: itemsExtracted, chunks: generatedChunks };
  }

  async extractMultiplePdfChunks(files: File[], onProgress: (msg: string) => void): Promise<{ pages: PdfPageData[], chunks: any[] }> {
    if (!this.pdfjsLib) {
      await this.loadPdfEngine(() => undefined, () => undefined);
    }
    const lib = this.getPdfjsLib();
    if (!lib) {
      throw new Error('Thư viện PDF.js chưa sẵn sàng.');
    }

    this.pageToDocMap.clear();
    this.loadedPdfDocs.clear();

    const allPages: PdfPageData[] = [];
    const allChunks: any[] = [];
    let globalPageCounter = 1;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      onProgress(`Đang nạp tệp ${i + 1}/${files.length}: ${file.name}...`);

      const arrayBuffer = await file.arrayBuffer();
      const srcDoc = await PDFDocument.load(arrayBuffer);
      const fileNumPages = srcDoc.getPageCount();

      if (fileNumPages > 12) {
        throw new Error(`Tệp "${file.name}" có ${fileNumPages} trang, vượt quá giới hạn tối đa 12 trang/tệp cho chế độ tải nhiều file.`);
      }

      // Load pdfjsDoc for this file
      let filePdfDoc: any = null;
      try {
        const loadingTask = lib.getDocument({ data: arrayBuffer.slice(0) });
        filePdfDoc = await loadingTask.promise;
        this.loadedPdfDocs.set(file.name, filePdfDoc);
        if (i === 0) {
          this.currentPdfDoc = filePdfDoc;
          this.currentFileName = file.name;
        }
      } catch (e) {
        console.warn(`Could not load pdfjsDoc for ${file.name}:`, e);
      }

      const filePages: PdfPageData[] = [];
      const startPageNum = globalPageCounter;

      for (let localPageNum = 1; localPageNum <= fileNumPages; localPageNum++) {
        const globalPageNum = globalPageCounter++;
        const pageData: PdfPageData = {
          pageNum: globalPageNum,
          items: [],
          pageImageUrl: '',
          extractedImages: []
        };
        filePages.push(pageData);
        allPages.push(pageData);

        if (filePdfDoc) {
          this.pageToDocMap.set(globalPageNum, { doc: filePdfDoc, localPageNum });
        }
      }

      const endPageNum = globalPageCounter - 1;

      allChunks.push({
        id: `Phần ${i + 1}`,
        originalFileName: file.name,
        index: i,
        startPageNum,
        endPageNum,
        pages: filePages,
        status: 'pending',
        errorMessage: '',
        markdownContent: '',
        reflowHtml: ''
      });
    }

    // Render first chunk pages immediately
    if (allChunks.length > 0) {
      onProgress('Đang render trước ảnh Bản gốc Phần 1...');
      const firstChunk = allChunks[0];
      for (const page of firstChunk.pages) {
        if (!page.pageImageUrl) {
          page.pageImageUrl = await this.renderPageToPng(page.pageNum);
        }
      }
    }

    return { pages: allPages, chunks: allChunks };
  }

  formatBytes(bytes: number, decimals = 1): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  }

  /**
   * IndexedDB Persistence delegation
   */
  openDb(): Promise<IDBDatabase> {
    return this.db.openDb();
  }

  saveHistoryItem(item: any): Promise<void> {
    return this.db.saveHistoryItem(item);
  }

  getAllHistoryItems(): Promise<any[]> {
    return this.db.getAllHistoryItems();
  }

  deleteHistoryItem(id: string): Promise<void> {
    return this.db.deleteHistoryItem(id);
  }

  /**
   * Markdown preview and XHTML compiler rendering delegation
   */
  renderMarkdownToHtml(markdown: string, pdfPages: PdfPageData[]): string {
    return MarkdownRenderer.renderMarkdownToHtml(markdown, pdfPages);
  }

  renderHtmlContent(htmlContent: string, pdfPages: PdfPageData[]): string {
    return MarkdownRenderer.renderHtmlContent(htmlContent, pdfPages);
  }

  renderContent(content: string, pdfPages: PdfPageData[], outputMode: 'markdown' | 'html' = 'markdown'): string {
    if (outputMode === 'html') {
      return MarkdownRenderer.renderHtmlContent(content, pdfPages);
    }
    return MarkdownRenderer.renderMarkdownToHtml(content, pdfPages);
  }

  /**
   * DOCX Word generation delegation
   */
  generateDocx(title: string, content: string, pdfPages: PdfPageData[]): Promise<Blob> {
    return DocxExporter.generateDocx(title, content, pdfPages);
  }
}


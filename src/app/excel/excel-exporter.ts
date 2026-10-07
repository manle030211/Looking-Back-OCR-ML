/* eslint-disable @typescript-eslint/no-explicit-any */
import * as XLSX from 'xlsx';
import { PdfChunk } from '../services/document-processing.service';

export interface TableData {
  title?: string;
  headers: string[];
  rows: (string | number)[][];
  sourceInfo?: string;
}

export class ExcelExporter {
  /**
   * Tạo workbook Excel hoàn chỉnh từ danh sách các khối (chunks) đã xử lý
   */
  static generateExcelBlob(fileName: string, chunks: PdfChunk[]): Blob {
    const wb = XLSX.utils.book_new();

    // 1. Thu thập tất cả các bảng biểu từ HTML và Markdown
    const allTables: TableData[] = [];
    chunks.forEach((chunk) => {
      const chunkLabel = `${chunk.id} (Trang ${chunk.startPageNum}-${chunk.endPageNum})`;
      
      // Trích xuất từ HTML nếu có
      if (chunk.reflowHtml) {
        const htmlTables = this.extractTablesFromHtml(chunk.reflowHtml, chunkLabel);
        allTables.push(...htmlTables);
      }
      
      // Nếu chưa có bảng nào từ HTML và có markdownContent, thử trích xuất từ Markdown
      if (allTables.length === 0 && chunk.markdownContent) {
        const mdTables = this.extractTablesFromMarkdown(chunk.markdownContent, chunkLabel);
        allTables.push(...mdTables);
      }
    });

    // 2. Tạo Sheet "Tổng hợp nội dung" (Structured Content Sheet)
    const summaryRows: any[][] = [
      ['TÀI LIỆU', fileName],
      ['TỔNG SỐ KHỐI', chunks.length],
      ['SỐ BẢNG PHÁT HIỆN', allTables.length],
      ['NGÀY XUẤT', new Date().toLocaleString('vi-VN')],
      [] // Dòng trống
    ];

    // Tiêu đề các cột dữ liệu
    summaryRows.push(['Khối', 'Trang', 'Phân loại', 'Nội dung chi tiết']);

    chunks.forEach((chunk) => {
      const chunkName = chunk.id;
      const pagesStr = `${chunk.startPageNum} - ${chunk.endPageNum}`;

      // Phân tách nội dung thành các dòng cho bảng tính
      const content = chunk.markdownContent || this.stripHtml(chunk.reflowHtml || '');
      const paragraphs = content
        .split(/\n\s*\n/)
        .map(p => p.trim())
        .filter(p => p.length > 0);

      if (paragraphs.length === 0) {
        summaryRows.push([chunkName, pagesStr, 'Chưa có nội dung', '']);
      } else {
        paragraphs.forEach((para) => {
          let type = 'Đoạn văn';
          let text = para;

          if (para.startsWith('# ')) {
            type = 'Tiêu đề 1';
            text = para.replace(/^#\s+/, '');
          } else if (para.startsWith('## ')) {
            type = 'Tiêu đề 2';
            text = para.replace(/^##\s+/, '');
          } else if (para.startsWith('### ')) {
            type = 'Tiêu đề 3';
            text = para.replace(/^###\s+/, '');
          } else if (para.startsWith('- ') || para.startsWith('* ') || /^\d+\.\s/.test(para)) {
            type = 'Danh sách';
          } else if (para.startsWith('|') && para.endsWith('|')) {
            type = 'Bảng dữ liệu';
          }

          summaryRows.push([chunkName, pagesStr, type, text]);
        });
      }
      // Dòng ngăn cách giữa các khối
      summaryRows.push([]);
    });

    const summaryWs = XLSX.utils.aoa_to_sheet(summaryRows);
    summaryWs['!cols'] = this.autoFitColumns(summaryRows, [14, 12, 16, 75]);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Nội dung tổng hợp');

    // 3. Nếu phát hiện các bảng biểu, tạo từng Sheet riêng cho từng bảng
    if (allTables.length > 0) {
      allTables.forEach((table, idx) => {
        const sheetName = this.sanitizeSheetName(`Bảng ${idx + 1}`);
        const tableRows: any[][] = [];

        if (table.sourceInfo) {
          tableRows.push(['Nguồn:', table.sourceInfo]);
          tableRows.push([]);
        }

        if (table.headers.length > 0) {
          tableRows.push(table.headers);
        }

        table.rows.forEach(row => {
          tableRows.push(row);
        });

        const tableWs = XLSX.utils.aoa_to_sheet(tableRows);
        tableWs['!cols'] = this.autoFitColumns(tableRows);
        XLSX.utils.book_append_sheet(wb, tableWs, sheetName);
      });
    }

    // 4. Xuất file binary sang Blob
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Blob([wbout], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  /**
   * Tạo file Excel cho duy nhất một khối
   */
  static generateChunkExcelBlob(fileName: string, chunk: PdfChunk): Blob {
    return this.generateExcelBlob(fileName, [chunk]);
  }

  /**
   * Trích xuất bảng từ mã HTML
   */
  private static extractTablesFromHtml(html: string, sourceInfo: string): TableData[] {
    const tables: TableData[] = [];
    if (typeof DOMParser === 'undefined') return tables;

    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      const tableEls = doc.querySelectorAll('table');

      tableEls.forEach((tableEl) => {
        let headers: string[] = [];
        const rows: (string | number)[][] = [];

        const trEls = tableEl.querySelectorAll('tr');
        trEls.forEach((tr) => {
          const thEls = tr.querySelectorAll('th');
          const tdEls = tr.querySelectorAll('td');

          if (thEls.length > 0 && headers.length === 0) {
            headers = Array.from(thEls).map(th => (th.textContent || '').trim());
          } else {
            const cells = Array.from(thEls.length > 0 ? thEls : tdEls).map(cell => {
              const text = (cell.textContent || '').trim();
              return this.parseCellNumber(text);
            });
            if (cells.length > 0) {
              rows.push(cells);
            }
          }
        });

        if (headers.length === 0 && rows.length > 0) {
          headers = rows.shift()?.map(c => String(c)) || [];
        }

        if (headers.length > 0 || rows.length > 0) {
          tables.push({ headers, rows, sourceInfo });
        }
      });
    } catch (e) {
      console.warn('Lỗi phân tích bảng HTML cho Excel:', e);
    }

    return tables;
  }

  /**
   * Trích xuất bảng từ cú pháp Markdown
   */
  private static extractTablesFromMarkdown(md: string, sourceInfo: string): TableData[] {
    const tables: TableData[] = [];
    const lines = md.split(/\r?\n/);
    let currentTableLines: string[] = [];

    const flushTable = () => {
      if (currentTableLines.length >= 2) {
        const rowsData = currentTableLines.filter(line => {
          const clean = line.trim();
          return !/^[|:\-\s]+$/.test(clean);
        });

        if (rowsData.length > 0) {
          const parsed = rowsData.map(line => {
            const cols = line.split('|').map(c => c.trim());
            if (line.startsWith('|')) cols.shift();
            if (line.endsWith('|')) cols.pop();
            return cols;
          });

          const headers = parsed[0] || [];
          const rows = parsed.slice(1).map(row => row.map(cell => this.parseCellNumber(cell)));

          if (headers.length > 0 || rows.length > 0) {
            tables.push({ headers, rows, sourceInfo });
          }
        }
      }
      currentTableLines = [];
    };

    for (const line of lines) {
      const clean = line.trim();
      if (clean.startsWith('|') && clean.endsWith('|')) {
        currentTableLines.push(clean);
      } else {
        flushTable();
      }
    }
    flushTable();

    return tables;
  }

  /**
   * Tự động nhận diện chuỗi số để chuyển sang kiểu Number trong Excel
   */
  private static parseCellNumber(val: string): string | number {
    if (!val) return '';
    const clean = val.replace(/\s+/g, '');
    // Chuỗi số nguyên hoặc thập phân hợp lệ (ví dụ: 123, 123.45, 123,45)
    if (/^-?\d+([.,]\d+)?$/.test(clean)) {
      const normalized = clean.replace(',', '.');
      const num = parseFloat(normalized);
      if (!isNaN(num) && isFinite(num)) {
        return num;
      }
    }
    return val;
  }

  /**
   * Loại bỏ các thẻ HTML để lấy văn bản thuần
   */
  private static stripHtml(html: string): string {
    return html
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/\s+/g, ' ')
      .trim();
  }

  /**
   * Đảm bảo tên Sheet hợp lệ trong Excel (tối đa 31 ký tự, không chứa ký tự cấm: : \ / ? * [ ])
   */
  private static sanitizeSheetName(name: string): string {
    const cleaned = name.replace(/[\\/?*[\]:]/g, '_').trim();
    return cleaned.slice(0, 31) || 'Sheet1';
  }

  /**
   * Tính toán độ rộng cột tự động dựa trên độ dài dữ liệu
   */
  private static autoFitColumns(rows: any[][], customDefaults?: number[]): { wch: number }[] {
    const colWidths: number[] = [...(customDefaults || [])];

    rows.forEach(row => {
      if (!Array.isArray(row)) return;
      row.forEach((cell, colIdx) => {
        const str = cell !== null && cell !== undefined ? String(cell) : '';
        const len = Math.min(Math.max(str.length, 8), 65);
        colWidths[colIdx] = Math.max(colWidths[colIdx] || 10, len + 2);
      });
    });

    return colWidths.map(w => ({ wch: w || 12 }));
  }
}

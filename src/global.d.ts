// Global type declarations for external libraries loaded via CDN

declare global {
  interface Window {
    pdfjsLib: {
      getDocument: (source: { data: Uint8Array }) => { promise: Promise<any> };
      GlobalWorkerOptions: {
        workerSrc: string;
      };
      Util: {
        transform: (a: number[], b: number[]) => number[];
      };
    };
    XLSX: {
      read: (data: ArrayBuffer, options: { type: string }) => {
        SheetNames: string[];
        Sheets: Record<string, any>;
      };
      utils: {
        sheet_to_json: (sheet: any, options: { header: number; defval: string }) => any[];
      };
    };
    docx: {
      renderAsync: (buffer: ArrayBuffer, container: HTMLElement) => Promise<void>;
    };
    tf: {
      tensor2d: (data: number[][]) => any;
      tidy: (fn: () => void) => void;
    };
  }
}

export {};

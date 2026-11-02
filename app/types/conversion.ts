export type ConversionId = 
| "word-to-markdown"
| "pdf-to-markdown"
| "powerpoint-to-markdown"
| "markdown-to-word"
| "markdown-to-pdf"
| "markdown-to-powerpoint";

export type ConversionStatus = 
| "idle"
| "validating"
| "converting"
| "completed"
| "error";

export interface ConversionOption {
  id: ConversionId;
  title: string;
  description: string;
  inputExtensions: ".docx" | ".pdf" | ".pptx" | ".md";
  outputExtensions: ".md" | ".docx" | ".pdf" | ".pptx";
}

export interface ConversionResult {
  blob: Blob;
  fileName: string;
  mimetype: string;
}

export interface ConversionError {
  message: string;
  code: 
  | "UNSUPPORTED_FILE"
  | "INVALID_FILE"
  | "FILE_TOO_LARGE"
  | "ENCRYPTED_FILE"
  | "CONVERSION_FAILED"
}


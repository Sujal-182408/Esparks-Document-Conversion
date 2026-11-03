export type Format = "docx" | "pdf" | "pptx" | "md";

export const conversionOptions: Record<Format, Format[]> = {
  docx: ["md"],
  pdf: ["md"],
  pptx: ["md"],
  md: ["docx", "pdf", "pptx"],
};
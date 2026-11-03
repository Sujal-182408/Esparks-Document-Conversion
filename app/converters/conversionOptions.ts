import { ConversionOption } from "../types/conversion";
type Format = "docx" | "pdf" | "pptx" | "md";
const conversionOptions: Record<Format, Format[]> = {
  docx: ["pdf", "pptx", "md"],
  pdf: ["docx", "pptx", "md"],
  pptx: ["docx", "pdf", "md"],
  md: ["docx", "pdf", "pptx"],
};
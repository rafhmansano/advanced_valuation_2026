import { getDocument, GlobalWorkerOptions } from "pdfjs-dist/build/pdf.mjs";

// Use the bundled worker via URL import for Vite
GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url
).toString();

/**
 * Extract all text from a PDF file (ArrayBuffer).
 * Returns { pages: [{ pageNum, text }], fullText }
 */
export async function extractPdfText(arrayBuffer) {
  const pdf = await getDocument({ data: arrayBuffer }).promise;
  const pages = [];
  let fullText = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const text = content.items.map((item) => item.str).join(" ");
    pages.push({ pageNum: i, text });
    fullText += text + "\n";
  }

  return { pages, fullText, numPages: pdf.numPages };
}

// ─── KEYWORD PATTERNS PER SECTOR FIELD ───
// Each entry: { field key -> [regex patterns that match labels near the value] }
// Patterns are case-insensitive and look for a number after the keyword.

const NUMBER_RE = /[R$\\s]*(-?[\\d.,]+)/;

function extractNumber(text, patterns) {
  for (const pat of patterns) {
    const match = text.match(pat);
    if (match) {
      // Clean: remove dots as thousands sep, replace comma with dot
      let raw = match[1].replace(/\\.(?=\\d{3})/g, "").replace(",", ".");
      const val = parseFloat(raw);
      if (!isNaN(val)) return val;
    }
  }
  return null;
}

const FIELD_PATTERNS = {
  bancos: {
    bookValuePerShare: [
      /(?:VPA|valor\s+patrimonial\s+por\s+a[cç][aã]o)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
      /(?:patrimônio|patrimonial)\s*(?:\/|por)\s*(?:ação|a[cç][aã]o)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    roe: [
      /ROE\s*(?:recorrente|ajustado|anualizado)?[^\\d]*?(-?[\d.,]+)\s*%/i,
      /(?:retorno\s+sobre\s+(?:o\s+)?(?:patrimônio|PL))[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
    earningsPerShare: [
      /(?:LPA|lucro\s+por\s+a[cç][aã]o)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
  },
  transmissoras: {
    rapAnual: [
      /(?:RAP|receita\s+anual\s+permitida)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    opex: [
      /(?:OPEX|despesas?\s+operacionai?s?)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
  },
  geradoras: {
    ebitda: [
      /EBITDA\s*(?:ajustado|recorrente)?[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    netDebt: [
      /(?:d[ií]vida\s+l[ií]quida|endividamento\s+l[ií]quido)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    mwInstalled: [
      /(?:capacidade\s+instalada|MW\s+instalad[oa])[^\\d]*?(-?[\d.,]+)/i,
    ],
  },
  seguros: {
    bookValuePerShare: [
      /(?:VPA|valor\s+patrimonial)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    roe: [
      /ROE[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
    combinedRatio: [
      /(?:combined\s+ratio|[ií]ndice\s+combinad[oa])[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
  },
  saneamento: {
    ebitda: [
      /EBITDA[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    netDebt: [
      /(?:d[ií]vida\s+l[ií]quida)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
  },
  holdings: {
    totalInvestments: [
      /(?:total\s+(?:de\s+)?(?:participações|investimentos))[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    holdingDebt: [
      /(?:d[ií]vida\s+l[ií]quida)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
  },
  telecom: {
    ebitda: [
      /EBITDA[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    netDebt: [
      /(?:d[ií]vida\s+l[ií]quida)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
  },
  fii_logistico: {
    noiAnual: [
      /(?:NOI|receita\s+operacional\s+l[ií]quida)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    vpPerShare: [
      /(?:VP\s*(?:por|\/)\s*cota|valor\s+patrimonial\s*(?:por|\/)\s*cota)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
  },
  fii_shopping: {
    noiAnual: [
      /(?:NOI|receita\s+operacional\s+l[ií]quida)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    vpPerShare: [
      /(?:VP\s*(?:por|\/)\s*cota|valor\s+patrimonial\s*(?:por|\/)\s*cota)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    vendasM2: [
      /(?:vendas\s*(?:por|\/)\s*m[²2])[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    ablTotal: [
      /(?:ABL\s+total|[aá]rea\s+bruta\s+loc[aá]vel)[^\\d]*?(-?[\d.,]+)/i,
    ],
  },
  fii_lajes: {
    noiAnual: [
      /(?:NOI|receita\s+operacional\s+l[ií]quida)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    vacancyActual: [
      /(?:vac[aâ]ncia\s*(?:f[ií]sica|atual)?)[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
    vpPerShare: [
      /(?:VP\s*(?:por|\/)\s*cota|valor\s+patrimonial\s*(?:por|\/)\s*cota)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
  },
  fii_papel: {
    vpPerShare: [
      /(?:VP\s*(?:por|\/)\s*cota|valor\s+patrimonial\s*(?:por|\/)\s*cota)[^\\d]*?R?\$?\s*(-?[\d.,]+)/i,
    ],
    spreadMedio: [
      /(?:spread\s+m[eé]dio)[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
    ltvMedio: [
      /(?:LTV\s+m[eé]dio)[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
    inadimplencia: [
      /(?:inadimpl[eê]ncia)[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
    dividendYield: [
      /(?:dividend\s*yield|DY)\s*(?:12\s*m|anualizado)?[^\\d]*?(-?[\d.,]+)\s*%/i,
    ],
  },
};

// Common fields extracted regardless of sector
const COMMON_PATTERNS = {
  ticker: [
    /(?:ticker|c[oó]digo)[:\s]*([A-Z]{4}\d{1,2})/i,
    /\b([A-Z]{4}\d{1,2})\b/,
  ],
};

function extractTicker(text) {
  for (const pat of COMMON_PATTERNS.ticker) {
    const match = text.match(pat);
    if (match) return match[1].toUpperCase();
  }
  return null;
}

/**
 * Given full PDF text and a sector, attempt to extract field values.
 * Returns { extractedFields: { key: value }, confidence: { key: "high"|"low" } }
 */
export function extractFieldsFromText(fullText, sector) {
  const patterns = FIELD_PATTERNS[sector] || {};
  const extractedFields = {};
  const confidence = {};

  // Try ticker
  const ticker = extractTicker(fullText);
  if (ticker) {
    extractedFields.ticker = ticker;
    confidence.ticker = "high";
  }

  // Try sector-specific fields
  for (const [fieldKey, fieldPatterns] of Object.entries(patterns)) {
    const val = extractNumber(fullText, fieldPatterns);
    if (val !== null) {
      extractedFields[fieldKey] = val;
      confidence[fieldKey] = "high";
    }
  }

  return { extractedFields, confidence };
}

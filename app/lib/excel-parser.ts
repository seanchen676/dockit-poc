import ExcelJS from "exceljs";

export interface ParsedColumn {
  index: number;
  width?: number;
  hidden: boolean;
}

export interface ParsedImage {
  imageId: string;
  type?: string;
  range: string;
  base64: string | null;
}

export interface ParsedCell {
  address: string;
  value: unknown;
  type: string;
  style: Partial<ExcelJS.Style> | null;
  numFmt?: string;
}

export interface ParsedRow {
  index: number;
  cells: ParsedCell[];
}

export interface ParsedSheet {
  name: string;
  properties?: {
    dimension?: string;
    tabColor?: string;
  };
  views?: Partial<ExcelJS.WorksheetView>[];
  merges: string[];
  columns: ParsedColumn[];
  images: ParsedImage[];
  dataStruct: {
    rows: ParsedRow[];
  };
}

export interface ParsedResult {
  fileName: string;
  generatedAt: string;
  sheets: ParsedSheet[];
}

/**
 * Main parsing function
 */
export async function parseExcel(
  fileBuffer: ArrayBuffer,
  fileName: string,
): Promise<ParsedResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuffer);

  const result: ParsedResult = {
    fileName: fileName,
    generatedAt: new Date().toISOString(),
    sheets: [],
  };

  workbook.eachSheet((sheet) => {
    // 1. Basic Sheet Info
    const sheetData: ParsedSheet = {
      name: sheet.name,
      properties: {
        // dimensions IS actually exposed in recent exceljs versions or via prototype, but declared optionally.
        // If types complain, we access it "safely".
        dimension: sheet.dimensions?.toString() || undefined,
        tabColor: sheet.properties.tabColor?.argb || undefined,
      },
      views: sheet.views as Partial<ExcelJS.WorksheetView>[],
      merges: [],
      columns: [],
      images: [],
      dataStruct: { rows: [] },
    };

    // 2. Merges
    // ExcelJS stores merges in `_merges` (private) or we can access via other means.
    // However, the spec pseudo code used `sheet._merges`.
    // Typescript might complain about private properties. We'll try to cast or use a public getter if available.
    // It seems `sheet.model.merges` might be available or we iterate.
    // Let's stick to the spec's logic but handle TS.
    // Explicitly define a type for the internal structure we are accessing
    interface InternalSheet extends ExcelJS.Worksheet {
      _merges?: Record<string, { shortRange: string }>;
    }
    const internalSheet = sheet as unknown as InternalSheet;

    if (internalSheet._merges) {
      Object.keys(internalSheet._merges).forEach((key) => {
        if (internalSheet._merges && internalSheet._merges[key].shortRange) {
          sheetData.merges.push(internalSheet._merges[key].shortRange);
        }
      });
    }

    // 3. Columns
    if (sheet.columns) {
      sheet.columns.forEach((col, index) => {
        // ExcelJS columns can be objects or just sparse array entries
        // We safely access properties
        const width = col.width;
        const hidden = col.hidden;
        // Note: col.header might be present, but we prioritize row scanning
        sheetData.columns.push({
          index: index + 1,
          width: width,
          hidden: hidden || false,
        });
      });
    }

    // 4. Images
    // Type definition for what getImages actually returns (based on observation/error)
    interface SheetImage {
      type: "image";
      imageId: string;
      range: ExcelJS.ImageRange;
    }
    // We cast the return value because the official type definition might be lagging or strictly typed to something else
    const images = sheet.getImages() as unknown as SheetImage[];

    if (images && images.length > 0) {
      images.forEach((img) => {
        const imgId = img.imageId;
        const imgData = workbook.getImage(Number(imgId));

        // img.range is { tl: { col: 1.5, row: 1.5 }, br: ... }
        // We want integers or a clear format.
        // The spec suggested `${img.range.tl.nativeCol}:${img.range.tl.nativeRow}`
        // Let's protect against undefined
        let rangeStr = "";
        // Define a loose type for range anchor to access potential native properties safely
        interface NativeAnchor {
          nativeCol?: number;
          nativeRow?: number;
          col: number;
          row: number;
        }

        const tl = img.range.tl as NativeAnchor;

        if (tl) {
          rangeStr = `${tl.nativeCol ?? tl.col}:${tl.nativeRow ?? tl.row}`;
        }

        const base64Str = imgData.buffer
          ? arrayBufferToBase64(imgData.buffer)
          : null;

        sheetData.images.push({
          imageId: imgId,
          type: imgData.extension,
          range: rangeStr,
          base64: base64Str,
        });
      });
    }

    // 5. Universal Data Scan (Sparse)
    // sheet.eachRow({ includeEmpty: false }) skips empty rows
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const rowData: ParsedRow = {
        index: rowNumber,
        cells: [],
      };

      let hasContent = false;

      // Removed unused _colNumber argument
      row.eachCell({ includeEmpty: false }, (cell) => {
        // Updated: Capture all style properties without filtering "defaults"
        const cellStyle = extractCompleteStyle(cell.style);

        // Use ExcelJS.CellValue or a type that allows result to be any compatible value (but not any type)
        let rawValue:
          | ExcelJS.CellValue
          | { formula: string; result: ExcelJS.CellValue } = cell.value;
        let cellType = "string";

        // A. Formula
        // Check cell.formula directly as it handles Shared Formulas better than just cell.type
        // In ExcelJS, cell.formula is a getter that resolves the formula even in shared contexts if loaded correctly
        if (cell.formula) {
          cellType = "formula";
          rawValue = {
            formula: cell.formula,
            // result can be number, string, Date, error... unknown is safest if we want no-any
            result: cell.result as ExcelJS.CellValue,
          };
        } else if (cell.type === ExcelJS.ValueType.Formula) {
          // Fallback if cell.formula is empty but type says Formula (unlikely but safe)
          cellType = "formula";
          const val = cell.value as ExcelJS.CellFormulaValue;
          rawValue = {
            formula: val.formula,
            result: val.result as ExcelJS.CellValue,
          };
        }
        // B. Date
        else if (cell.type === ExcelJS.ValueType.Date) {
          cellType = "date";
          // cell.value should be Date object
          if (cell.value instanceof Date) {
            rawValue = cell.value.toISOString();
          }
        }
        // C. Rich Text
        else if (cell.type === ExcelJS.ValueType.RichText) {
          cellType = "richText";
          // rawValue is already the logical structure
        }
        // D. Number
        else if (cell.type === ExcelJS.ValueType.Number) {
          cellType = "number";
        }
        // E. Hyperlink
        else if (cell.type === ExcelJS.ValueType.Hyperlink) {
          cellType = "hyperlink";
          // cell.value might be { text: '...', hyperlink: '...' }
        }

        const cellObj: ParsedCell = {
          address: cell.address,
          value: rawValue,
          type: cellType,
          style: cellStyle,
          numFmt: cell.numFmt,
        };

        // Filter out completely empty useful info (though eachCell skips empty values, style might exist)
        if (rawValue !== null && rawValue !== undefined) {
          rowData.cells.push(cellObj);
          hasContent = true;
        } else if (cellStyle !== null) {
          // Even if value is empty, if style exists, we might want it (e.g. background color in empty cell)
          // But usually for LLM we care about data.
          // The spec says "Only extract ... to prevent output ... redundant data".
          // Let's keep it if style is present.
          rowData.cells.push(cellObj);
          hasContent = true;
        }
      });

      if (hasContent) {
        sheetData.dataStruct.rows.push(rowData);
      }
    });

    result.sheets.push(sheetData);
  });

  return result;
}

/**
 * Helper to convert ArrayBuffer/Buffer to Base64 string in Browser
 */
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return typeof btoa === "function"
    ? btoa(binary)
    : Buffer.from(binary, "binary").toString("base64");
}

/**
 * Extract complete style object without aggressive filtering
 * Captures all defined properties for Font, Fill, Border, Alignment, Protection.
 */
function extractCompleteStyle(style: Partial<ExcelJS.Style>) {
  if (!style) return null;
  const complete: Partial<ExcelJS.Style> = {};

  // 1. Font
  if (style.font) {
    // Copy all properties that are not undefined
    // We explicitly avoid filtering "Calibri" or "11" as per request
    const f = style.font;
    const mFont: Partial<ExcelJS.Font> = {};

    if (f.name) mFont.name = f.name;
    if (f.size) mFont.size = f.size;
    if (f.family) mFont.family = f.family;
    if (f.scheme) mFont.scheme = f.scheme;
    if (f.charset) mFont.charset = f.charset;
    if (f.color) mFont.color = f.color;
    if (f.bold) mFont.bold = true;
    if (f.italic) mFont.italic = true;
    if (f.underline) mFont.underline = f.underline; // can be boolean or string ('double', etc)
    if (f.strike) mFont.strike = true;
    if (f.outline) mFont.outline = true;
    if (f.vertAlign) mFont.vertAlign = f.vertAlign;

    if (Object.keys(mFont).length > 0) complete.font = mFont;
  }

  // 2. Fill
  if (style.fill) {
    // Capture all fill types (pattern, gradient)
    complete.fill = style.fill;
  }

  // 3. Alignment
  if (style.alignment) {
    complete.alignment = style.alignment;
  }

  // 4. Border
  if (style.border) {
    // Capture full border object
    complete.border = style.border;
  }

  // 5. Protection
  if (style.protection) {
    complete.protection = style.protection;
  }

  return Object.keys(complete).length > 0 ? complete : null;
}

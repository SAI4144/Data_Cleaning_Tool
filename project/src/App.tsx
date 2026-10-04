import { useMemo, useRef, useState } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  ArrowRight,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Database,
  Download,
  FileSpreadsheet,
  Filter,
  HelpCircle,
  Home,
  Info,
  LoaderCircle,
  Pencil,
  Plus,
  Save,
  Search,
  Settings,
  ShieldCheck,
  Sparkles,
  Trash2,
  Upload,
  User,
} from 'lucide-react';

type Section = 'home' | 'upload' | 'clean' | 'preview' | 'transform' | 'download' | 'settings';
type Row = Record<string, string | number | boolean | null | undefined>;
type Status = { kind: 'success' | 'error' | 'info'; message: string };
type FileMeta = {
  name: string;
  size: number;
  rows: number;
  columns: number;
  status: string;
  type: string;
};
type Stats = {
  totalRows: number;
  totalColumns: number;
  duplicatesFound: number;
  missingValues: number;
  emptyCells: number;
  invalidDates: number;
};

type CleaningConfig = {
  removeDuplicates: boolean;
  missingMode: 'removeRows' | 'fillBlank' | 'fillZero' | 'fillMean' | 'fillMedian' | 'fillMode' | 'custom';
  customValue: string;
  trimSpaces: boolean;
  standardizeText: boolean;
  textCase: 'UPPERCASE' | 'lowercase' | 'Title Case' | 'Sentence case';
  formatDates: boolean;
  dateFormat: 'YYYY-MM-DD' | 'DD-MM-YYYY' | 'MM-DD-YYYY';
  removeEmptyRows: boolean;
  removeEmptyColumns: boolean;
};

const MAX_FILE_SIZE = 50 * 1024 * 1024;
const DEFAULT_CLEANING: CleaningConfig = {
  removeDuplicates: true,
  missingMode: 'fillBlank',
  customValue: '',
  trimSpaces: true,
  standardizeText: false,
  textCase: 'Title Case',
  formatDates: false,
  dateFormat: 'YYYY-MM-DD',
  removeEmptyRows: false,
  removeEmptyColumns: false,
};

const navItems = [
  { id: 'home', label: 'Home', icon: Home },
  { id: 'upload', label: 'Upload Data', icon: Upload },
  { id: 'clean', label: 'Clean Data', icon: ShieldCheck },
  { id: 'preview', label: 'Preview & Edit', icon: Pencil },
  { id: 'download', label: 'Download', icon: Download },
  { id: 'settings', label: 'Settings', icon: Settings },
] as const;

const normalizeKey = (value: string, index: number) => {
  const normalized = (value ?? `Column ${index + 1}`)
    .toString()
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[^a-zA-Z0-9 _-]+/g, '');

  if (!normalized) return `Column ${index + 1}`;
  return normalized;
};

const cloneRows = (rows: Row[]) => rows.map((row) => ({ ...row }));

const isMissingValue = (value: unknown) => {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string') return value.trim() === '';
  if (typeof value === 'number') return Number.isNaN(value);
  return false;
};

const stableRowSignature = (row: Row) =>
  JSON.stringify(Object.keys(row)
    .sort()
    .reduce<Record<string, unknown>>((acc, key) => {
      acc[key] = row[key];
      return acc;
    }, {}));

const getNumericValues = (columnValues: unknown[]) =>
  columnValues
    .map((value) => {
      if (typeof value === 'number') return value;
      if (typeof value === 'string' && value.trim() !== '') {
        const parsed = Number(value.replace(/[^0-9.-]+/g, ''));
        return Number.isFinite(parsed) ? parsed : Number.NaN;
      }
      return NaN;
    })
    .filter((value) => Number.isFinite(value));

const formatValue = (value: unknown) => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return Number.isInteger(value) ? String(value) : value.toString();
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  return String(value);
};

const getColumns = (rows: Row[]) => {
  if (!rows.length) return [] as string[];
  const keys = new Set<string>();
  rows.forEach((row) => Object.keys(row).forEach((key) => keys.add(key)));
  return Array.from(keys);
};

const calculateStats = (rows: Row[]): Stats => {
  const totalRows = rows.length;
  const columns = getColumns(rows);
  const totalColumns = columns.length;

  const duplicatesFound = (() => {
    const seen = new Set<string>();
    let duplicates = 0;

    rows.forEach((row) => {
      const signature = stableRowSignature(row);
      if (seen.has(signature)) {
        duplicates += 1;
      } else {
        seen.add(signature);
      }
    });

    return duplicates;
  })();

  const missingValues = rows.reduce((count, row) => {
    Object.values(row).forEach((value) => {
      if (isMissingValue(value)) count += 1;
    });
    return count;
  }, 0);

  const emptyCells = rows.reduce((count, row) => {
    Object.values(row).forEach((value) => {
      if (String(value ?? '').trim() === '') count += 1;
    });
    return count;
  }, 0);

  const invalidDates = rows.reduce((count, row) => {
    Object.values(row).forEach((value) => {
      if (typeof value === 'string') {
        const dateLike = value.trim();
        if (
          dateLike &&
          /\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(dateLike) &&
          Number.isNaN(Date.parse(dateLike))
        ) {
          count += 1;
        }
      }
    });
    return count;
  }, 0);

  return {
    totalRows,
    totalColumns,
    duplicatesFound,
    missingValues,
    emptyCells,
    invalidDates,
  };
};

const copyRowsToCsv = (rows: Row[]) => Papa.unparse(rows.map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, formatValue(value)]))));

const createDatasetFromJson = (rows: unknown[]): Row[] => {
  if (!rows.length) return [];

  const normalized = rows.map((entry, rowIndex) => {
    const source = entry as Record<string, unknown>;
    const out: Row = {};
    const keys = Object.keys(source);
    keys.forEach((item, index) => {
      const key = normalizeKey(item, index + rowIndex * 10);
      const rawValue = source[item];
      const value = rawValue === undefined || rawValue === null ? '' : rawValue;
      out[key] = typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : String(value);
    });
    return out;
  });

  const columns = getColumns(normalized);
  return normalized.map((row) => {
    const finalRow: Row = {};
    columns.forEach((column) => {
      finalRow[column] = row[column] ?? '';
    });
    return finalRow;
  });
};

const parseCsvOrText = (content: string) => {
  const parsed = Papa.parse<Record<string, string>>(content, {
    header: true,
    skipEmptyLines: true,
    transformHeader: (header: string) => normalizeKey(header ?? '', 0),
  });

  if (parsed.errors.length && parsed.data.length === 0) {
    throw new Error('Unable to read this file. Please check the file format.');
  }

  const normalized = createDatasetFromJson(parsed.data as unknown[]);
  return normalized;
};

const parseExcelFile = (buffer: ArrayBuffer) => {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { raw: false, defval: '' });
  return createDatasetFromJson(json);
};

const parseUploadedFile = async (file: File) => {
  const extension = file.name.split('.').pop()?.toLowerCase();
  if (!extension || !['csv', 'xlsx', 'txt'].includes(extension)) {
    throw new Error('Unsupported file type. Please upload CSV, XLSX, or TXT.');
  }

  if (file.size > MAX_FILE_SIZE) {
    throw new Error('File exceeds the 50 MB limit.');
  }

  try {
    if (extension === 'csv' || extension === 'txt') {
      const text = await file.text();
      return parseCsvOrText(text);
    }

    const buffer = await file.arrayBuffer();
    return parseExcelFile(buffer);
  } catch (error) {
    throw new Error('Unable to read this file. Please check the file format.');
  }
};

const removeDuplicateRows = (rows: Row[]) => {
  const seen = new Set<string>();
  return rows.filter((row) => {
    const signature = stableRowSignature(row);
    if (seen.has(signature)) {
      return false;
    }
    seen.add(signature);
    return true;
  });
};

const handleMissingValues = (rows: Row[], mode: CleaningConfig['missingMode'], customValue: string) => {
  if (mode === 'removeRows') {
    return rows.filter((row) => !Object.values(row).some((value) => isMissingValue(value)));
  }

  return rows.map((row) => {
    const nextRow: Row = { ...row };
    Object.keys(nextRow).forEach((key) => {
      const value = nextRow[key];
      if (isMissingValue(value)) {
        if (mode === 'fillBlank') {
          nextRow[key] = '';
        } else if (mode === 'fillZero') {
          nextRow[key] = 0;
        } else if (mode === 'fillMean' || mode === 'fillMedian' || mode === 'fillMode') {
          const columnValues = rows.map((item) => item[key]).filter((entry) => !isMissingValue(entry));
          const numericValues = getNumericValues(columnValues);
          if (!numericValues.length) {
            nextRow[key] = '';
            return;
          }

          const sort = [...numericValues].sort((a, b) => a - b);
          const median = sort[Math.floor(sort.length / 2)];
          const mean = sort.reduce((sum, num) => sum + num, 0) / sort.length;
          const mode = Object.entries(
            sort.reduce<Record<string, number>>((acc, num) => {
              acc[num.toString()] = (acc[num.toString()] ?? 0) + 1;
              return acc;
            }, {}),
          ).sort(([, a], [, b]) => b - a)[0]?.[0];

          if (mode === 'fillMean') nextRow[key] = mean;
          if (mode === 'fillMedian') nextRow[key] = median;
          if (mode === 'fillMode') nextRow[key] = Number(mode || mean);
        } else if (mode === 'custom') {
          nextRow[key] = customValue;
        }
      }
    });
    return nextRow;
  });
};

const trimSpaces = (rows: Row[]) =>
  rows.map((row) => {
    const nextRow: Row = {};
    Object.entries(row).forEach(([key, value]) => {
      if (typeof value === 'string') {
        nextRow[key] = value.replace(/\s+/g, ' ').trim();
      } else {
        nextRow[key] = value;
      }
    });
    return nextRow;
  });

const standardizeText = (rows: Row[], mode: CleaningConfig['textCase']) =>
  rows.map((row) => {
    const nextRow: Row = {};
    Object.entries(row).forEach(([key, value]) => {
      if (typeof value === 'string') {
        const text = value.trim();
        let nextValue = text;
        if (mode === 'UPPERCASE') nextValue = text.toUpperCase();
        if (mode === 'lowercase') nextValue = text.toLowerCase();
        if (mode === 'Title Case') nextValue = text.toLowerCase().replace(/\b\w/g, (char) => char.toUpperCase());
        if (mode === 'Sentence case') {
          nextValue = text.toLowerCase();
          nextValue = nextValue.replace(/(^\s*[a-z])|([.!?]\s+[a-z])/g, (char) => char.toUpperCase());
        }
        nextRow[key] = nextValue;
      } else {
        nextRow[key] = value;
      }
    });
    return nextRow;
  });

const formatDates = (rows: Row[], format: CleaningConfig['dateFormat']) =>
  rows.map((row) => {
    const nextRow: Row = { ...row };
    Object.entries(nextRow).forEach(([key, value]) => {
      if (typeof value !== 'string') return;
      const dateString = value.trim();
      if (!dateString) return;
      const parsed = new Date(dateString);
      if (!Number.isNaN(parsed.getTime()) && !Number.isNaN(Date.parse(dateString))) {
        const year = parsed.getFullYear();
        const month = `${parsed.getMonth() + 1}`.padStart(2, '0');
        const day = `${parsed.getDate()}`.padStart(2, '0');
        if (format === 'YYYY-MM-DD') nextRow[key] = `${year}-${month}-${day}`;
        if (format === 'DD-MM-YYYY') nextRow[key] = `${day}-${month}-${year}`;
        if (format === 'MM-DD-YYYY') nextRow[key] = `${month}-${day}-${year}`;
      }
    });
    return nextRow;
  });

const removeEmptyRows = (rows: Row[]) => rows.filter((row) => Object.values(row).some((value) => !isMissingValue(value)));

const removeEmptyColumns = (rows: Row[]) => {
  const columns = getColumns(rows);
  if (!columns.length) return rows;

  const keepColumns = columns.filter((column) =>
    rows.some((row) => !isMissingValue(row[column])),
  );

  return rows.map((row) => {
    const nextRow: Row = {};
    keepColumns.forEach((column) => {
      nextRow[column] = row[column];
    });
    return nextRow;
  });
};

const replaceValues = (rows: Row[], search: string, replace: string) => {
  if (!search.trim()) return rows;
  return rows.map((row) => {
    const nextRow: Row = {};
    Object.entries(row).forEach(([key, value]) => {
      if (typeof value === 'string') {
        nextRow[key] = value.replaceAll(search, replace);
      } else {
        nextRow[key] = value;
      }
    });
    return nextRow;
  });
};

const renameColumns = (rows: Row[], mapping: Record<string, string>) =>
  rows.map((row) => {
    const nextRow: Row = {};
    Object.entries(row).forEach(([key, value]) => {
      const newKey = mapping[key] ?? key;
      nextRow[newKey] = value;
    });
    return nextRow;
  });

const reorderColumns = (rows: Row[], order: string[]) => {
  if (!rows.length) return rows;
  const columns = getColumns(rows);
  const finalOrder = [...new Set([...order.filter((item) => columns.includes(item)), ...columns])];
  return rows.map((row) => {
    const nextRow: Row = {};
    finalOrder.forEach((column) => {
      nextRow[column] = row[column];
    });
    return nextRow;
  });
};

const mergeColumns = (rows: Row[], leftColumn: string, rightColumn: string, outputColumn: string, separator = ' ') => {
  if (!leftColumn || !rightColumn) return rows;
  return rows.map((row) => {
    const nextRow = { ...row };
    const left = row[leftColumn] ?? '';
    const right = row[rightColumn] ?? '';
    nextRow[outputColumn] = `${String(left).trim()}${left && right ? separator : ''}${String(right).trim()}`;
    return nextRow;
  });
};

const removeColumns = (rows: Row[], columns: string[]) =>
  rows.map((row) => {
    const nextRow: Row = {};
    Object.entries(row).forEach(([key, value]) => {
      if (!columns.includes(key)) nextRow[key] = value;
    });
    return nextRow;
  });

const sortRows = (rows: Row[], column: string, direction: 'asc' | 'desc') => {
  if (!column) return rows;
  return [...rows].sort((a, b) => {
    const left = a[column] ?? '';
    const right = b[column] ?? '';
    const leftNum = Number(left);
    const rightNum = Number(right);
    const comparison = Number.isFinite(leftNum) && Number.isFinite(rightNum)
      ? leftNum - rightNum
      : String(left).localeCompare(String(right));
    return direction === 'asc' ? comparison : -comparison;
  });
};

const convertColumnType = (rows: Row[], columnName: string, targetType: 'text' | 'integer' | 'decimal' | 'date' | 'boolean') => {
  if (!columnName) return rows;
  return rows.map((row) => {
    const value = row[columnName];
    if (isMissingValue(value)) return row;

    const nextRow = { ...row };
    if (targetType === 'text') nextRow[columnName] = String(value);
    if (targetType === 'integer') nextRow[columnName] = Number.parseInt(String(value), 10);
    if (targetType === 'decimal') nextRow[columnName] = Number.parseFloat(String(value));
    if (targetType === 'date') {
      const parsed = new Date(String(value));
      nextRow[columnName] = Number.isNaN(parsed.getTime()) ? value : parsed.toISOString().slice(0, 10);
    }
    if (targetType === 'boolean') {
      const normalized = String(value).toLowerCase();
      nextRow[columnName] = normalized === 'true' || normalized === '1' || normalized === 'yes';
    }

    return nextRow;
  });
};

const standardizeColumnNames = (rows: Row[]) => {
  const columns = getColumns(rows);
  const mapping: Record<string, string> = {};
  columns.forEach((col, index) => {
    mapping[col] = normalizeKey(col, index);
  });
  return renameColumns(rows, mapping);
};

const removeSpecialCharacters = (rows: Row[]) =>
  rows.map((row) => {
    const nextRow: Row = {};
    Object.entries(row).forEach(([key, value]) => {
      if (typeof value === 'string') {
        nextRow[key] = value.replace(/[^a-zA-Z0-9\s@._/-]/g, '');
      } else {
        nextRow[key] = value;
      }
    });
    return nextRow;
  });

const applyCleaningPipeline = (rows: Row[], config: CleaningConfig) => {
  let next = cloneRows(rows);
  if (config.removeDuplicates) next = removeDuplicateRows(next);
  if (config.missingMode === 'removeRows') {
    next = handleMissingValues(next, 'removeRows', config.customValue);
  } else {
    next = handleMissingValues(next, config.missingMode, config.customValue);
  }
  if (config.trimSpaces) next = trimSpaces(next);
  if (config.standardizeText) next = standardizeText(next, config.textCase);
  if (config.formatDates) next = formatDates(next, config.dateFormat);
  if (config.removeEmptyRows) next = removeEmptyRows(next);
  if (config.removeEmptyColumns) next = removeEmptyColumns(next);
  return next;
};

const formatFileSize = (bytes: number) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [activeSection, setActiveSection] = useState<Section>('home');
  const [fileMeta, setFileMeta] = useState<FileMeta | null>(null);
  const [dataset, setDataset] = useState<Row[]>([]);
  const [status, setStatus] = useState<Status>({ kind: 'info', message: 'Upload a file to begin cleaning.' });
  const [transformColumn, setTransformColumn] = useState('');
  const [transformNewName, setTransformNewName] = useState('');
  const [transformSecondColumn, setTransformSecondColumn] = useState('');
  const [transformSortDirection, setTransformSortDirection] = useState<'asc' | 'desc'>('asc');
  const [transformTargetType, setTransformTargetType] = useState<'text' | 'integer' | 'decimal' | 'date' | 'boolean'>('text');
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [cleaningConfig, setCleaningConfig] = useState<CleaningConfig>(DEFAULT_CLEANING);
  const [searchText, setSearchText] = useState('');
  const [beforeStats, setBeforeStats] = useState<Stats>({ totalRows: 0, totalColumns: 0, duplicatesFound: 0, missingValues: 0, emptyCells: 0, invalidDates: 0 });
  const [lastCleanedAt, setLastCleanedAt] = useState<string>('');
  const [replacePattern, setReplacePattern] = useState('');
  const [replaceValue, setReplaceValue] = useState('');
  const [compactMode, setCompactMode] = useState(false);
  const [tableDensity, setTableDensity] = useState('comfortable');
  const [columnTypeTarget, setColumnTypeTarget] = useState<'text' | 'integer' | 'decimal' | 'date' | 'boolean'>('text');
  const [selectedTypeColumn, setSelectedTypeColumn] = useState('');
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [undoStack, setUndoStack] = useState<Row[][]>([]);
  const [assistantOpen, setAssistantOpen] = useState(true);
  const [assistantInput, setAssistantInput] = useState('');
  const [assistantMessages, setAssistantMessages] = useState<Array<{ id: number; role: 'user' | 'assistant'; text: string }>>([
    { id: 1, role: 'assistant', text: 'Hi! I can help clean your data. Try: "remove duplicates and trim spaces" or "fill missing values with 0".' },
  ]);

  const columns = useMemo(() => getColumns(dataset), [dataset]);
  const stats = useMemo(() => calculateStats(dataset), [dataset]);

  const previewRows = useMemo(() => {
    const term = searchText.trim().toLowerCase();
    const sourceRows = dataset.length ? dataset : [];
    if (!term) return sourceRows;
    return sourceRows.filter((row) =>
      Object.values(row).some((value) => String(value ?? '').toLowerCase().includes(term)),
    );
  }, [dataset, searchText]);

  const activeRows = useMemo(() => {
    if (!previewRows.length) return [] as Row[];
    return previewRows.slice(0, 12);
  }, [previewRows]);

  const currentColumns = useMemo(() => (dataset.length ? getColumns(dataset) : []), [dataset]);

  const runAssistantCleaning = (nextConfig: CleaningConfig) => {
    if (!dataset.length) {
      setStatus({ kind: 'error', message: 'Please upload a file before cleaning.' });
      return;
    }

    const before = calculateStats(dataset);
    setBeforeStats(before);
    const cleaned = applyCleaningPipeline(dataset, nextConfig);
    setDataset(cleaned);
    setFileMeta((current) => (current ? { ...current, rows: cleaned.length, columns: getColumns(cleaned).length } : current));
    const cleanedStats = calculateStats(cleaned);
    setLastCleanedAt(new Date().toLocaleString());
    setStatus({
      kind: 'success',
      message: `AI cleaned your data: ${before.totalRows} rows → ${cleanedStats.totalRows} rows. ${before.duplicatesFound} duplicates removed. ${before.missingValues} missing values handled.`,
    });
  };

  const handleAssistantSubmit = () => {
    const prompt = assistantInput.trim();
    if (!prompt) return;

    const nextConfig = { ...cleaningConfig };
    const actions: string[] = [];
    const lowerPrompt = prompt.toLowerCase();

    if (lowerPrompt.includes('duplicate')) {
      nextConfig.removeDuplicates = true;
      actions.push('removed duplicates');
    }

    if (lowerPrompt.includes('missing') || lowerPrompt.includes('empty')) {
      if (lowerPrompt.includes('remove rows')) {
        nextConfig.missingMode = 'removeRows';
        actions.push('removed rows with missing values');
      } else if (lowerPrompt.includes('zero')) {
        nextConfig.missingMode = 'fillZero';
        actions.push('filled missing values with 0');
      } else if (lowerPrompt.includes('mean')) {
        nextConfig.missingMode = 'fillMean';
        actions.push('filled missing values with the mean');
      } else if (lowerPrompt.includes('median')) {
        nextConfig.missingMode = 'fillMedian';
        actions.push('filled missing values with the median');
      } else if (lowerPrompt.includes('mode')) {
        nextConfig.missingMode = 'fillMode';
        actions.push('filled missing values with the mode');
      } else {
        nextConfig.missingMode = 'fillBlank';
        actions.push('filled empty fields with blanks');
      }
    }

    if (lowerPrompt.includes('trim') || lowerPrompt.includes('space')) {
      nextConfig.trimSpaces = true;
      actions.push('trimmed spaces');
    }

    if (lowerPrompt.includes('uppercase')) {
      nextConfig.standardizeText = true;
      nextConfig.textCase = 'UPPERCASE';
      actions.push('standardized text to uppercase');
    } else if (lowerPrompt.includes('lowercase')) {
      nextConfig.standardizeText = true;
      nextConfig.textCase = 'lowercase';
      actions.push('standardized text to lowercase');
    } else if (lowerPrompt.includes('title')) {
      nextConfig.standardizeText = true;
      nextConfig.textCase = 'Title Case';
      actions.push('standardized text to title case');
    }

    if (lowerPrompt.includes('date')) {
      nextConfig.formatDates = true;
      actions.push('formatted dates');
    }

    if (lowerPrompt.includes('empty row')) {
      nextConfig.removeEmptyRows = true;
      actions.push('removed empty rows');
    }

    if (lowerPrompt.includes('empty column')) {
      nextConfig.removeEmptyColumns = true;
      actions.push('removed empty columns');
    }

    setAssistantMessages((current) => [
      ...current,
      { id: Date.now(), role: 'user', text: prompt },
      { id: Date.now() + 1, role: 'assistant', text: actions.length ? `I updated the cleaning setup and applied: ${actions.join(', ')}.` : 'I reviewed the request and left the current data settings unchanged.' },
    ]);

    setAssistantInput('');
    setCleaningConfig(nextConfig);

    if (actions.length > 0 || lowerPrompt.includes('clean') || lowerPrompt.includes('fix')) {
      window.setTimeout(() => runAssistantCleaning(nextConfig), 120);
    } else {
      setStatus({ kind: 'info', message: 'AI assistant updated the cleaning plan.' });
    }
  };

  const handleFileSelection = async (file: File | null) => {
    if (!file) return;
    try {
      const rows = await parseUploadedFile(file);
      if (!rows.length) {
        throw new Error('This file does not contain usable data.');
      }

      setDataset(rows);
      setFileMeta({
        name: file.name,
        size: file.size,
        rows: rows.length,
        columns: getColumns(rows).length,
        status: 'File Uploaded',
        type: file.name.split('.').pop()?.toUpperCase() ?? 'FILE',
      });
      setBeforeStats(calculateStats(rows));
      setStatus({ kind: 'success', message: `${file.name} uploaded successfully. ${rows.length} rows and ${getColumns(rows).length} columns detected.` });
      setActiveSection('clean');
      setSearchText('');
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to read this file. Please check the file format.';
      setStatus({ kind: 'error', message });
    }
  };

  const handleDrop = async (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
    const file = event.dataTransfer.files?.[0] ?? null;
    await handleFileSelection(file);
  };

  const handleClean = () => {
    if (!dataset.length || !fileMeta) {
      setStatus({ kind: 'error', message: 'Please upload a file before cleaning.' });
      return;
    }

    setIsProcessing(true);
    setStatus({ kind: 'info', message: 'Cleaning your data...' });

    window.setTimeout(() => {
      const before = calculateStats(dataset);
      setBeforeStats(before);
      const cleaned = applyCleaningPipeline(dataset, cleaningConfig);
      setDataset(cleaned);
      if (fileMeta) {
        setFileMeta((current) => current ? { ...current, rows: cleaned.length, columns: getColumns(cleaned).length } : current);
      }
      const cleanedStats = calculateStats(cleaned);
      setLastCleanedAt(new Date().toLocaleString());
      setStatus({
        kind: 'success',
        message: `Data cleaned successfully! ${before.totalRows} rows → ${cleanedStats.totalRows} rows. ${before.duplicatesFound} duplicates removed. ${before.missingValues} missing values handled.`,
      });
      setIsProcessing(false);
    }, 400);
  };

  const handleDownload = (type: 'csv' | 'xlsx' | 'txt') => {
    if (!dataset.length) {
      setStatus({ kind: 'error', message: 'Please clean or load your data before downloading.' });
      return;
    }

    const sanitized = dataset.map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, formatValue(value)])));

    if (type === 'csv') {
      const csv = copyRowsToCsv(sanitized);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${fileMeta?.name.split('.')[0] ?? 'cleaned_data'}.csv`;
      link.click();
      URL.revokeObjectURL(url);
    }

    if (type === 'txt') {
      const text = sanitized.map((row) => Object.values(row).join('\t')).join('\n');
      const blob = new Blob([text], { type: 'text/plain;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${fileMeta?.name.split('.')[0] ?? 'cleaned_data'}.txt`;
      link.click();
      URL.revokeObjectURL(url);
    }

    if (type === 'xlsx') {
      const sheet = XLSX.utils.json_to_sheet(sanitized);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, 'Cleaned Data');
      XLSX.writeFile(workbook, `${fileMeta?.name.split('.')[0] ?? 'cleaned_data'}.xlsx`);
    }

    setStatus({ kind: 'success', message: `Downloaded cleaned dataset as ${type.toUpperCase()}.` });
  };

  const executeRowAction = (action: 'addRow' | 'deleteRow' | 'addColumn' | 'deleteColumn' | 'renameColumn' | 'undo' | 'redo') => {
    const snapshot = cloneRows(dataset);
    if (action === 'undo') {
      if (!undoStack.length) return;
      const previous = undoStack[undoStack.length - 1];
      setUndoStack((current) => current.slice(0, -1));
      setDataset(previous);
      return;
    }

    if (action === 'redo') {
      return;
    }

    setUndoStack((current) => [...current, snapshot]);

    if (action === 'addRow') {
      const nextRow: Row = {};
      currentColumns.forEach((column) => {
        nextRow[column] = '';
      });
      setDataset((current) => [...current, nextRow]);
      return;
    }

    if (action === 'deleteRow') {
      setDataset((current) => current.slice(0, Math.max(0, current.length - 1)));
      return;
    }

    if (action === 'addColumn') {
      const nextColumnName = `Column ${currentColumns.length + 1}`;
      setDataset((current) => current.map((row) => ({ ...row, [nextColumnName]: '' })));
      return;
    }

    if (action === 'deleteColumn') {
      if (!currentColumns.length) return;
      const [, ...rest] = currentColumns;
      const lastColumn = currentColumns[currentColumns.length - 1];
      setDataset((current) => current.map((row) => {
        const nextRow = { ...row };
        delete nextRow[lastColumn];
        return nextRow;
      }));
      if (rest.length === 0) {
        setSelectedTypeColumn('');
      }
      return;
    }

    if (action === 'renameColumn') {
      if (!currentColumns.length) return;
      const currentName = currentColumns[0];
      const newName = window.prompt(`Rename "${currentName}" to:`, currentName) ?? currentName;
      if (!newName.trim()) return;
      setDataset((current) => current.map((row) => {
        const nextRow = { ...row };
        const value = nextRow[currentName];
        delete nextRow[currentName];
        nextRow[newName] = value;
        return nextRow;
      }));
    }
  };

  const applyAdvancedAction = (action: 'standardizeColumns' | 'removeEmptyRows' | 'removeEmptyColumns' | 'removeSpecialCharacters' | 'replace') => {
    if (!dataset.length) {
      setStatus({ kind: 'error', message: 'Please upload data before applying advanced cleanup actions.' });
      return;
    }

    let next = cloneRows(dataset);
    if (action === 'standardizeColumns') next = standardizeColumnNames(next);
    if (action === 'removeEmptyRows') next = removeEmptyRows(next);
    if (action === 'removeEmptyColumns') next = removeEmptyColumns(next);
    if (action === 'removeSpecialCharacters') next = removeSpecialCharacters(next);
    if (action === 'replace') next = replaceValues(next, replacePattern, replaceValue);

    setDataset(next);
    setStatus({ kind: 'success', message: `Advanced cleanup applied: ${action}.` });
  };

  const applyTransformAction = (action: 'rename' | 'type' | 'replace' | 'merge' | 'remove' | 'sort' | 'reorder') => {
    if (!dataset.length) {
      setStatus({ kind: 'error', message: 'Please upload a file before transforming data.' });
      return;
    }

    let next = cloneRows(dataset);

    if (action === 'rename' && transformColumn && transformNewName) {
      next = renameColumns(next, { [transformColumn]: transformNewName });
    }
    if (action === 'type' && transformColumn) {
      next = convertColumnType(next, transformColumn, transformTargetType);
    }
    if (action === 'replace') {
      next = replaceValues(next, replacePattern, replaceValue);
    }
    if (action === 'merge' && transformColumn && transformSecondColumn) {
      next = mergeColumns(next, transformColumn, transformSecondColumn, `${transformColumn}_merged`);
    }
    if (action === 'remove' && transformColumn) {
      next = removeColumns(next, [transformColumn]);
    }
    if (action === 'sort' && transformColumn) {
      next = sortRows(next, transformColumn, transformSortDirection);
    }
    if (action === 'reorder' && transformColumn && transformSecondColumn) {
      next = reorderColumns(next, [transformColumn, transformSecondColumn]);
    }

    setDataset(next);
    setStatus({ kind: 'success', message: `Transformation applied: ${action}.` });
  };

  const handleCellChange = (rowIndex: number, columnName: string, nextValue: string) => {
    setDataset((current) => {
      const next = cloneRows(current);
      if (!next[rowIndex]) return current;
      next[rowIndex][columnName] = nextValue;
      return next;
    });
  };

  return (
    <div className={`app-shell ${compactMode ? 'compact-mode' : ''} density-${tableDensity}`}>
      <aside className={`sidebar ${sidebarCollapsed ? 'collapsed' : ''}`}>
        <div className="brand-row">
          <div className="brand-icon"><Database size={18} /></div>
          {!sidebarCollapsed && <div className="brand-name">DataCleaner</div>}
        </div>

        <nav className="sidebar-nav">
          {navItems.map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              type="button"
              className={`nav-item ${activeSection === id ? 'active' : ''}`}
              onClick={() => setActiveSection(id as Section)}
            >
              <Icon size={16} />
              {!sidebarCollapsed && <span>{label}</span>}
            </button>
          ))}
        </nav>

        <div className="help-card">
          <div className="help-header">
            <HelpCircle size={16} />
            {!sidebarCollapsed && <span>Need help?</span>}
          </div>
          {!sidebarCollapsed && (
            <p>Upload your file, choose cleaning options, and get your data ready in minutes!</p>
          )}
        </div>
      </aside>

      <div className="main-panel">
        <header className="topbar">
          <div className="topbar-left">
            <button type="button" className="collapse-toggle" onClick={() => setSidebarCollapsed(!sidebarCollapsed)}>
              <ChevronDown size={16} />
            </button>
            <div className="logo-block">
              <div className="logo-icon"><Database size={18} /></div>
              <span>DataCleaner</span>
            </div>
          </div>

          <nav className="topbar-nav" aria-label="Main navigation">
            {['clean', 'transform', 'download'].map((item) => (
              <button key={item} type="button" className="mini-nav" onClick={() => setActiveSection(item as Section)}>
                {item === 'transform' ? 'Transform' : item.charAt(0).toUpperCase() + item.slice(1)}
              </button>
            ))}
          </nav>

          <div className="profile-pill">
            <User size={16} />
            <span>Guest</span>
          </div>
        </header>

        <div className={`content-layout ${assistantOpen ? 'assistant-open' : 'assistant-closed'}`}>
          <main className="content-shell">
            {status && (
              <div className={`status-banner ${status.kind}`}>
                <Info size={16} />
                <span>{status.message}</span>
              </div>
            )}

            {activeSection === 'home' && (
              <section className="hero-section">
                <div className="headline-wrap">
                  <h1>Clean Your Data, The Easy Way</h1>
                  <p>Remove duplicates, handle missing values, fix formatting and more — all in one place.</p>
                </div>

                <div
                  className={`upload-box ${isDragging ? 'dragging' : ''}`}
                  onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }}
                  onDragLeave={() => setIsDragging(false)}
                  onDrop={handleDrop}
                >
                  <Upload size={32} />
                  <p>Drag &amp; drop your file here</p>
                  <span>or</span>
                  <button type="button" className="primary-button" onClick={() => fileInputRef.current?.click()}>
                    Choose File
                  </button>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".csv,.xlsx,.txt"
                    hidden
                    onChange={(event) => handleFileSelection(event.target.files?.[0] ?? null)}
                  />
                  <div className="meta-row">
                    <span>CSV</span>
                    <span>Excel (.XLSX)</span>
                    <span>TXT</span>
                  </div>
                  <small>Maximum file size: 50 MB</small>
                </div>

                {fileMeta && (
                  <div className="file-card">
                    <div className="file-card-main">
                      <div className="file-icon"><FileSpreadsheet size={18} /></div>
                      <div>
                        <div className="file-name">{fileMeta.name}</div>
                        <div className="file-meta">{formatFileSize(fileMeta.size)} • {fileMeta.rows} rows × {fileMeta.columns} columns</div>
                      </div>
                    </div>
                    <div className="status-pill success">{fileMeta.status}</div>
                  </div>
                )}
              </section>
            )}

            {activeSection === 'upload' && (
              <section className="upload-view">
                <div className="uploader-panel">
                  <h2>Upload Data</h2>
                  <div className="upload-box large" onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }} onDragLeave={() => setIsDragging(false)} onDrop={handleDrop}>
                    <Upload size={28} />
                    <p>Drag &amp; drop your file here</p>
                    <button type="button" className="primary-button" onClick={() => fileInputRef.current?.click()}>
                      Choose File
                    </button>
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept=".csv,.xlsx,.txt"
                      hidden
                      onChange={(event) => handleFileSelection(event.target.files?.[0] ?? null)}
                    />
                  </div>
                </div>
              </section>
            )}

            {(activeSection === 'clean' || activeSection === 'home' || activeSection === 'upload') && (
              <div className="workspace-grid">
                <div className="main-workspace">
                  {dataset.length > 0 && (
                    <>
                      <div className="panel-header space-between">
                        <h3>Data Preview</h3>
                        <div className="search-box">
                          <Search size={14} />
                          <input value={searchText} onChange={(event) => setSearchText(event.target.value)} placeholder="Search rows" />
                        </div>
                      </div>

                      <div className="table-wrap">
                        <table className="data-table">
                          <thead>
                            <tr>
                              {currentColumns.map((column) => (
                                <th key={column}>{column}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {activeRows.map((row, rowIndex) => (
                              <tr key={`${rowIndex}-${Object.values(row).join('-')}`}>
                                {currentColumns.map((column) => {
                                  const value = row[column] ?? '';
                                  const isEmpty = String(value).trim() === '';
                                  return (
                                    <td key={`${rowIndex}-${column}`} className={isEmpty ? 'empty-cell' : ''}>
                                      <input
                                        value={String(value)}
                                        onChange={(event) => handleCellChange(rowIndex, column, event.target.value)}
                                      />
                                    </td>
                                  );
                                })}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>

                      <div className="stats-grid">
                        <div className="stat-card">
                          <span>Total Rows</span>
                          <strong>{stats.totalRows}</strong>
                        </div>
                        <div className="stat-card">
                          <span>Total Columns</span>
                          <strong>{stats.totalColumns}</strong>
                        </div>
                        <div className="stat-card">
                          <span>Duplicates Found</span>
                          <strong>{stats.duplicatesFound}</strong>
                        </div>
                        <div className="stat-card">
                          <span>Missing Values</span>
                          <strong>{stats.missingValues}</strong>
                        </div>
                      </div>

                      <div className="panel-header">
                        <h3>Before / After Comparison</h3>
                      </div>
                      <div className="comparison-grid">
                        <div className="comparison-card">
                          <h4>Before Cleaning</h4>
                          <div className="comparison-list">
                            <span>Rows: {beforeStats.totalRows}</span>
                            <span>Columns: {beforeStats.totalColumns}</span>
                            <span>Duplicates: {beforeStats.duplicatesFound}</span>
                            <span>Missing: {beforeStats.missingValues}</span>
                          </div>
                        </div>
                        <div className="comparison-card accent">
                          <h4>After Cleaning</h4>
                          <div className="comparison-list">
                            <span>Rows: {stats.totalRows}</span>
                            <span>Columns: {stats.totalColumns}</span>
                            <span>Duplicates: {stats.duplicatesFound}</span>
                            <span>Missing: {stats.missingValues}</span>
                          </div>
                        </div>
                      </div>
                    </>
                  )}
                </div>

                {activeSection === 'clean' && (
                  <aside className="options-panel">
                    <div className="panel-header">
                      <h3>Cleaning Options</h3>
                    </div>

                    <div className="option-toggle">
                      <label>
                        <input type="checkbox" checked={cleaningConfig.removeDuplicates} onChange={(event) => setCleaningConfig((current) => ({ ...current, removeDuplicates: event.target.checked }))} />
                        <span>Remove Duplicates</span>
                      </label>
                      <small>Keep only unique records</small>
                    </div>

                    <div className="option-toggle">
                      <label>
                        <span>Handle Missing Values</span>
                      </label>
                      <select value={cleaningConfig.missingMode} onChange={(event) => setCleaningConfig((current) => ({ ...current, missingMode: event.target.value as CleaningConfig['missingMode'] }))}>
                        <option value="removeRows">Remove rows</option>
                        <option value="fillBlank">Fill with blank</option>
                        <option value="fillZero">Fill with 0</option>
                        <option value="fillMean">Fill with mean</option>
                        <option value="fillMedian">Fill with median</option>
                        <option value="fillMode">Fill with mode</option>
                        <option value="custom">Custom value</option>
                      </select>
                      {cleaningConfig.missingMode === 'custom' && (
                        <input value={cleaningConfig.customValue} onChange={(event) => setCleaningConfig((current) => ({ ...current, customValue: event.target.value }))} placeholder="Custom value" />
                      )}
                    </div>

                    <div className="option-toggle">
                      <label>
                        <input type="checkbox" checked={cleaningConfig.trimSpaces} onChange={(event) => setCleaningConfig((current) => ({ ...current, trimSpaces: event.target.checked }))} />
                        <span>Trim Spaces</span>
                      </label>
                    </div>

                    <div className="option-toggle">
                      <label>
                        <input type="checkbox" checked={cleaningConfig.standardizeText} onChange={(event) => setCleaningConfig((current) => ({ ...current, standardizeText: event.target.checked }))} />
                        <span>Standardize Text</span>
                      </label>
                      {cleaningConfig.standardizeText && (
                        <select value={cleaningConfig.textCase} onChange={(event) => setCleaningConfig((current) => ({ ...current, textCase: event.target.value as CleaningConfig['textCase'] }))}>
                          <option value="UPPERCASE">UPPERCASE</option>
                          <option value="lowercase">lowercase</option>
                          <option value="Title Case">Title Case</option>
                          <option value="Sentence case">Sentence case</option>
                        </select>
                      )}
                    </div>

                    <div className="option-toggle">
                      <label>
                        <input type="checkbox" checked={cleaningConfig.formatDates} onChange={(event) => setCleaningConfig((current) => ({ ...current, formatDates: event.target.checked }))} />
                        <span>Format Dates</span>
                      </label>
                      {cleaningConfig.formatDates && (
                        <select value={cleaningConfig.dateFormat} onChange={(event) => setCleaningConfig((current) => ({ ...current, dateFormat: event.target.value as CleaningConfig['dateFormat'] }))}>
                          <option value="YYYY-MM-DD">YYYY-MM-DD</option>
                          <option value="DD-MM-YYYY">DD-MM-YYYY</option>
                          <option value="MM-DD-YYYY">MM-DD-YYYY</option>
                        </select>
                      )}
                    </div>

                    <div className="option-toggle">
                      <label>
                        <input type="checkbox" checked={cleaningConfig.removeEmptyRows} onChange={(event) => setCleaningConfig((current) => ({ ...current, removeEmptyRows: event.target.checked }))} />
                        <span>Remove Empty Rows</span>
                      </label>
                    </div>

                    <div className="option-toggle">
                      <label>
                        <input type="checkbox" checked={cleaningConfig.removeEmptyColumns} onChange={(event) => setCleaningConfig((current) => ({ ...current, removeEmptyColumns: event.target.checked }))} />
                        <span>Remove Empty Columns</span>
                      </label>
                    </div>

                    <div className="advanced-block">
                      <div className="advanced-header">Advanced Options</div>
                      <div className="inline-actions">
                        <button type="button" onClick={() => applyAdvancedAction('standardizeColumns')}>Standardize Names</button>
                        <button type="button" onClick={() => applyAdvancedAction('removeEmptyRows')}>Remove Empty Rows</button>
                        <button type="button" onClick={() => applyAdvancedAction('removeEmptyColumns')}>Remove Empty Columns</button>
                        <button type="button" onClick={() => applyAdvancedAction('removeSpecialCharacters')}>Remove Special Characters</button>
                      </div>
                      <div className="replace-box">
                        <input value={replacePattern} onChange={(event) => setReplacePattern(event.target.value)} placeholder="Find" />
                        <input value={replaceValue} onChange={(event) => setReplaceValue(event.target.value)} placeholder="Replace" />
                        <button type="button" onClick={() => applyAdvancedAction('replace')}>Replace All</button>
                      </div>
                    </div>

                    <button type="button" className="clean-button" onClick={handleClean} disabled={isProcessing || !dataset.length}>
                      {isProcessing ? <LoaderCircle className="spin" size={16} /> : <Sparkles size={16} />}
                      {isProcessing ? 'Cleaning your data...' : '✨ Clean Data'}
                    </button>
                  </aside>
                )}
              </div>
            )}

            {activeSection === 'preview' && (
              <section className="preview-page">
                <div className="toolbar-row">
                  <button type="button" onClick={() => executeRowAction('addRow')}><Plus size={15} /> Add Row</button>
                  <button type="button" onClick={() => executeRowAction('deleteRow')}><Trash2 size={15} /> Delete Row</button>
                  <button type="button" onClick={() => executeRowAction('addColumn')}><Plus size={15} /> Add Column</button>
                  <button type="button" onClick={() => executeRowAction('renameColumn')}><Pencil size={15} /> Rename Column</button>
                  <button type="button" onClick={() => executeRowAction('undo')}><ArrowRight size={15} /> Undo</button>
                  <button type="button" onClick={() => executeRowAction('redo')}><ArrowRight size={15} /> Redo</button>
                  <button type="button" className="primary-button" onClick={() => setStatus({ kind: 'success', message: 'Changes saved successfully.' })}><Save size={15} /> Save Changes</button>
                </div>

                <div className="preview-tools">
                  <div className="search-box wide">
                    <Filter size={14} />
                    <input value={searchText} onChange={(event) => setSearchText(event.target.value)} placeholder="Filter data" />
                  </div>
                  <select value={selectedTypeColumn} onChange={(event) => setSelectedTypeColumn(event.target.value)}>
                    <option value="">Choose column</option>
                    {columns.map((column) => <option key={column} value={column}>{column}</option>)}
                  </select>
                  <select value={columnTypeTarget} onChange={(event) => setColumnTypeTarget(event.target.value as 'text' | 'integer' | 'decimal' | 'date' | 'boolean')}>
                    <option value="text">Text</option>
                    <option value="integer">Integer</option>
                    <option value="decimal">Decimal</option>
                    <option value="date">Date</option>
                    <option value="boolean">Boolean</option>
                  </select>
                  <button type="button" onClick={() => {
                    if (!selectedTypeColumn) return;
                    setDataset((current) => convertColumnType(current, selectedTypeColumn, columnTypeTarget));
                    setStatus({ kind: 'success', message: `Column "${selectedTypeColumn}" converted to ${columnTypeTarget}.` });
                  }}>Apply Type</button>
                </div>

                <div className="table-wrap large-table">
                  <table className="data-table">
                    <thead>
                      <tr>
                        {columns.map((column) => <th key={column}>{column}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows.map((row, rowIndex) => (
                        <tr key={`preview-${rowIndex}`}>
                          {columns.map((column) => (
                            <td key={`${rowIndex}-${column}`}>
                              <input
                                value={String(row[column] ?? '')}
                                onChange={(event) => handleCellChange(rowIndex, column, event.target.value)}
                              />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            )}

            {activeSection === 'transform' && (
              <section className="transform-page">
                <div className="settings-card">
                  <h2>Transform Data</h2>
                  <div className="toggle-row">
                    <span>Column</span>
                    <select value={transformColumn} onChange={(event) => setTransformColumn(event.target.value)}>
                      <option value="">Select column</option>
                      {columns.map((column) => <option key={column} value={column}>{column}</option>)}
                    </select>
                  </div>
                  <div className="toggle-row">
                    <span>Rename to</span>
                    <input value={transformNewName} onChange={(event) => setTransformNewName(event.target.value)} placeholder="New column name" />
                  </div>
                  <div className="toggle-row">
                    <span>Second column</span>
                    <select value={transformSecondColumn} onChange={(event) => setTransformSecondColumn(event.target.value)}>
                      <option value="">Select second column</option>
                      {columns.map((column) => <option key={column} value={column}>{column}</option>)}
                    </select>
                  </div>
                  <div className="toggle-row">
                    <span>Target type</span>
                    <select value={transformTargetType} onChange={(event) => setTransformTargetType(event.target.value as 'text' | 'integer' | 'decimal' | 'date' | 'boolean')}>
                      <option value="text">Text</option>
                      <option value="integer">Integer</option>
                      <option value="decimal">Decimal</option>
                      <option value="date">Date</option>
                      <option value="boolean">Boolean</option>
                    </select>
                  </div>
                  <div className="toggle-row">
                    <span>Sort direction</span>
                    <select value={transformSortDirection} onChange={(event) => setTransformSortDirection(event.target.value as 'asc' | 'desc')}>
                      <option value="asc">Ascending</option>
                      <option value="desc">Descending</option>
                    </select>
                  </div>
                  <div className="download-actions">
                    <button type="button" onClick={() => applyTransformAction('rename')}>Rename Columns</button>
                    <button type="button" onClick={() => applyTransformAction('type')}>Change Data Type</button>
                    <button type="button" onClick={() => applyTransformAction('replace')}>Replace Values</button>
                    <button type="button" onClick={() => applyTransformAction('merge')}>Merge Columns</button>
                    <button type="button" onClick={() => applyTransformAction('remove')}>Remove Column</button>
                    <button type="button" onClick={() => applyTransformAction('sort')}>Sort Data</button>
                    <button type="button" onClick={() => applyTransformAction('reorder')}>Reorder Columns</button>
                  </div>
                </div>
                <div className="table-wrap large-table">
                  <table className="data-table">
                    <thead>
                      <tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr>
                    </thead>
                    <tbody>
                      {previewRows.map((row, rowIndex) => (
                        <tr key={`transform-${rowIndex}`}>
                          {columns.map((column) => <td key={`${rowIndex}-${column}`}>{String(row[column] ?? '')}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            )}

            {activeSection === 'download' && (
              <section className="download-page">
                <div className="download-card">
                  <div className="download-badge">Your cleaned data is ready!</div>
                  <h2>{fileMeta?.name ?? 'cleaned_data.csv'}</h2>
                  <div className="download-info-grid">
                    <div><span>Rows</span><strong>{dataset.length}</strong></div>
                    <div><span>Columns</span><strong>{columns.length}</strong></div>
                    <div><span>Operations</span><strong>{Object.values(cleaningConfig).filter(Boolean).length}</strong></div>
                  </div>
                  <div className="download-actions">
                    <button type="button" onClick={() => handleDownload('csv')}>Download CSV</button>
                    <button type="button" onClick={() => handleDownload('xlsx')}>Download Excel</button>
                    <button type="button" onClick={() => handleDownload('txt')}>Download TXT</button>
                  </div>
                  {lastCleanedAt && <small>Last cleaned: {lastCleanedAt}</small>}
                </div>
              </section>
            )}

            {activeSection === 'settings' && (
              <section className="settings-page">
                <div className="settings-card">
                  <h2>Appearance</h2>
                  <label className="toggle-row">
                    <span>Dark theme</span>
                    <input type="checkbox" checked readOnly />
                  </label>
                  <label className="toggle-row">
                    <span>Compact mode</span>
                    <input type="checkbox" checked={compactMode} onChange={(event) => setCompactMode(event.target.checked)} />
                  </label>
                  <label className="toggle-row">
                    <span>Table density</span>
                    <select value={tableDensity} onChange={(event) => setTableDensity(event.target.value)}>
                      <option value="comfortable">Comfortable</option>
                      <option value="compact">Compact</option>
                      <option value="dense">Dense</option>
                    </select>
                  </label>
                </div>

                <div className="settings-card">
                  <h2>Data Settings</h2>
                  <label className="toggle-row">
                    <span>Default date format</span>
                    <select value={cleaningConfig.dateFormat} onChange={(event) => setCleaningConfig((current) => ({ ...current, dateFormat: event.target.value as CleaningConfig['dateFormat'] }))}>
                      <option value="YYYY-MM-DD">YYYY-MM-DD</option>
                      <option value="DD-MM-YYYY">DD-MM-YYYY</option>
                      <option value="MM-DD-YYYY">MM-DD-YYYY</option>
                    </select>
                  </label>
                  <label className="toggle-row">
                    <span>Default missing-value behavior</span>
                    <select value={cleaningConfig.missingMode} onChange={(event) => setCleaningConfig((current) => ({ ...current, missingMode: event.target.value as CleaningConfig['missingMode'] }))}>
                      <option value="removeRows">Remove rows</option>
                      <option value="fillBlank">Fill blank</option>
                      <option value="fillZero">Fill with 0</option>
                      <option value="fillMean">Fill with mean</option>
                      <option value="fillMedian">Fill with median</option>
                      <option value="fillMode">Fill with mode</option>
                      <option value="custom">Custom value</option>
                    </select>
                  </label>
                  <label className="toggle-row">
                    <span>Default export format</span>
                    <select>
                      <option value="csv">CSV</option>
                      <option value="xlsx">XLSX</option>
                      <option value="txt">TXT</option>
                    </select>
                  </label>
                </div>

                <div className="settings-card">
                  <h2>File Settings</h2>
                  <div className="info-line">Maximum file size: 50 MB</div>
                  <div className="info-line">Supported types: CSV, XLSX, TXT</div>
                </div>
              </section>
            )}
          </main>

          <aside
            id="assistant-panel"
            className={`assistant-panel ${assistantOpen ? 'open' : 'closed'}`}
          >
            <div className="assistant-header">
              <div className="assistant-brand">
                <Sparkles size={16} />
                <span>ChatGPT</span>
              </div>
              <button
                type="button"
                className="assistant-toggle"
                aria-label={assistantOpen ? 'Hide ChatGPT panel' : 'Open ChatGPT panel'}
                aria-expanded={assistantOpen}
                aria-controls="assistant-content"
                title={assistantOpen ? 'Hide ChatGPT panel' : 'Open ChatGPT panel'}
                onClick={() => setAssistantOpen((current) => !current)}
              >
                {assistantOpen ? <ChevronRight size={17} /> : <ChevronLeft size={17} />}
              </button>
            </div>

            <div
              id="assistant-content"
              className="assistant-content"
              aria-hidden={!assistantOpen}
            >
                <div className="assistant-chat">
                  {assistantMessages.map((message) => (
                    <div key={message.id} className={`assistant-message ${message.role}`}>
                      {message.text}
                    </div>
                  ))}
                </div>

                <div className="assistant-quick-actions">
                  <button type="button" onClick={() => setAssistantInput('Remove duplicates and trim spaces')}>Remove duplicates</button>
                  <button type="button" onClick={() => setAssistantInput('Fill missing values with 0')}>Fill missing</button>
                  <button type="button" onClick={() => setAssistantInput('Convert dates to YYYY-MM-DD')}>Format dates</button>
                </div>

                <div className="assistant-input-row">
                  <input
                    value={assistantInput}
                    onChange={(event) => setAssistantInput(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter') { handleAssistantSubmit(); }
                    }}
                    placeholder="Ask the AI to clean your data..."
                  />
                  <button type="button" className="primary-button" onClick={handleAssistantSubmit}>Send</button>
                </div>
            </div>
          </aside>
        </div>
      </div>
    </div>
  );
}

export default App;

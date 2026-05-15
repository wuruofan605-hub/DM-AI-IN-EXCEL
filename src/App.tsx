import {
  ArrowUpDown,
  BarChart3,
  Bell,
  BookmarkPlus,
  Bot,
  Check,
  ChevronDown,
  ChevronRight,
  CircleUserRound,
  Clipboard,
  ClipboardCheck,
  ClipboardPaste,
  Copy,
  Database,
  Eraser,
  FileSpreadsheet,
  FilePlus2,
  Filter,
  FolderClosed,
  FunctionSquare,
  Grid2X2,
  History,
  Link,
  LayoutTemplate,
  LineChart,
  MessageSquarePlus,
  MessageSquareText,
  MoreHorizontal,
  Newspaper,
  PanelRightClose,
  Paperclip,
  Plus,
  RefreshCw,
  Search,
  Send,
  ShieldCheck,
  Sparkles,
  Scissors,
  Settings2,
  SquarePen,
  Table2,
  Tag,
  Trash2,
  UploadCloud,
  X,
} from 'lucide-react';
import type { LucideIcon } from 'lucide-react';
import { ChangeEvent, FormEvent, MouseEvent, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  baseRows,
  buildInitialOutputConfigs,
  columns,
  defaultTemplates,
  findFieldCandidates,
  findMetricCandidates,
  initialMetricBindings,
  makeCellAddress,
  metricDefinitions,
  mockFetchMetricValue,
  sheet,
  workbook,
} from './mockData';
import type {
  AiConfirmationCard,
  AiContextChip,
  AiRequest,
  CellSelection,
  ChatMessage,
  ColumnParamType,
  FieldDefinition,
  MetricBinding,
  MetricDefinition,
  OutputCellConfig,
  OutputStatus,
  SaveTemplateDraft,
  TemplateConfig,
} from './types';

type MenuState = {
  x: number;
  y: number;
  selection: CellSelection;
} | null;

type InputCellMode = 'system' | 'custom';
type HeaderColumnMode = 'system' | 'custom';
type SuggestionContext = 'metricRow' | 'headerCol';

type AiConversation = {
  id: string;
  title: string;
  messages: ChatMessage[];
  pinned: boolean;
  updatedAt: string;
};

const initialSelection: CellSelection = {
  sheetId: workbook.activeSheetId,
  sheetName: workbook.activeSheetName,
  address: 'A13',
  activeCell: 'A13',
};

const contextReferences: AiContextChip[] = [{ id: 'ref_sheet1', type: 'reference', label: '@Sheet1' }];

function createIntroMessages(): ChatMessage[] {
  return [
    {
      id: `assistant_intro_${Date.now()}`,
      role: 'assistant',
      text: '我已绑定当前工作簿。输入区用于放指标和参数，输出区只显示刷新后的结果；你可以右键选区，也可以在这里用自然语言告诉我想新增什么指标。',
    },
  ];
}

const initialConversationId = 'conversation_current';

type DetectedObjectType = 'issuer' | 'bond' | 'unknown';

type ImportedColumnRole = 'entity_key' | 'recognized_metric' | 'custom_metric' | 'ignored';

type ImportedColumnProfile = {
  index: number;
  header: string;
  role: ImportedColumnRole;
  metricCode?: string;
  fetchInterface?: string;
  fieldName?: string;
};

type WorkbookSheetState = {
  sheetId: string;
  sheetName: string;
  rows: string[][];
  years: string[];
  outputConfigs: Record<string, OutputCellConfig>;
  metricDrafts: Record<string, string>;
  metricBindings: MetricBinding[];
  inputCellModes: Record<string, InputCellMode>;
  headerColumnModes: Record<string, HeaderColumnMode>;
  columnParamType: ColumnParamType;
  activeColumnCount: number;
  detectedObjectType: DetectedObjectType;
  columnProfiles: ImportedColumnProfile[];
};

type ImportDraftSheet = WorkbookSheetState & {
  previewRows: string[][];
  customMetricCount: number;
};

type ImportDraft = {
  fileName: string;
  sheets: ImportDraftSheet[];
  activeSheetId: string;
};

type ProductModule = {
  id: string;
  title: string;
  subtitle: string;
  status: string;
  icon: LucideIcon;
  tone: string;
  summary: string;
  flow: string[];
  output: string;
};

const productModules: ProductModule[] = [
  {
    id: 'data-assistant',
    title: '数据助手',
    subtitle: 'Excel 提数 / 指标绑定',
    status: '当前主功能',
    icon: Database,
    tone: 'blue',
    summary: '围绕 Excel 选区识别指标、确认口径、批量刷新数据，并沉淀为可复用模板。',
    flow: ['选中表格区域', '识别指标和年份', '确认数据口径', '刷新并审计来源'],
    output: '结构化数值表、隐藏取数配置、可复用模板',
  },
  {
    id: 'research-view',
    title: '研报观点',
    subtitle: '观点抽取 / 证据追踪',
    status: '规划中',
    icon: Newspaper,
    tone: 'green',
    summary: '把研报、纪要和公告拆成观点、论据、风险提示，并能回填到 Excel 或生成对比摘要。',
    flow: ['上传研报材料', '抽取核心观点', '关联数据证据', '生成观点卡片'],
    output: '观点摘要、原文出处、行业/主体标签',
  },
  {
    id: 'credit-analysis',
    title: '信用分析',
    subtitle: '主体画像 / 风险预警',
    status: '规划中',
    icon: ShieldCheck,
    tone: 'amber',
    summary: '面向发行人和债项整合财务、舆情、估值、评级和条款信息，形成信用风险工作流。',
    flow: ['输入主体或债券', '拉取财务和估值', '识别风险变化', '输出信用结论'],
    output: '主体画像、风险因子、跟踪清单',
  },
  {
    id: 'quant-backtest',
    title: '量化回测',
    subtitle: '策略验证 / 组合复盘',
    status: '规划中',
    icon: LineChart,
    tone: 'purple',
    summary: '支持用自然语言配置利率、信用、存单策略参数，回测收益、回撤和归因。',
    flow: ['定义策略条件', '选择资产池和区间', '运行回测', '分析收益归因'],
    output: '净值曲线、绩效指标、交易明细',
  },
];

const initialWorkbookSheet: WorkbookSheetState = {
  sheetId: workbook.activeSheetId,
  sheetName: workbook.activeSheetName,
  rows: baseRows,
  years: baseRows[0].slice(1, 8),
  outputConfigs: buildInitialOutputConfigs(baseRows[0].slice(1, 8), initialMetricBindings),
  metricDrafts: {},
  metricBindings: initialMetricBindings,
  inputCellModes: Object.fromEntries(initialMetricBindings.map((binding) => [binding.cell, 'system' as InputCellMode])),
  headerColumnModes: Object.fromEntries(columns.slice(1, 8).map((column) => [column, 'system' as HeaderColumnMode])),
  columnParamType: 'year',
  activeColumnCount: 8,
  detectedObjectType: 'unknown',
  columnProfiles: [],
};

function App() {
  const [gridRows, setGridRows] = useState(baseRows);
  const [years, setYears] = useState(baseRows[0].slice(1, 8));
  const [metricBindings, setMetricBindings] = useState<MetricBinding[]>(initialMetricBindings);
  const [outputConfigs, setOutputConfigs] = useState<Record<string, OutputCellConfig>>(() =>
    buildInitialOutputConfigs(baseRows[0].slice(1, 8), initialMetricBindings),
  );
  const [selection, setSelection] = useState<CellSelection>(initialSelection);
  const [menu, setMenu] = useState<MenuState>(null);
  const [input, setInput] = useState('????');
  const [messages, setMessages] = useState<ChatMessage[]>(() => createIntroMessages());
  const [activeConversationId, setActiveConversationId] = useState(initialConversationId);
  const [activeConversationTitle, setActiveConversationTitle] = useState('???????');
  const [aiConversations, setAiConversations] = useState<AiConversation[]>([
    {
      id: initialConversationId,
      title: '???????',
      messages: createIntroMessages(),
      pinned: true,
      updatedAt: '??',
    },
    {
      id: 'conversation_template_setup',
      title: '??????????',
      messages: [{ id: 'history_template_1', role: 'assistant', text: '?????????????????????????' }],
      pinned: false,
      updatedAt: '??',
    },
    {
      id: 'conversation_import_review',
      title: 'Excel ????????',
      messages: [{ id: 'history_import_1', role: 'assistant', text: '?????/????????????????????????' }],
      pinned: false,
      updatedAt: '??',
    },
  ]);
  const [aiPanelCollapsed, setAiPanelCollapsed] = useState(false);
  const [confirmationCard, setConfirmationCard] = useState<AiConfirmationCard | null>(null);
  const [saveDraft, setSaveDraft] = useState<SaveTemplateDraft | null>(null);
  const [customTemplates, setCustomTemplates] = useState<TemplateConfig[]>([]);
  const [templateModalOpen, setTemplateModalOpen] = useState(false);
  const [formulaPanelMode, setFormulaPanelMode] = useState<'audit' | 'export' | null>(null);
  const [appliedNote, setAppliedNote] = useState('');
  const [activeSuggestionCell, setActiveSuggestionCell] = useState<string | null>(null);
  const [metricDrafts, setMetricDrafts] = useState<Record<string, string>>({});
  const [inputCellModes, setInputCellModes] = useState<Record<string, InputCellMode>>(initialWorkbookSheet.inputCellModes);
  const [headerColumnModes, setHeaderColumnModes] = useState<Record<string, HeaderColumnMode>>(initialWorkbookSheet.headerColumnModes);
  const [columnParamType, setColumnParamType] = useState<ColumnParamType>('year');
  const [activeColumnCount, setActiveColumnCount] = useState(8);
  const [activeModuleId, setActiveModuleId] = useState(productModules[0].id);
  const [viewMode, setViewMode] = useState<'templateHome' | 'workbook'>('templateHome');
  const [selectedTemplateId, setSelectedTemplateId] = useState<string | null>(null);
  const [templateSearch, setTemplateSearch] = useState('');
  const [templatePickerOpen, setTemplatePickerOpen] = useState(false);
  const [templatePickerSearch, setTemplatePickerSearch] = useState('');
  const [workbookSheets, setWorkbookSheets] = useState<WorkbookSheetState[]>([initialWorkbookSheet]);
  const [activeSheetId, setActiveSheetId] = useState(workbook.activeSheetId);
  const [workbookFileName, setWorkbookFileName] = useState(workbook.workbookName);
  const [importDialogOpen, setImportDialogOpen] = useState(false);
  const [importDraft, setImportDraft] = useState<ImportDraft | null>(null);
  const [importError, setImportError] = useState('');

  const displayRows = useMemo(() => {
    const activeSheetForRows = workbookSheets.find((item) => item.sheetId === activeSheetId);
    const rows = gridRows.map((row) => [...row]);
    if (!activeSheetForRows?.columnProfiles.length) {
      rows[0] = ['指标', ...years];
    }
    Object.values(outputConfigs).forEach((config) => {
      const { row, column } = parseCell(config.targetCell);
      rows[row - 1] = [...(rows[row - 1] ?? [])];
      rows[row - 1][column - 1] = renderOutputValue(config);
    });
    return rows;
  }, [activeSheetId, gridRows, outputConfigs, workbookSheets, years]);

  const selectedChip = useMemo<AiContextChip>(
    () => ({
      id: 'selected_range',
      type: 'selected',
      label: `${selection.sheetName} ${selection.address} selected`,
    }),
    [selection],
  );

  const selectedFormulaValue = useMemo(() => {
    const config = outputConfigs[selection.activeCell];
    if (config) {
      return `隐藏取数配置: ${config.fetchInterface ?? 'DMDATA'}(${config.metricCode}; param=${config.parameterCell ?? config.metricCell}; field=${config.fieldName ?? config.columnParamCell})`;
    }
    return getCellValue(displayRows, selection.activeCell);
  }, [displayRows, outputConfigs, selection.activeCell]);

  const activeModule = useMemo(
    () => productModules.find((module) => module.id === activeModuleId) ?? productModules[0],
    [activeModuleId],
  );

  const selectedTemplate = useMemo(
    () => defaultTemplates.find((template) => template.templateId === selectedTemplateId) ?? null,
    [selectedTemplateId],
  );

  const filteredTemplates = useMemo(() => {
    return filterTemplates(defaultTemplates, templateSearch);
  }, [templateSearch]);

  const pickerTemplates = useMemo(() => filterTemplates(defaultTemplates, templatePickerSearch), [templatePickerSearch]);

  const activeWorkbookSheet = useMemo(
    () => workbookSheets.find((item) => item.sheetId === activeSheetId) ?? workbookSheets[0],
    [activeSheetId, workbookSheets],
  );

  const visibleAiConversations = useMemo(() => {
    const snapshot: AiConversation = {
      id: activeConversationId,
      title: activeConversationTitle,
      messages,
      pinned: aiConversations.find((item) => item.id === activeConversationId)?.pinned ?? false,
      updatedAt: '刚刚',
    };
    return [snapshot, ...aiConversations.filter((item) => item.id !== activeConversationId)];
  }, [activeConversationId, activeConversationTitle, aiConversations, messages]);

  const currentWorkbook = useMemo(
    () => ({
      ...workbook,
      workbookName: workbookFileName,
      activeSheetId: activeWorkbookSheet?.sheetId ?? workbook.activeSheetId,
      activeSheetName: activeWorkbookSheet?.sheetName ?? workbook.activeSheetName,
    }),
    [activeWorkbookSheet, workbookFileName],
  );

  const currentSheet = useMemo(
    () => ({
      ...sheet,
      sheetId: activeWorkbookSheet?.sheetId ?? sheet.sheetId,
      sheetName: activeWorkbookSheet?.sheetName ?? sheet.sheetName,
      rowCount: Math.max(55, activeWorkbookSheet?.rows.length ?? sheet.rowCount),
      columnCount: columns.length,
    }),
    [activeWorkbookSheet],
  );

  const handleSelect = (address: string) => {
    setSelection({
      sheetId: activeWorkbookSheet?.sheetId ?? workbook.activeSheetId,
      sheetName: activeWorkbookSheet?.sheetName ?? workbook.activeSheetName,
      address,
      activeCell: address.split(':')[0],
    });
    setActiveSuggestionCell(!selectedTemplate && getSuggestionContext(address, activeColumnCount) ? address : null);
  };

  const normalizeColumnParam = (value: string) => {
    if (columnParamType === 'year') return value.replace(/[^\d]/g, '').slice(0, 4);
    return value.trim();
  };

  const refreshConfiguredCells = (cellKeys: string[]) => {
    if (cellKeys.length === 0) return;
    const now = new Date().toISOString().slice(0, 16).replace('T', ' ');
    setOutputConfigs((current) => {
      const next = { ...current };
      cellKeys.forEach((cell) => {
        const config = next[cell];
        if (!config) return;
        next[cell] = {
          ...config,
          status: config.metricCode && config.columnParam ? 'ready' : 'empty',
          value: config.metricCode && config.columnParam ? mockFetchMetricValue(config.metricCode, config.columnParam) : '',
          lastRefreshAt: config.metricCode && config.columnParam ? now : undefined,
        };
      });
      return next;
    });
  };

  const updateYear = (columnIndex: number, value: string) => {
    const cleanYear = normalizeColumnParam(value);
    const nextYears = [...years];
    nextYears[columnIndex - 2] = cleanYear;
    setYears(nextYears);

    setOutputConfigs((current) => {
      const next = { ...current };
      Object.values(next).forEach((config) => {
        if (config.columnParamCell === `${columns[columnIndex - 1]}1`) {
          config.columnParam = cleanYear;
          config.status = config.metricCode && cleanYear ? 'ready' : 'empty';
          config.value = config.metricCode && cleanYear ? mockFetchMetricValue(config.metricCode, cleanYear) : '';
          config.lastRefreshAt = config.metricCode && cleanYear ? new Date().toISOString().slice(0, 16).replace('T', ' ') : undefined;
        }
      });
      return next;
    });
  };

  const bindMetricToRow = (row: number, metric: MetricDefinition) => {
    const nextRows = gridRows.map((item) => [...item]);
    nextRows[row - 1] = [...(nextRows[row - 1] ?? [])];
    nextRows[row - 1][0] = metric.label;
    setGridRows(nextRows);

    const binding: MetricBinding = {
      row,
      cell: `A${row}`,
      displayName: metric.label,
      metricCode: metric.code,
      source: metric.source,
      disambiguationStatus: 'confirmed',
    };
    setMetricBindings((current) => [...current.filter((item) => item.row !== row), binding]);

    setOutputConfigs((current) => {
      const next = { ...current };
      years.forEach((year, index) => {
        const columnIndex = index + 2;
        const targetCell = makeCellAddress(columnIndex, row);
        next[targetCell] = {
          targetCell,
          metricCell: `A${row}`,
          columnParamCell: `${columns[columnIndex - 1]}1`,
          metricCode: metric.code,
          columnParam: year,
          status: year ? 'ready' : 'empty',
          value: year ? mockFetchMetricValue(metric.code, year) : '',
          lastRefreshAt: year ? new Date().toISOString().slice(0, 16).replace('T', ' ') : undefined,
        };
      });
      return next;
    });
    setActiveSuggestionCell(null);
    setMetricDrafts((current) => ({ ...current, [`A${row}`]: metric.label }));
    setInputCellModes((current) => ({ ...current, [`A${row}`]: 'system' }));
  };

  const keepCustomInput = (cell: string, value: string) => {
    const { row } = parseCell(cell);
    const nextRows = gridRows.map((item) => [...item]);
    nextRows[row - 1] = [...(nextRows[row - 1] ?? [])];
    nextRows[row - 1][0] = value;
    setGridRows(nextRows);
    setMetricDrafts((current) => ({ ...current, [cell]: value }));
    setInputCellModes((current) => ({ ...current, [cell]: 'custom' }));
    setMetricBindings((current) => current.filter((item) => item.row !== row));
    setOutputConfigs((current) => {
      const next = { ...current };
      years.forEach((_, index) => {
        delete next[makeCellAddress(index + 2, row)];
      });
      return next;
    });
    setActiveSuggestionCell(null);
  };

  const updateInputDraft = (cell: string, value: string) => {
    const { row } = parseCell(cell);
    const nextRows = gridRows.map((item) => [...item]);
    nextRows[row - 1] = [...(nextRows[row - 1] ?? [])];
    nextRows[row - 1][0] = value;
    setGridRows(nextRows);
    setMetricDrafts((current) => ({ ...current, [cell]: value }));
    setInputCellModes((current) => ({ ...current, [cell]: 'custom' }));
  };

  const updateResultCell = (cell: string, value: string) => {
    const { row, column } = parseCell(cell);
    const nextRows = gridRows.map((item) => [...item]);
    nextRows[row - 1] = [...(nextRows[row - 1] ?? [])];
    nextRows[row - 1][column - 1] = value;
    setGridRows(nextRows);
  };

  const bindFieldToColumn = (columnIndex: number, field: FieldDefinition) => {
    const columnLetter = columns[columnIndex - 1];
    const nextYears = [...years];
    nextYears[columnIndex - 2] = field.code;
    setYears(nextYears);
    setHeaderColumnModes((current) => ({ ...current, [columnLetter]: 'system' }));
    setOutputConfigs((current) => {
      const next = { ...current };
      metricBindings.forEach((binding) => {
        const targetCell = makeCellAddress(columnIndex, binding.row);
        next[targetCell] = {
          targetCell,
          metricCell: `A${binding.row}`,
          columnParamCell: `${columnLetter}1`,
          metricCode: binding.metricCode,
          columnParam: field.code,
          status: 'ready',
          value: mockFetchMetricValue(binding.metricCode, field.code),
          lastRefreshAt: new Date().toISOString().slice(0, 16).replace('T', ' '),
          fieldName: field.label,
        };
      });
      return next;
    });
    setActiveSuggestionCell(null);
  };

  const keepCustomHeader = (columnIndex: number, value: string) => {
    const columnLetter = columns[columnIndex - 1];
    const nextYears = [...years];
    nextYears[columnIndex - 2] = value.trim();
    setYears(nextYears);
    setHeaderColumnModes((current) => ({ ...current, [columnLetter]: 'custom' }));
    setOutputConfigs((current) => {
      const next = { ...current };
      Object.keys(next).forEach((cell) => {
        if (next[cell].columnParamCell === `${columnLetter}1`) delete next[cell];
      });
      return next;
    });
    setActiveSuggestionCell(null);
  };

  const updateHeaderDraft = (columnIndex: number, value: string) => {
    const columnLetter = columns[columnIndex - 1];
    const nextYears = [...years];
    nextYears[columnIndex - 2] = value;
    setYears(nextYears);
    setHeaderColumnModes((current) => ({ ...current, [columnLetter]: 'custom' }));
  };

  const addColumn = () => {
    if (activeColumnCount >= columns.length) return;
    const newActiveCount = activeColumnCount + 1;
    const columnLetter = columns[newActiveCount - 1];
    setActiveColumnCount(newActiveCount);
    setYears((current) => {
      const next = [...current];
      next[newActiveCount - 2] = '';
      return next;
    });
    setHeaderColumnModes((current) => ({ ...current, [columnLetter]: 'custom' }));
    setActiveSuggestionCell(`${columnLetter}1`);
  };

  const refreshOutputs = () => {
    refreshConfiguredCells(Object.keys(outputConfigs).filter((cell) => outputConfigs[cell].status === 'ready'));
    setMessages((current) => [
      ...current,
      {
        id: `assistant_refresh_${Date.now()}`,
        role: 'assistant',
        text: '已按隐藏取数配置刷新待更新单元格。输出区仍然只展示数值，配置链路可在“公式审计”里查看。',
      },
    ]);
  };

  const submitRequest = (query: string, source: 'chat' | 'context-menu', selected = selection) => {
    const request = buildAiRequest(query, selected, currentWorkbook, currentSheet);
    setMessages((current) => [...current, { id: `user_${Date.now()}`, role: 'user', text: query, request }]);

    const candidates = findMetricCandidates(query);
    const looksLikeMetricRequest = candidates.some((item) => item.aliases.some((alias) => query.includes(alias)));
    if (looksLikeMetricRequest || query.includes('生猪')) {
      const targetCell = isMetricInputCell(selected.activeCell) ? selected.activeCell : 'A13';
      setConfirmationCard({
        id: `confirm_${Date.now()}`,
        type: 'metricBinding',
        status: 'pending',
        targetCell,
        query,
        candidates: candidates.length ? candidates : findMetricCandidates('生猪价格'),
        selectedMetricCode: (candidates[0] ?? findMetricCandidates('生猪价格')[0]).code,
      });
      setMessages((current) => [
        ...current,
        {
          id: `assistant_confirm_${Date.now()}`,
          role: 'assistant',
          text: `我识别到你想在 ${targetCell} 新增指标。请先确认口径，确认后我会写入指标绑定，并把对应输出区标记为待刷新。`,
        },
      ]);
      return;
    }

    if (source === 'context-menu' && query.includes('补全数据')) {
      refreshOutputs();
      return;
    }

    setMessages((current) => [
      ...current,
      {
        id: `assistant_generic_${Date.now()}`,
        role: 'assistant',
        text: `已定位到 ${request.selectedRange.sheetName}!${request.selectedRange.address}。当前原型重点展示指标绑定、手动刷新和模板沉淀。`,
      },
    ]);
  };

  const confirmMetricBinding = () => {
    if (!confirmationCard) return;
    const metric = metricDefinitions.find((item) => item.code === confirmationCard.selectedMetricCode);
    if (!metric) return;
    const row = parseCell(confirmationCard.targetCell).row;
    bindMetricToRow(row, metric);
    setConfirmationCard({ ...confirmationCard, status: 'confirmed' });
    setSelection({
      sheetId: activeWorkbookSheet?.sheetId ?? workbook.activeSheetId,
      sheetName: activeWorkbookSheet?.sheetName ?? workbook.activeSheetName,
      address: confirmationCard.targetCell,
      activeCell: confirmationCard.targetCell,
    });
    setMessages((current) => [
      ...current,
      {
        id: `assistant_bound_${Date.now()}`,
        role: 'assistant',
        text: `已将 ${metric.label} 绑定到 ${confirmationCard.targetCell}。这一步保存的是指标口径和取数配置，不是写死数值；点击“刷新”后会拉取 B${row}:H${row}。`,
      },
    ]);
  };

  const requestSaveTemplate = () => {
    setSaveDraft({
      id: `save_${Date.now()}`,
      status: 'pending',
      templateName: '我的宏观指标模板',
      inputRanges: ['A2:A13', 'B1:H1'],
      outputRange: 'B2:H13',
      refreshRules: '按指标绑定和年份输入批量刷新输出区。',
      formatPreset: '淡黄输入区 + 淡蓝输出区',
    });
  };

  const confirmSaveTemplate = () => {
    if (!saveDraft) return;
    const template: TemplateConfig = {
      templateId: `tpl_custom_${Date.now()}`,
      templateName: saveDraft.templateName,
      description: '由当前工作表保存，可复用指标输入、年份参数和刷新配置。',
      inputRanges: saveDraft.inputRanges,
      outputRange: saveDraft.outputRange,
      metricBindings,
      refreshRules: saveDraft.refreshRules,
      formatPreset: saveDraft.formatPreset,
      type: 'custom',
      applicableObject: 'bond_set',
      requiredParams: [
        { name: '指标或证券代码', description: '从当前输入区保存的核心参数。', placeholder: '沿用当前工作表输入区' },
      ],
      optionalParams: [
        { name: '时间或筛选条件', description: '从当前表头和筛选条件保存。', placeholder: '沿用当前工作表表头' },
      ],
      defaultOutputFields: years.filter(Boolean),
      hiddenFetchConfig: {
        interfaceName: 'DM_CUSTOM_TEMPLATE',
        parameterMapping: { inputRange: saveDraft.inputRanges.join(',') },
        fieldMapping: Object.fromEntries(years.filter(Boolean).map((year) => [year, year])),
      },
      previewRows: [
        ['输入区', saveDraft.inputRanges.join(' / '), ''],
        ['输出区', saveDraft.outputRange, ''],
        ['规则', saveDraft.refreshRules, ''],
      ],
    };
    setCustomTemplates((current) => [template, ...current]);
    setSaveDraft({ ...saveDraft, status: 'saved' });
    setTemplateModalOpen(true);
    setAppliedNote('已保存到模板广场');
  };

  const handleSubmit = (event: FormEvent) => {
    event.preventDefault();
    const query = input.trim();
    if (!query) return;
    submitRequest(query, 'chat');
    setInput('');
  };

  const upsertActiveConversation = (items: AiConversation[]) => {
    const snapshot: AiConversation = {
      id: activeConversationId,
      title: activeConversationTitle,
      messages,
      pinned: items.find((item) => item.id === activeConversationId)?.pinned ?? false,
      updatedAt: '刚刚',
    };
    const withoutCurrent = items.filter((item) => item.id !== activeConversationId);
    return [snapshot, ...withoutCurrent].sort((a, b) => Number(b.pinned) - Number(a.pinned));
  };

  const startNewAiConversation = () => {
    setAiConversations((current) => upsertActiveConversation(current));
    const nextId = `conversation_${Date.now()}`;
    setActiveConversationId(nextId);
    setActiveConversationTitle('新的 AI 对话');
    setMessages([
      {
        id: `assistant_new_${Date.now()}`,
        role: 'assistant',
        text: '新的 AI 对话已创建。我会继续绑定当前工作簿和选区，你可以直接提出新的取数、审计或模板问题。',
      },
    ]);
    setInput('');
    setFormulaPanelMode(null);
    setConfirmationCard(null);
    setSaveDraft(null);
  };

  const openAiConversation = (conversation: AiConversation) => {
    setAiConversations((current) => upsertActiveConversation(current));
    setActiveConversationId(conversation.id);
    setActiveConversationTitle(conversation.title);
    setMessages(conversation.messages);
    setInput('');
    setFormulaPanelMode(null);
  };

  const toggleConversationPin = (conversationId: string) => {
    setAiConversations((current) =>
      current
        .map((item) => (item.id === conversationId ? { ...item, pinned: !item.pinned } : item))
        .sort((a, b) => Number(b.pinned) - Number(a.pinned)),
    );
  };

  const renameConversation = (conversationId: string, title: string) => {
    const cleanTitle = title.trim();
    if (!cleanTitle) return;
    setAiConversations((current) => current.map((item) => (item.id === conversationId ? { ...item, title: cleanTitle } : item)));
    if (conversationId === activeConversationId) setActiveConversationTitle(cleanTitle);
  };

  const deleteConversation = (conversationId: string) => {
    setAiConversations((current) => current.filter((item) => item.id !== conversationId));
    if (conversationId === activeConversationId) {
      const nextId = `conversation_${Date.now()}`;
      setActiveConversationId(nextId);
      setActiveConversationTitle('新的 AI 对话');
      setMessages([
        {
          id: `assistant_new_${Date.now()}`,
          role: 'assistant',
          text: '已删除当前会话，并为你打开一个新的 AI 对话。',
        },
      ]);
      setInput('');
      setFormulaPanelMode(null);
    }
  };

  const openContextMenu = (event: MouseEvent, address = selection.address) => {
    event.preventDefault();
    setMenu({
      x: event.clientX,
      y: event.clientY,
      selection: {
        sheetId: activeWorkbookSheet?.sheetId ?? workbook.activeSheetId,
        sheetName: activeWorkbookSheet?.sheetName ?? workbook.activeSheetName,
        address,
        activeCell: address.split(':')[0],
      },
    });
  };

  const runMenuAction = (label: string) => {
    if (!menu) return;
    setSelection(menu.selection);
    submitRequest(`${label}：${menu.selection.sheetName}!${menu.selection.address}`, 'context-menu', menu.selection);
    setMenu(null);
  };

  const saveActiveSheetSnapshot = () => {
    setWorkbookSheets((current) =>
      current.map((item) =>
        item.sheetId === activeSheetId
          ? {
              ...item,
              rows: gridRows,
              years,
              outputConfigs,
              metricDrafts,
              metricBindings,
              inputCellModes,
              headerColumnModes,
              columnParamType,
              activeColumnCount,
            }
          : item,
      ),
    );
  };

  const switchSheet = (sheetId: string) => {
    if (sheetId === activeSheetId) return;
    saveActiveSheetSnapshot();
    const target = workbookSheets.find((item) => item.sheetId === sheetId);
    if (!target) return;
    setActiveSheetId(target.sheetId);
    setGridRows(target.rows);
    setYears(target.years);
    setOutputConfigs(target.outputConfigs);
    setMetricDrafts(target.metricDrafts);
    setMetricBindings(target.metricBindings);
    setInputCellModes(target.inputCellModes);
    setHeaderColumnModes(target.headerColumnModes);
    setColumnParamType(target.columnParamType);
    setActiveColumnCount(target.activeColumnCount);
    setSelection({
      sheetId: target.sheetId,
      sheetName: target.sheetName,
      address: 'A1',
      activeCell: 'A1',
    });
    setActiveSuggestionCell(null);
  };

  const openImportDialog = () => {
    setImportError('');
    setImportDraft(null);
    setImportDialogOpen(true);
  };

  const handleImportFile = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setImportError('');
    try {
      const buffer = await file.arrayBuffer();
      const parsedWorkbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheets = parsedWorkbook.SheetNames.map((sheetName, index) => {
        const worksheet = parsedWorkbook.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json<Array<string | number | Date | boolean | null>>(worksheet, {
          header: 1,
          defval: '',
          blankrows: false,
        });
        return buildImportedSheetState(rawRows, sheetName, index);
      }).filter((item) => item.rows.some((row) => row.some(Boolean)));

      if (sheets.length === 0) {
        setImportError('这个 Excel 没有识别到可导入的二维表。');
        return;
      }

      setImportDraft({
        fileName: file.name,
        sheets,
        activeSheetId: sheets[0].sheetId,
      });
    } catch {
      setImportError('Excel 解析失败，请确认文件是 .xlsx 格式，且没有被加密。');
    } finally {
      event.target.value = '';
    }
  };

  const updateImportSheetType = (sheetId: string, objectType: DetectedObjectType) => {
    setImportDraft((current) => {
      if (!current) return current;
      const sheets = current.sheets.map((item) =>
        item.sheetId === sheetId ? rebuildImportedSheetWithObjectType(item, objectType) : item,
      );
      return { ...current, sheets };
    });
  };

  const confirmImportWorkbook = () => {
    if (!importDraft) return;
    const sheets = importDraft.sheets.map(({ previewRows, customMetricCount, ...sheetState }) => sheetState);
    const firstSheet = sheets[0];
    setWorkbookSheets(sheets);
    setActiveSheetId(firstSheet.sheetId);
    setWorkbookFileName(importDraft.fileName);
    setSelectedTemplateId(null);
    setGridRows(firstSheet.rows);
    setYears(firstSheet.years);
    setOutputConfigs(firstSheet.outputConfigs);
    setMetricDrafts(firstSheet.metricDrafts);
    setMetricBindings(firstSheet.metricBindings);
    setInputCellModes(firstSheet.inputCellModes);
    setHeaderColumnModes(firstSheet.headerColumnModes);
    setColumnParamType(firstSheet.columnParamType);
    setActiveColumnCount(firstSheet.activeColumnCount);
    setSelection({
      sheetId: firstSheet.sheetId,
      sheetName: firstSheet.sheetName,
      address: 'A1',
      activeCell: 'A1',
    });
    setImportDialogOpen(false);
    setViewMode('workbook');
    setFormulaPanelMode('audit');
    setMessages((current) => [
      ...current,
      {
        id: `assistant_import_${Date.now()}`,
        role: 'assistant',
        text: `已导入 ${importDraft.fileName}，保留了 ${sheets.length} 个 Sheet。蓝色字段已识别为可取数指标，红色字段是自定义或暂未识别字段，不会被刷新覆盖。`,
      },
    ]);
  };

  const updateImportedHeader = (columnIndex: number, value: string) => {
    if (!activeWorkbookSheet || activeWorkbookSheet.columnProfiles.length === 0) return;
    const nextRows = gridRows.map((row) => [...row]);
    nextRows[0] = [...(nextRows[0] ?? [])];
    nextRows[0][columnIndex] = value;
    const rebuilt = buildImportedSheetState(
      nextRows,
      activeWorkbookSheet.sheetName,
      Number(activeWorkbookSheet.sheetId.replace(/\D/g, '')) || 0,
      activeWorkbookSheet.detectedObjectType,
    );
    const nextSheet: WorkbookSheetState = {
      ...rebuilt,
      sheetId: activeWorkbookSheet.sheetId,
      sheetName: activeWorkbookSheet.sheetName,
    };
    setGridRows(nextSheet.rows);
    setYears(nextSheet.years);
    setOutputConfigs(nextSheet.outputConfigs);
    setMetricDrafts(nextSheet.metricDrafts);
    setMetricBindings(nextSheet.metricBindings);
    setInputCellModes(nextSheet.inputCellModes);
    setHeaderColumnModes(nextSheet.headerColumnModes);
    setColumnParamType(nextSheet.columnParamType);
    setActiveColumnCount(nextSheet.activeColumnCount);
    setWorkbookSheets((current) => current.map((item) => (item.sheetId === activeSheetId ? nextSheet : item)));
  };

  const applyTemplate = (template: TemplateConfig) => {
    const workbookState = buildTemplateWorkbook(template);
    const nextSheet: WorkbookSheetState = {
      sheetId: `sheet_${template.templateId}`,
      sheetName: 'Sheet1',
      rows: workbookState.rows,
      years: workbookState.headers,
      outputConfigs: workbookState.outputConfigs,
      metricDrafts: workbookState.metricDrafts,
      metricBindings: [],
      inputCellModes: Object.fromEntries(Object.keys(workbookState.metricDrafts).map((cell) => [cell, 'custom' as InputCellMode])),
      headerColumnModes: Object.fromEntries(columns.slice(1, workbookState.headers.length + 1).map((column) => [column, 'system' as HeaderColumnMode])),
      columnParamType: template.columnParamType ?? 'fieldCode',
      activeColumnCount: Math.max(2, workbookState.headers.length + 1),
      detectedObjectType: template.applicableObject === 'issuer' ? 'issuer' : 'bond',
      columnProfiles: [],
    };
    setTemplatePickerOpen(false);
    setSelectedTemplateId(template.templateId);
    setWorkbookSheets([nextSheet]);
    setActiveSheetId(nextSheet.sheetId);
    setWorkbookFileName(`${template.templateName}.xlsx`);
    setGridRows(workbookState.rows);
    setYears(workbookState.headers);
    setMetricBindings([]);
    setOutputConfigs(workbookState.outputConfigs);
    setMetricDrafts(workbookState.metricDrafts);
    setInputCellModes(Object.fromEntries(Object.keys(workbookState.metricDrafts).map((cell) => [cell, 'custom' as InputCellMode])));
    setHeaderColumnModes(Object.fromEntries(columns.slice(1, workbookState.headers.length + 1).map((column) => [column, 'system' as HeaderColumnMode])));
    setColumnParamType(template.columnParamType ?? 'fieldCode');
    setActiveColumnCount(Math.max(2, workbookState.headers.length + 1));
    setSelection({
      sheetId: nextSheet.sheetId,
      sheetName: nextSheet.sheetName,
      address: 'A2',
      activeCell: 'A2',
    });
    setActiveSuggestionCell(null);
    setViewMode('workbook');
    setFormulaPanelMode('audit');
    setMessages((current) => [
      ...current,
      {
        id: `assistant_template_${Date.now()}`,
        role: 'assistant',
        text: `当前已套用「${template.templateName}」。我已经放好必填参数、可选参数、默认输出字段和隐藏取数配置，后续刷新会按 ${template.hiddenFetchConfig.interfaceName} 的字段映射执行。`,
      },
    ]);
  };

  return (
    <div className="terminal-shell" onClick={() => setMenu(null)}>
      <TopBar
        activeModuleId={activeModuleId}
        modules={productModules}
        onModuleSelect={(moduleId) => {
          setActiveModuleId(moduleId);
          if (moduleId === 'data-assistant') setViewMode('templateHome');
        }}
      />
      {activeModule.id === 'data-assistant' && viewMode === 'templateHome' ? (
        <div className="template-home-layout">
          <LeftRail
            activeModule={activeModule}
            onOpenTemplates={() => setViewMode('templateHome')}
            onOpenTemplatePicker={() => {
              setTemplatePickerSearch('');
              setTemplatePickerOpen(true);
            }}
            customTemplateCount={customTemplates.length}
          />
          <TemplateHome
            templates={filteredTemplates}
            search={templateSearch}
            onSearchChange={setTemplateSearch}
            onOpenTemplate={applyTemplate}
            onOpenImport={openImportDialog}
          />
        </div>
      ) : (
        <div className={`workspace ${aiPanelCollapsed ? 'ai-collapsed' : ''}`}>
          <LeftRail
            activeModule={activeModule}
            onOpenTemplates={() => setViewMode('templateHome')}
            onOpenTemplatePicker={() => {
              setTemplatePickerSearch('');
              setTemplatePickerOpen(true);
            }}
            customTemplateCount={customTemplates.length}
          />
          {activeModule.id === 'data-assistant' ? (
            <main className="excel-pane">
              <WorkbookHeader
                selectedTemplate={selectedTemplate}
                workbookFileName={workbookFileName}
                onBackToTemplates={() => setViewMode('templateHome')}
              />
              <Ribbon
                onRefresh={refreshOutputs}
                onSaveTemplate={requestSaveTemplate}
                onShowAudit={() => setFormulaPanelMode('audit')}
                onExportTable={() => setFormulaPanelMode('export')}
                onOpenImport={openImportDialog}
              />
              <FormulaBar activeCell={selection.activeCell} value={selectedFormulaValue} />
              <SpreadsheetGrid
                rows={displayRows}
                selection={selection}
                years={years}
                outputConfigs={outputConfigs}
                columnProfiles={activeWorkbookSheet?.columnProfiles ?? []}
                activeSuggestionCell={activeSuggestionCell}
                enableMetricSuggestions={!selectedTemplate}
                metricDrafts={metricDrafts}
                inputCellModes={inputCellModes}
                headerColumnModes={headerColumnModes}
                columnParamType={columnParamType}
                activeColumnCount={activeColumnCount}
                onSelect={handleSelect}
                onYearChange={updateYear}
                onMetricDraftChange={updateInputDraft}
                onBindMetric={bindMetricToRow}
                onBindField={bindFieldToColumn}
                onHeaderChange={updateImportedHeader}
                onHeaderDraftChange={updateHeaderDraft}
                onKeepCustomInput={keepCustomInput}
                onKeepCustomHeader={keepCustomHeader}
                onResultDraftChange={updateResultCell}
                onAddColumn={addColumn}
                onContextMenu={openContextMenu}
              />
              <SheetFooter
                selectedRange={selection.address}
                sheets={workbookSheets}
                activeSheetId={activeSheetId}
                onSwitchSheet={switchSheet}
              />
            </main>
          ) : (
            <ModuleWorkspace module={activeModule} />
          )}
          {aiPanelCollapsed ? (
            <button className="ai-collapsed-tab" type="button" onClick={() => setAiPanelCollapsed(false)} aria-label="展开 DM AI">
              <Bot size={18} />
              <span>AI</span>
            </button>
          ) : (
            <AiPanel
              input={input}
              setInput={setInput}
              selectedChip={selectedChip}
              messages={messages}
              conversations={visibleAiConversations}
              activeConversationId={activeConversationId}
              activeConversationTitle={activeConversationTitle}
              confirmationCard={confirmationCard}
              saveDraft={saveDraft}
              appliedNote={appliedNote}
              outputConfigs={outputConfigs}
              metricBindings={metricBindings}
              formulaPanelMode={formulaPanelMode}
              onSubmit={handleSubmit}
              onNewConversation={startNewAiConversation}
              onOpenConversation={openAiConversation}
              onToggleConversationPin={toggleConversationPin}
              onRenameConversation={renameConversation}
              onDeleteConversation={deleteConversation}
              onCollapse={() => setAiPanelCollapsed(true)}
              onMetricChoice={(metricCode) =>
                confirmationCard && setConfirmationCard({ ...confirmationCard, selectedMetricCode: metricCode })
              }
              onConfirmMetric={confirmMetricBinding}
              onCancelMetric={() => confirmationCard && setConfirmationCard({ ...confirmationCard, status: 'cancelled' })}
              onTemplateNameChange={(templateName) => saveDraft && setSaveDraft({ ...saveDraft, templateName })}
              onConfirmSaveTemplate={confirmSaveTemplate}
              onCancelSaveTemplate={() => saveDraft && setSaveDraft({ ...saveDraft, status: 'cancelled' })}
              onCloseFormulaPanel={() => setFormulaPanelMode(null)}
            />
          )}
        </div>
      )}
      {menu && <ContextMenu x={menu.x} y={menu.y} onRun={runMenuAction} />}
      {templateModalOpen && (
        <TemplateGalleryModal
          defaultTemplates={defaultTemplates}
          customTemplates={customTemplates}
          onClose={() => setTemplateModalOpen(false)}
        />
      )}
      {templatePickerOpen && (
        <TemplatePickerModal
          templates={pickerTemplates}
          search={templatePickerSearch}
          onSearchChange={setTemplatePickerSearch}
          onOpenTemplate={applyTemplate}
          onClose={() => setTemplatePickerOpen(false)}
        />
      )}
      {importDialogOpen && (
        <ImportWorkbookModal
          draft={importDraft}
          error={importError}
          onFileChange={handleImportFile}
          onClose={() => setImportDialogOpen(false)}
          onConfirm={confirmImportWorkbook}
          onSheetTypeChange={updateImportSheetType}
        />
      )}
    </div>
  );
}

function TopBar({
  activeModuleId,
  modules,
  onModuleSelect,
}: {
  activeModuleId: string;
  modules: ProductModule[];
  onModuleSelect: (moduleId: string) => void;
}) {
  return (
    <header className="topbar">
      <div className="brand">
        <Grid2X2 size={18} />
        <strong className="dm-logo">
          D<span>M</span>
        </strong>
        <em>Intl.</em>
      </div>
      <nav className="global-modules" aria-label="一级能力导航">
        {modules.map((module) => {
          const Icon = module.icon;
          return (
            <button
              className={activeModuleId === module.id ? 'active' : ''}
              key={module.id}
              onClick={() => onModuleSelect(module.id)}
            >
              <Icon size={14} />
              {module.title}
            </button>
          );
        })}
      </nav>
      <div className="top-actions">
        <Bell size={16} />
        <MessageSquareText size={16} />
        <CircleUserRound size={20} />
        <span>吴若凡</span>
        <b>3,293</b>
      </div>
    </header>
  );
}

function LeftRail({
  activeModule,
  onOpenTemplates,
  onOpenTemplatePicker,
  customTemplateCount,
}: {
  activeModule: ProductModule;
  onOpenTemplates: () => void;
  onOpenTemplatePicker: () => void;
  customTemplateCount: number;
}) {
  const workspaces = ['资金', '利率债', '信用债', '存单'];

  return (
    <aside className="left-rail">
      <h2>{activeModule.title}</h2>
      {activeModule.id === 'data-assistant' ? (
        <>
          <button className="rail-item active" onClick={onOpenTemplates}>
            <LayoutTemplate size={16} />
            模板广场
            <span>{customTemplateCount ? `我的 ${customTemplateCount}` : '模板中心'}</span>
          </button>
          <div className="rail-section">
            <div className="rail-section-title">
              <strong>工作区</strong>
              <Search size={14} />
              <button type="button" className="rail-icon-button" onClick={onOpenTemplatePicker} aria-label="新建工作区">
                <Plus size={14} />
              </button>
            </div>
            {workspaces.map((item) => (
              <div className="file-row" key={item}>
                <FolderClosed size={14} />
                <span>{item}</span>
              </div>
            ))}
          </div>
        </>
      ) : (
        <ModuleRail module={activeModule} />
      )}
    </aside>
  );
}

function ModuleRail({ module }: { module: ProductModule }) {
  const sections = {
    'research-view': ['观点库', '研报收件箱', '主题跟踪', '证据摘录'],
    'credit-analysis': ['主体池', '债项跟踪', '风险预警', '评级变动'],
    'quant-backtest': ['策略库', '资产池', '回测任务', '绩效归因'],
  }[module.id] ?? ['工作台', '任务', '输出'];

  return (
    <div className="rail-section">
      <div className="rail-section-title">
        <strong>模块资源</strong>
        <Plus size={14} />
      </div>
      {sections.map((item, index) => (
        <div className={`file-row ${index === 0 ? 'active-resource' : ''}`} key={item}>
          <FolderClosed size={14} />
          <span>{item}</span>
        </div>
      ))}
    </div>
  );
}

function ModuleWorkspace({ module }: { module: ProductModule }) {
  const Icon = module.icon;

  return (
    <main className="module-workspace">
      <section className={`module-hero ${module.tone}`}>
        <div className="module-hero-title">
          <Icon size={24} />
          <div>
            <strong>{module.title}</strong>
            <span>{module.subtitle}</span>
          </div>
          <b>{module.status}</b>
        </div>
        <p>{module.summary}</p>
      </section>
      <section className="module-board">
        <div className="module-board-head">
          <strong>未来工作流</strong>
          <span>{module.output}</span>
        </div>
        <div className="module-stage-grid">
          {module.flow.map((step, index) => (
            <article className="module-stage" key={step}>
              <span>{String(index + 1).padStart(2, '0')}</span>
              <strong>{step}</strong>
              <p>{moduleStageCopy(module.id, index)}</p>
            </article>
          ))}
        </div>
      </section>
    </main>
  );
}

function TemplateHome({
  templates,
  search,
  onSearchChange,
  onOpenTemplate,
  onOpenImport,
}: {
  templates: TemplateConfig[];
  search: string;
  onSearchChange: (value: string) => void;
  onOpenTemplate: (template: TemplateConfig) => void;
  onOpenImport: () => void;
}) {
  const issuerTemplates = templates.filter((template) => template.applicableObject === 'issuer');
  const bondTemplates = templates.filter((template) => template.applicableObject !== 'issuer');

  return (
    <main className="template-home">
      <section className="template-home-toolbar">
        <label className="template-search">
          <Search size={22} />
          <input
            value={search}
            onChange={(event) => onSearchChange(event.target.value)}
            placeholder="搜索模板：城投、评级、余额、估值"
            aria-label="搜索模板"
          />
        </label>
      </section>
      <section className="template-home-content">
        <section className="template-import-entry">
          <button className="import-entry-card" type="button" onClick={onOpenImport}>
            <span className="import-entry-icon">
              <UploadCloud size={30} />
            </span>
            <div>
              <strong>导入 Excel</strong>
              <p>上传客户已有工作簿，保留多个 Sheet，自动识别主体/债券列和可取数字段。</p>
            </div>
            <small>未识别指标会标红保留，不参与刷新覆盖</small>
          </button>
        </section>
        <TemplateGroup title="主体模板" subtitle="按主体名称、报告期和授信指标取数" templates={issuerTemplates} onOpenTemplate={onOpenTemplate} />
        <TemplateGroup title="债券模板" subtitle="按债券代码、估值日和发行区间取数" templates={bondTemplates} onOpenTemplate={onOpenTemplate} />
      </section>
    </main>
  );
}

function TemplateGroup({
  title,
  subtitle,
  templates,
  onOpenTemplate,
}: {
  title: string;
  subtitle: string;
  templates: TemplateConfig[];
  onOpenTemplate: (template: TemplateConfig) => void;
}) {
  if (templates.length === 0) {
    return (
      <section className="template-group">
        <div className="template-section-title">
          <h1>{title}</h1>
          <span>没有匹配模板</span>
        </div>
      </section>
    );
  }

  return (
    <section className="template-group">
      <div className="template-section-title">
        <h1>{title}</h1>
        <span>{subtitle}</span>
      </div>
      <div className="template-card-grid">
        {templates.map((template) => (
          <TemplateLibraryCard template={template} key={template.templateId} onOpen={() => onOpenTemplate(template)} />
        ))}
      </div>
    </section>
  );
}

function TemplateLibraryCard({ template, onOpen }: { template: TemplateConfig; onOpen: () => void }) {
  return (
    <button className="template-library-card" type="button" onClick={onOpen}>
      <div className="template-card-heading">
        <span className="template-app-icon">
          <Grid2X2 size={15} />
        </span>
        <strong>{template.templateName}</strong>
      </div>
      <div className="template-preview-table">
        {template.previewRows.map((row, rowIndex) => (
          <div className="template-preview-row" key={`${template.templateId}-${rowIndex}`}>
            {row.map((cell, cellIndex) => (
              <span key={`${template.templateId}-${rowIndex}-${cellIndex}`}>{cell}</span>
            ))}
          </div>
        ))}
      </div>
      <div className="template-boundary">
        <span>适用对象：{objectScopeText(template.applicableObject)}</span>
        <span>必填：{template.requiredParams.map((param) => param.name).join('、')}</span>
        <span>输出：{template.defaultOutputFields.slice(0, 4).join('、')}</span>
      </div>
      <code>{template.hiddenFetchConfig.interfaceName}</code>
    </button>
  );
}

function TemplatePickerModal({
  templates,
  search,
  onSearchChange,
  onOpenTemplate,
  onClose,
}: {
  templates: TemplateConfig[];
  search: string;
  onSearchChange: (value: string) => void;
  onOpenTemplate: (template: TemplateConfig) => void;
  onClose: () => void;
}) {
  const issuerTemplates = templates.filter((template) => template.applicableObject === 'issuer');
  const bondTemplates = templates.filter((template) => template.applicableObject !== 'issuer');

  return (
    <div className="modal-backdrop template-picker-backdrop" onClick={onClose}>
      <section className="template-picker-modal" onClick={(event) => event.stopPropagation()}>
        <div className="template-picker-header">
          <label className="template-search picker-search">
            <Search size={22} />
            <input
              value={search}
              onChange={(event) => onSearchChange(event.target.value)}
              placeholder="搜索模板：城投、评级、余额、估值"
              aria-label="搜索弹窗模板"
            />
          </label>
          <button type="button" onClick={onClose} aria-label="关闭模板选择">
            <X size={18} />
          </button>
        </div>
        <div className="template-picker-content">
          <TemplateGroup title="主体模板" subtitle="按主体名称、报告期和授信指标取数" templates={issuerTemplates} onOpenTemplate={onOpenTemplate} />
          <TemplateGroup title="债券模板" subtitle="按债券代码、估值日和发行区间取数" templates={bondTemplates} onOpenTemplate={onOpenTemplate} />
        </div>
      </section>
    </div>
  );
}

function ImportWorkbookModal({
  draft,
  error,
  onFileChange,
  onClose,
  onConfirm,
  onSheetTypeChange,
}: {
  draft: ImportDraft | null;
  error: string;
  onFileChange: (event: ChangeEvent<HTMLInputElement>) => void;
  onClose: () => void;
  onConfirm: () => void;
  onSheetTypeChange: (sheetId: string, objectType: DetectedObjectType) => void;
}) {
  return (
    <div className="modal-backdrop import-backdrop" onClick={onClose}>
      <section className="import-modal" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <div>
            <UploadCloud size={20} />
            <strong>导入 Excel</strong>
          </div>
          <button type="button" onClick={onClose} aria-label="关闭导入">
            <X size={16} />
          </button>
        </div>
        <div className="import-body">
          <label className="excel-upload-zone">
            <UploadCloud size={34} />
            <strong>上传 .xlsx 工作簿</strong>
            <span>系统会保留多个 Sheet，并逐表识别主体/债券字段和指标列。</span>
            <input type="file" accept=".xlsx" onChange={onFileChange} />
          </label>
          {error && <div className="import-error">{error}</div>}
          {draft && (
            <div className="import-preview">
              <div className="import-summary">
                <strong>{draft.fileName}</strong>
                <span>{draft.sheets.length} 个 Sheet</span>
              </div>
              <div className="import-sheet-list">
                {draft.sheets.map((sheetItem) => (
                  <article className="import-sheet-card" key={sheetItem.sheetId}>
                    <div className="import-sheet-head">
                      <div>
                        <strong>{sheetItem.sheetName}</strong>
                        <span>
                          {objectTypeLabel(sheetItem.detectedObjectType)} · {sheetItem.customMetricCount} 个未识别指标
                        </span>
                      </div>
                      <select
                        value={sheetItem.detectedObjectType}
                        onChange={(event) => onSheetTypeChange(sheetItem.sheetId, event.target.value as DetectedObjectType)}
                        aria-label={`${sheetItem.sheetName} 对象类型`}
                      >
                        <option value="bond">债券表</option>
                        <option value="issuer">主体表</option>
                        <option value="unknown">待确认</option>
                      </select>
                    </div>
                    <div className="import-preview-table">
                      {sheetItem.previewRows.map((row, rowIndex) => (
                        <div className="import-preview-row" key={`${sheetItem.sheetId}-${rowIndex}`}>
                          {row.slice(0, 6).map((cell, cellIndex) => (
                            <span
                              className={sheetItem.columnProfiles[cellIndex]?.role === 'custom_metric' ? 'preview-custom' : ''}
                              key={`${sheetItem.sheetId}-${rowIndex}-${cellIndex}`}
                            >
                              {cell}
                            </span>
                          ))}
                        </div>
                      ))}
                    </div>
                  </article>
                ))}
              </div>
            </div>
          )}
        </div>
        <div className="import-actions">
          <button type="button" onClick={onClose}>取消</button>
          <button className="primary" type="button" disabled={!draft} onClick={onConfirm}>
            导入工作台
          </button>
        </div>
      </section>
    </div>
  );
}

function WorkbookHeader({
  selectedTemplate,
  workbookFileName,
  onBackToTemplates,
}: {
  selectedTemplate: TemplateConfig | null;
  workbookFileName: string;
  onBackToTemplates: () => void;
}) {
  return (
    <section className="workbook-header">
      <div className="tabs">
        <div className="file-tab active">
          <FileSpreadsheet size={14} />
          {selectedTemplate ? `${selectedTemplate.templateName}.xlsx` : workbookFileName}
          <X size={13} />
        </div>
        <div className="file-tab">
          107_编辑在线文档与Excel.md
          <X size={13} />
        </div>
        <button className="tab-plus">+</button>
      </div>
      <div className="pathline">
        tree &gt; 模板库 &gt; {selectedTemplate?.templateName ?? '工作表'}
        <span>最近修改: 05月07 15:46</span>
        <button type="button" onClick={onBackToTemplates}>返回模板库</button>
      </div>
    </section>
  );
}

function Ribbon({
  onRefresh,
  onSaveTemplate,
  onShowAudit,
  onExportTable,
  onOpenImport,
}: {
  onRefresh: () => void;
  onSaveTemplate: () => void;
  onShowAudit: () => void;
  onExportTable: () => void;
  onOpenImport: () => void;
}) {
  return (
    <section className="ribbon">
      <nav>
        {['开始', '插入', '公式', '数据', '视图', '设置'].map((item, index) => (
          <button className={index === 0 ? 'active' : ''} key={item}>
            {item}
          </button>
        ))}
      </nav>
      <div className="ribbon-tools">
        <button>
          <Copy size={15} />
          复制
        </button>
        <button>
          <FunctionSquare size={15} />
          函数
        </button>
        <button onClick={onRefresh}>
          <RefreshCw size={15} />
          刷新
        </button>
        <button onClick={onSaveTemplate}>
          <BookmarkPlus size={15} />
          保存为模板
        </button>
        <button>
          <BarChart3 size={15} />
          图表
        </button>
        <button onClick={onShowAudit}>
          <FunctionSquare size={15} />
          公式审计
        </button>
        <button onClick={onExportTable}>
          <FileSpreadsheet size={15} />
          导出表格
        </button>
        <button onClick={onOpenImport}>
          <UploadCloud size={15} />
          导入 Excel
        </button>
      </div>
    </section>
  );
}

function FormulaBar({ activeCell, value }: { activeCell: string; value: string }) {
  return (
    <section className="formula-bar">
      <div className="cell-name">{activeCell}</div>
      <div className="formula-actions">
        <X size={15} />
        <Check size={15} />
        <FunctionSquare size={16} />
      </div>
      <input value={value} readOnly aria-label="公式栏" />
    </section>
  );
}

function SpreadsheetGrid({
  rows,
  selection,
  years,
  outputConfigs,
  columnProfiles,
  activeSuggestionCell,
  enableMetricSuggestions,
  metricDrafts,
  inputCellModes,
  headerColumnModes,
  columnParamType,
  activeColumnCount,
  onSelect,
  onYearChange,
  onMetricDraftChange,
  onBindMetric,
  onBindField,
  onHeaderChange,
  onHeaderDraftChange,
  onKeepCustomInput,
  onKeepCustomHeader,
  onResultDraftChange,
  onAddColumn,
  onContextMenu,
}: {
  rows: string[][];
  selection: CellSelection;
  years: string[];
  outputConfigs: Record<string, OutputCellConfig>;
  columnProfiles: ImportedColumnProfile[];
  activeSuggestionCell: string | null;
  enableMetricSuggestions: boolean;
  metricDrafts: Record<string, string>;
  inputCellModes: Record<string, InputCellMode>;
  headerColumnModes: Record<string, HeaderColumnMode>;
  columnParamType: ColumnParamType;
  activeColumnCount: number;
  onSelect: (address: string) => void;
  onYearChange: (columnIndex: number, value: string) => void;
  onMetricDraftChange: (cell: string, value: string) => void;
  onBindMetric: (row: number, metric: MetricDefinition) => void;
  onBindField: (columnIndex: number, field: FieldDefinition) => void;
  onHeaderChange: (columnIndex: number, value: string) => void;
  onHeaderDraftChange: (columnIndex: number, value: string) => void;
  onKeepCustomInput: (cell: string, value: string) => void;
  onKeepCustomHeader: (columnIndex: number, value: string) => void;
  onResultDraftChange: (cell: string, value: string) => void;
  onAddColumn: () => void;
  onContextMenu: (event: MouseEvent, address?: string) => void;
}) {
  const rowCount = 55;
  const activePosition = activeSuggestionCell ? parseCell(activeSuggestionCell) : null;
  const activeSuggestionRow = activePosition?.row ?? 0;
  const activeSuggestionColumn = activePosition?.column ?? 0;
  const activeSuggestionContext = activeSuggestionCell ? getSuggestionContext(activeSuggestionCell, activeColumnCount) : null;
  const activeQuery = activeSuggestionCell ? metricDrafts[activeSuggestionCell] ?? getCellValue(rows, activeSuggestionCell) : '';
  const fieldQuery =
    activeSuggestionCell && activeSuggestionContext === 'headerCol' ? years[activeSuggestionColumn - 2] ?? activeQuery : activeQuery;
  const metricSuggestions =
    activeSuggestionCell && enableMetricSuggestions && activeSuggestionContext === 'metricRow'
      ? findMetricCandidates(activeQuery)
      : [];
  const fieldSuggestions =
    activeSuggestionCell && enableMetricSuggestions && activeSuggestionContext === 'headerCol'
      ? findFieldCandidates(fieldQuery)
      : [];

  return (
    <section className="grid-wrap" onContextMenu={(event) => onContextMenu(event, selection.address)}>
      <table className="sheet-grid">
        <thead>
          <tr>
            <th className="corner" />
            {columns.slice(0, activeColumnCount).map((column) => (
              <th key={column}>{column}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: rowCount }, (_, rowIndex) => {
            const rowNumber = rowIndex + 1;
            const row = rows[rowIndex] ?? [];
            return (
              <tr key={rowNumber}>
                <th>{rowNumber}</th>
                {columns.slice(0, Math.min(columns.length, activeColumnCount + 1)).map((column, columnIndex) => {
                  const address = `${column}${rowNumber}`;
                  const isAddColumnCell = rowNumber === 1 && columnIndex === activeColumnCount && activeColumnCount < columns.length;
                  if (isAddColumnCell) {
                    return (
                      <td className="add-column-cell" key={address} onClick={onAddColumn}>
                        <Plus size={16} />
                      </td>
                    );
                  }
                  const value = row[columnIndex] ?? '';
                  const selected = isSelectedCell(rowNumber, columnIndex + 1, selection.address);
                  const isHeader = rowNumber === 1 && columnIndex < activeColumnCount;
                  const isMetricInput = columnIndex === 0 && rowNumber >= 2 && rowNumber <= 13;
                  const isYearInput = rowNumber === 1 && columnIndex >= 1 && columnIndex < activeColumnCount;
                  const isResultArea = columnIndex >= 1 && columnIndex < activeColumnCount && rowNumber >= 2 && rowNumber <= 13;
                  const outputConfig = outputConfigs[address];
                  const isOutput = Boolean(outputConfig);
                  const columnProfile = columnProfiles.find((profile) => profile.index === columnIndex);
                  const isImportedHeader = rowNumber === 1 && Boolean(columnProfile);
                  const isCustomMetric = isImportedHeader && columnProfile?.role === 'custom_metric';
                  const isRecognizedMetric = isImportedHeader && columnProfile?.role === 'recognized_metric';
                  const isEntityKey = isImportedHeader && columnProfile?.role === 'entity_key';

                  return (
                    <td
                      className={[
                        selected ? 'selected-cell' : '',
                        isHeader ? 'table-header-cell' : '',
                        isImportedHeader ? 'imported-header-cell' : '',
                        isMetricInput || isYearInput ? 'input-zone-cell' : '',
                        isMetricInput ? `input-mode-${inputCellModes[address] ?? 'custom'}` : '',
                        isYearInput ? 'column-param-input-cell' : '',
                        isYearInput ? `header-mode-${headerColumnModes[column] ?? 'custom'}` : '',
                        isResultArea ? 'result-zone-cell' : '',
                        isResultArea && !isOutput ? 'result-zone-custom' : '',
                        isOutput ? `output-cell status-${outputConfig.status}` : '',
                        isOutput ? 'number-cell' : '',
                      ].join(' ')}
                      key={address}
                      onClick={(event) => {
                        event.stopPropagation();
                        onSelect(address);
                      }}
                      onContextMenu={(event) => {
                        event.stopPropagation();
                        onContextMenu(event, selected ? selection.address : address);
                      }}
                    >
                      {isImportedHeader && (isCustomMetric || isRecognizedMetric) ? (
                        <input
                          className="metric-input imported-header-input"
                          value={value}
                          onChange={(event) => onHeaderChange(columnIndex, event.target.value)}
                          aria-label={`${address} 导入指标列名`}
                        />
                      ) : isYearInput ? (
                        <div className="input-cell-editor header-cell-editor">
                          <input
                            value={years[columnIndex - 1] ?? ''}
                            onChange={(event) => {
                              onHeaderDraftChange(columnIndex + 1, event.target.value);
                              onSelect(address);
                            }}
                            onBlur={(event) => {
                              if (headerColumnModes[column] !== 'system') onKeepCustomHeader(columnIndex + 1, event.target.value);
                            }}
                            onFocus={() => onSelect(address)}
                            aria-label={`${address} column parameter`}
                          />
                          <span>{headerColumnModes[column] === 'system' ? '字段' : '自定义'}</span>
                        </div>
                      ) : isMetricInput ? (
                        <div className="input-cell-editor">
                          <input
                            className="metric-input"
                            value={metricDrafts[address] ?? value}
                            placeholder={rowNumber === 13 ? '输入函数 / 指标 / 自定义项' : ''}
                            onChange={(event) => onMetricDraftChange(address, event.target.value)}
                            onFocus={() => onSelect(address)}
                          />
                          <span>{inputCellModes[address] === 'system' ? '函数' : '自定义'}</span>
                        </div>
                      ) : isResultArea && !isOutput ? (
                        <input
                          className="result-input"
                          value={value}
                          placeholder="可输入文本或函数"
                          onChange={(event) => onResultDraftChange(address, event.target.value)}
                          onFocus={() => onSelect(address)}
                        />
                      ) : (
                        value
                      )}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
      {activeSuggestionCell && enableMetricSuggestions && activeSuggestionContext && (
        <div className="suggestion-popover" style={{ top: activeSuggestionContext === 'headerCol' ? 78 : 78 + activeSuggestionRow * 22 }}>
          <div className="suggestion-title">
            <Search size={14} />
            {activeSuggestionContext === 'headerCol'
              ? `表头：${columnParamType === 'date' ? '选择日期字段' : columnParamType === 'year' ? '输入年份' : '选择系统字段'}`
              : '输入项：选择函数或保留自定义'}
          </div>
          {activeSuggestionContext === 'metricRow' && metricSuggestions.map((metric) => (
            <button key={metric.code} onMouseDown={(event) => event.preventDefault()} onClick={() => onBindMetric(activeSuggestionRow, metric)}>
              <strong>{metric.label}</strong>
              <span>{metric.category} · {metric.unit} · {metric.source}</span>
              <small>{metric.description}</small>
            </button>
          ))}
          {activeSuggestionContext === 'headerCol' && fieldSuggestions.map((field) => (
            <button key={field.code} onMouseDown={(event) => event.preventDefault()} onClick={() => onBindField(activeSuggestionColumn, field)}>
              <strong>{field.label}</strong>
              <span>{field.code} · {field.category} · {field.unit}</span>
              <small>{field.description}</small>
            </button>
          ))}
          <button
            className="custom-suggestion"
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => {
              if (!activeSuggestionCell) return;
              if (activeSuggestionContext === 'headerCol') onKeepCustomHeader(activeSuggestionColumn, fieldQuery);
              else onKeepCustomInput(activeSuggestionCell, activeQuery);
            }}
          >
            <strong>{activeSuggestionContext === 'headerCol' ? '保留为自定义列' : '保留为自定义输入'}</strong>
            <span>{activeSuggestionContext === 'headerCol' ? fieldQuery || '空白自定义列' : activeQuery || '空白自定义项'}</span>
            <small>自定义内容会保留，不绑定系统后端取数函数。</small>
          </button>
        </div>
      )}
    </section>
  );
}

function AiPanel({
  input,
  setInput,
  selectedChip,
  messages,
  conversations,
  activeConversationId,
  activeConversationTitle,
  confirmationCard,
  saveDraft,
  appliedNote,
  outputConfigs,
  metricBindings,
  formulaPanelMode,
  onSubmit,
  onNewConversation,
  onOpenConversation,
  onToggleConversationPin,
  onRenameConversation,
  onDeleteConversation,
  onCollapse,
  onMetricChoice,
  onConfirmMetric,
  onCancelMetric,
  onTemplateNameChange,
  onConfirmSaveTemplate,
  onCancelSaveTemplate,
  onCloseFormulaPanel,
}: {
  input: string;
  setInput: (value: string) => void;
  selectedChip: AiContextChip;
  messages: ChatMessage[];
  conversations: AiConversation[];
  activeConversationId: string;
  activeConversationTitle: string;
  confirmationCard: AiConfirmationCard | null;
  saveDraft: SaveTemplateDraft | null;
  appliedNote: string;
  outputConfigs: Record<string, OutputCellConfig>;
  metricBindings: MetricBinding[];
  formulaPanelMode: 'audit' | 'export' | null;
  onSubmit: (event: FormEvent) => void;
  onNewConversation: () => void;
  onOpenConversation: (conversation: AiConversation) => void;
  onToggleConversationPin: (conversationId: string) => void;
  onRenameConversation: (conversationId: string, title: string) => void;
  onDeleteConversation: (conversationId: string) => void;
  onCollapse: () => void;
  onMetricChoice: (metricCode: string) => void;
  onConfirmMetric: () => void;
  onCancelMetric: () => void;
  onTemplateNameChange: (value: string) => void;
  onConfirmSaveTemplate: () => void;
  onCancelSaveTemplate: () => void;
  onCloseFormulaPanel: () => void;
}) {
  const [historyOpen, setHistoryOpen] = useState(false);
  const [conversationMenuId, setConversationMenuId] = useState<string | null>(null);
  const sortedConversations = [...conversations].sort((a, b) => Number(b.pinned) - Number(a.pinned));

  return (
    <aside className="ai-panel">
      <header className="ai-header">
        <div>
          <Bot size={18} />
          <strong>{activeConversationTitle || 'DM AI'}</strong>
        </div>
        <div className="panel-icons">
          <button type="button" onClick={onNewConversation} aria-label="新建 AI 对话">
            <Plus size={15} />
          </button>
          <button type="button" onClick={() => setHistoryOpen(true)} aria-label="历史记录">
            <History size={16} />
          </button>
          <button type="button" onClick={onCollapse} aria-label="收起 DM AI">
            <PanelRightClose size={16} />
          </button>
        </div>
      </header>
      {historyOpen && (
        <section className="ai-history-panel">
          <div className="ai-history-header">
            <strong>历史记录</strong>
            <button type="button" onClick={() => setHistoryOpen(false)} aria-label="关闭历史记录">
              <X size={16} />
            </button>
          </div>
          <label className="ai-history-search">
            <Search size={14} />
            <input placeholder="搜索历史会话" />
          </label>
          <div className="ai-history-tabs">
            <button className="active" type="button">全部</button>
          </div>
          <span className="ai-history-period">近一周</span>
          <div className="ai-history-list">
            {sortedConversations.map((conversation) => (
              <article
                className={`ai-history-item ${conversation.id === activeConversationId ? 'active' : ''}`}
                key={conversation.id}
              >
                <button
                  className="ai-history-title"
                  type="button"
                  onClick={() => {
                    onOpenConversation(conversation);
                    setHistoryOpen(false);
                  }}
                >
                  <strong>{conversation.title}</strong>
                  <span>{conversation.pinned ? '置顶' : conversation.updatedAt}</span>
                </button>
                <button
                  className="ai-history-more"
                  type="button"
                  onClick={() => setConversationMenuId(conversationMenuId === conversation.id ? null : conversation.id)}
                  aria-label={`${conversation.title} 更多操作`}
                >
                  <MoreHorizontal size={16} />
                </button>
                {conversationMenuId === conversation.id && (
                  <div className="ai-history-menu">
                    <button type="button" onClick={() => onToggleConversationPin(conversation.id)}>
                      <BookmarkPlus size={14} />
                      {conversation.pinned ? '取消置顶' : '置顶'}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        const nextTitle = window.prompt('重命名会话', conversation.title);
                        if (nextTitle) onRenameConversation(conversation.id, nextTitle);
                        setConversationMenuId(null);
                      }}
                    >
                      <SquarePen size={14} />
                      重命名
                    </button>
                    <button type="button" onClick={() => onDeleteConversation(conversation.id)}>
                      <Trash2 size={14} />
                      删除会话
                    </button>
                  </div>
                )}
              </article>
            ))}
          </div>
        </section>
      )}
      <section className="chat-log">
        {formulaPanelMode && (
          <AuditPanel
            mode={formulaPanelMode}
            outputConfigs={outputConfigs}
            metricBindings={metricBindings}
            onClose={onCloseFormulaPanel}
          />
        )}
        {saveDraft && (
          <SaveTemplateCard
            draft={saveDraft}
            appliedNote={appliedNote}
            onNameChange={onTemplateNameChange}
            onConfirm={onConfirmSaveTemplate}
            onCancel={onCancelSaveTemplate}
          />
        )}
        {confirmationCard && (
          <MetricConfirmationCard
            card={confirmationCard}
            onChoice={onMetricChoice}
            onConfirm={onConfirmMetric}
            onCancel={onCancelMetric}
          />
        )}
        {messages.map((message) => (
          <article className={`message ${message.role}`} key={message.id}>
            <div className="message-meta">
              {message.role === 'assistant' ? 'DM AI 已绑定回答' : '你'}
              {message.request && (
                <span>
                  {message.request.workbook.activeSheetName} · {message.request.selectedRange.address}
                </span>
              )}
            </div>
            <p>{message.text}</p>
          </article>
        ))}
      </section>
      <form className="composer" onSubmit={onSubmit}>
        <div className="chips">
          <span className="chip selected">
            <Table2 size={13} />
            {selectedChip.label}
          </span>
          <span className="chip reference">@Sheet1</span>
          <button type="button" aria-label="添加上下文">
            <Plus size={14} />
          </button>
        </div>
        <textarea
          value={input}
          onChange={(event) => setInput(event.target.value)}
          placeholder="输入文字提问，或描述要新增的指标"
        />
        <div className="composer-actions">
          <div>
            <button type="button" aria-label="附件">
              <Paperclip size={16} />
            </button>
          </div>
          <button className="send" type="submit" aria-label="发送">
            <Send size={17} />
          </button>
        </div>
      </form>
    </aside>
  );
}

function MetricConfirmationCard({
  card,
  onChoice,
  onConfirm,
  onCancel,
}: {
  card: AiConfirmationCard;
  onChoice: (metricCode: string) => void;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  return (
    <section className={`confirm-card ${card.status}`}>
      <div className="card-title">
        <Sparkles size={16} />
        <strong>确认指标口径</strong>
        <span>{card.targetCell}</span>
      </div>
      <p>我不会直接写死数值。请先确认指标口径，确认后写入输入区，并把输出区标记为待刷新。</p>
      <div className="candidate-list">
        {card.candidates.map((candidate) => (
          <label className="candidate-option" key={candidate.code}>
            <input
              type="radio"
              checked={card.selectedMetricCode === candidate.code}
              onChange={() => onChoice(candidate.code)}
            />
            <span>
              <strong>{candidate.label}</strong>
              <small>{candidate.description}</small>
            </span>
          </label>
        ))}
      </div>
      {card.status === 'pending' && (
        <div className="patch-actions">
          <button onClick={onCancel}>取消</button>
          <button className="primary" onClick={onConfirm}>
            确认并写入 A13
          </button>
        </div>
      )}
    </section>
  );
}

function SaveTemplateCard({
  draft,
  appliedNote,
  onNameChange,
  onConfirm,
  onCancel,
}: {
  draft: SaveTemplateDraft;
  appliedNote: string;
  onNameChange: (value: string) => void;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  return (
    <section className={`confirm-card template-save ${draft.status}`}>
      <div className="card-title">
        <BookmarkPlus size={16} />
        <strong>保存为模板</strong>
        <span>{draft.status === 'saved' ? '已保存' : '待确认'}</span>
      </div>
      <label className="template-name-field">
        模板名称
        <input value={draft.templateName} onChange={(event) => onNameChange(event.target.value)} />
      </label>
      <div className="config-summary">
        <span>输入区：{draft.inputRanges.join(' / ')}</span>
        <span>输出区：{draft.outputRange}</span>
        <span>刷新配置：{draft.refreshRules}</span>
      </div>
      {draft.status === 'pending' && (
        <div className="patch-actions">
          <button onClick={onCancel}>取消</button>
          <button className="primary" onClick={onConfirm}>
            保存到模板广场
          </button>
        </div>
      )}
      {appliedNote && draft.status === 'saved' && <div className="applied-note">{appliedNote}</div>}
    </section>
  );
}

function AuditPanel({
  mode,
  outputConfigs,
  metricBindings,
  onClose,
}: {
  mode: 'audit' | 'export';
  outputConfigs: Record<string, OutputCellConfig>;
  metricBindings: MetricBinding[];
  onClose: () => void;
}) {
  const configs = Object.values(outputConfigs).slice(0, mode === 'audit' ? 12 : 18);

  return (
    <section className="formula-panel">
      <div className="formula-panel-title">
        <FunctionSquare size={16} />
        <strong>{mode === 'audit' ? '公式审计' : '导出表格'}</strong>
        <button type="button" onClick={onClose} aria-label="关闭面板">
          <X size={14} />
        </button>
      </div>
      <p>输出区不展示公式，只显示刷新后的数值。这里展示的是隐藏取数配置，用于审计、导出和回放。</p>
      <div className="role-summary">
        <div>
          <strong>输入区</strong>
          <span>A2:A13 指标，B1:H1 年份参数。</span>
        </div>
        <div>
          <strong>输出区</strong>
          <span>B2:H13 由隐藏配置刷新生成，只显示结果。</span>
        </div>
      </div>
      <div className="formula-list">
        {configs.map((config) => (
          <div className="formula-row" key={config.targetCell}>
            <span>{config.targetCell}</span>
            <code>
              metric={config.metricCode}; columnParamCell={config.columnParamCell}; status={statusText(config.status)}
            </code>
          </div>
        ))}
      </div>
      {mode === 'export' && (
        <div className="export-note">
          导出表格会包含可见数值，并附带 {metricBindings.length} 条指标绑定和隐藏取数配置。
        </div>
      )}
    </section>
  );
}

function TemplateGalleryModal({
  defaultTemplates,
  customTemplates,
  onClose,
}: {
  defaultTemplates: TemplateConfig[];
  customTemplates: TemplateConfig[];
  onClose: () => void;
}) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <section className="template-modal" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <div>
            <LayoutTemplate size={20} />
            <strong>模板广场</strong>
          </div>
          <button onClick={onClose}>
            <X size={16} />
          </button>
        </div>
        <div className="template-columns">
          <TemplateColumn title="默认模板" templates={defaultTemplates} emptyText="" />
          <TemplateColumn title="我的模板" templates={customTemplates} emptyText="还没有自定义模板。点击 Ribbon 的“保存为模板”后会出现在这里。" />
        </div>
      </section>
    </div>
  );
}

function TemplateColumn({ title, templates, emptyText }: { title: string; templates: TemplateConfig[]; emptyText: string }) {
  return (
    <div className="template-column">
      <h3>{title}</h3>
      {templates.length === 0 && <p className="empty-template">{emptyText}</p>}
      {templates.map((template) => (
        <article className="template-card" key={template.templateId}>
          <strong>{template.templateName}</strong>
          <p>{template.description}</p>
          <span>{template.inputRanges.join(' / ')} · {template.outputRange}</span>
        </article>
      ))}
    </div>
  );
}

function ContextMenu({ x, y, onRun }: { x: number; y: number; onRun: (label: string) => void }) {
  type ContextMenuAction = { icon: LucideIcon; label: string; trailing?: LucideIcon };
  type ContextMenuDivider = { type: 'divider' };

  const pasteOptions: ContextMenuAction[] = [
    { icon: ClipboardPaste, label: '粘贴' },
    { icon: FunctionSquare, label: '公式' },
    { icon: ClipboardCheck, label: '值' },
    { icon: SquarePen, label: '格式' },
    { icon: Table2, label: '数值格式' },
    { icon: Link, label: '链接' },
  ];

  const items: (ContextMenuAction | ContextMenuDivider)[] = [
    { icon: FilePlus2, label: '插入...' },
    { icon: Trash2, label: '删除...' },
    { icon: Eraser, label: '清除内容' },
    { type: 'divider' },
    { icon: Filter, label: '筛选' },
    { icon: ArrowUpDown, label: '排序', trailing: ChevronRight },
    { type: 'divider' },
    { icon: MessageSquarePlus, label: '插入批注' },
    { type: 'divider' },
    { icon: Settings2, label: '设置单元格格式...' },
    { icon: RefreshCw, label: '转换单元格...' },
    { icon: Link, label: '超链接...' },
    { type: 'divider' },
    { icon: SquarePen, label: '编辑富文本...' },
    { icon: Tag, label: '定义名称...' },
    { icon: BookmarkPlus, label: '标签...' },
    { icon: Clipboard, label: '默认值...' },
  ];

  return (
    <div className="context-menu" style={{ left: x, top: y }} onClick={(event) => event.stopPropagation()}>
      <div className="context-title">
        <Sparkles size={15} />
        DM AI for selected
      </div>
      <button onClick={() => onRun('剪切')}>
        <Scissors size={16} />
        剪切
      </button>
      <button onClick={() => onRun('复制')}>
        <Copy size={16} />
        复制
      </button>
      <div className="paste-header">
        <ClipboardPaste size={16} />
        粘贴选项:
      </div>
      <div className="paste-options" aria-label="粘贴选项">
        {pasteOptions.map(({ icon: Icon, label }) => (
          <button key={label} title={label} aria-label={label} onClick={() => onRun(label)}>
            <Icon size={18} />
          </button>
        ))}
      </div>
      {items.map((item, index) => {
        if (!('label' in item)) {
          return <div className="context-divider" key={`divider-${index}`} />;
        }

        const Icon = item.icon;
        const Trailing = item.trailing;

        return (
          <button key={item.label} onClick={() => onRun(item.label)}>
            <Icon size={16} />
            <span>{item.label}</span>
            {Trailing && <Trailing className="context-trailing" size={15} />}
          </button>
        );
      })}
    </div>
  );
}

function SheetFooter({
  selectedRange,
  sheets,
  activeSheetId,
  onSwitchSheet,
}: {
  selectedRange: string;
  sheets: WorkbookSheetState[];
  activeSheetId: string;
  onSwitchSheet: (sheetId: string) => void;
}) {
  const activeSheet = sheets.find((item) => item.sheetId === activeSheetId) ?? sheets[0];
  const customCount = activeSheet?.columnProfiles.filter((item) => item.role === 'custom_metric').length ?? 0;
  return (
    <footer className="sheet-footer">
      <div className="sheet-tabs">
        <button className="sheet-nav">‹</button>
        <button className="sheet-nav">›</button>
        {sheets.map((item) => (
          <button
            className={`sheet-tab ${item.sheetId === activeSheetId ? 'active' : ''}`}
            key={item.sheetId}
            onClick={() => onSwitchSheet(item.sheetId)}
          >
            {item.sheetName}
          </button>
        ))}
        <button className="sheet-plus">+</button>
      </div>
      <div className="status">
        <span>就绪</span>
        <span>选区: {selectedRange}</span>
        <span>{activeSheet ? objectTypeLabel(activeSheet.detectedObjectType) : '工作表'}</span>
        {customCount > 0 && <span>{customCount} 个自定义指标</span>}
        <span>100%</span>
      </div>
    </footer>
  );
}

function buildAiRequest(
  query: string,
  selection: CellSelection,
  activeWorkbook = workbook,
  activeSheet = sheet,
): AiRequest {
  return {
    query,
    workbook: activeWorkbook,
    sheet: activeSheet,
    selectedRange: selection,
    referencedContexts: query.includes('@Sheet1') ? contextReferences : [],
  };
}

function renderOutputValue(config: OutputCellConfig) {
  if (config.status === 'loading') return '刷新中';
  if (config.status === 'failed') return '失败';
  return config.value ?? '';
}

function statusText(status: OutputStatus) {
  return {
    empty: '空',
    loading: '刷新中',
    ready: '已刷新',
    failed: '失败',
  }[status];
}

function objectScopeText(scope: TemplateConfig['applicableObject']) {
  return {
    issuer: '主体',
    bond: '债券',
    bond_set: '债券集合',
  }[scope];
}

function objectTypeLabel(type: DetectedObjectType) {
  return {
    issuer: '主体表',
    bond: '债券表',
    unknown: '待确认',
  }[type];
}

function normalizeHeader(value: string) {
  return value.replace(/\s+/g, '').replace(/[()（）_\-—/\\]/g, '').toLowerCase();
}

function stringifyCell(value: string | number | Date | boolean | null | undefined) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function detectObjectType(rows: string[][], headerRowIndex: number): DetectedObjectType {
  const sample = rows.slice(headerRowIndex, headerRowIndex + 8).flat().join(' ');
  const normalizedSample = normalizeHeader(sample);
  const hasBondHeader = ['债券代码', '证券代码', '债券简称', '债券全称'].some((keyword) =>
    normalizedSample.includes(normalizeHeader(keyword)),
  );
  const hasBondCodeValue = /\b\d{6,9}\.(IB|SH|SZ|BJ|银行间)\b/i.test(sample);
  if (hasBondHeader || hasBondCodeValue) return 'bond';

  const hasIssuerHeader = ['主体名称', '发行人', '公司名称', '企业名称'].some((keyword) =>
    normalizedSample.includes(normalizeHeader(keyword)),
  );
  return hasIssuerHeader ? 'issuer' : 'unknown';
}

function findHeaderRow(rows: string[][]) {
  const keyWords = ['债券代码', '证券代码', '债券简称', '主体名称', '发行人', '公司名称', '估值日期', '报告期'];
  let bestIndex = rows.findIndex((row) => row.some(Boolean));
  let bestScore = -1;

  rows.slice(0, 8).forEach((row, index) => {
    const text = normalizeHeader(row.join(' '));
    const nonEmptyCount = row.filter(Boolean).length;
    const keywordScore = keyWords.filter((keyword) => text.includes(normalizeHeader(keyword))).length * 4;
    const score = keywordScore + Math.min(nonEmptyCount, 8);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return Math.max(bestIndex, 0);
}

function buildMetricDictionary(objectType: DetectedObjectType) {
  return defaultTemplates
    .filter((template) => {
      if (objectType === 'issuer') return template.applicableObject === 'issuer';
      if (objectType === 'bond') return template.applicableObject !== 'issuer';
      return true;
    })
    .flatMap((template) =>
      template.defaultOutputFields.map((fieldName) => ({
        normalized: normalizeHeader(fieldName),
        fieldName,
        metricCode: `${template.hiddenFetchConfig.interfaceName}.${template.hiddenFetchConfig.fieldMapping[fieldName] ?? fieldName}`,
        fetchInterface: template.hiddenFetchConfig.interfaceName,
      })),
    );
}

function classifyImportedColumn(
  header: string,
  index: number,
  objectType: DetectedObjectType,
): ImportedColumnProfile {
  const normalized = normalizeHeader(header);
  if (!normalized) return { index, header, role: 'ignored' };

  const bondKeys = ['债券代码', '证券代码', '债券简称', '债券全称', 'isin代码', '发行人名称'];
  const issuerKeys = ['主体名称', '发行人', '公司名称', '企业名称'];
  const entityKeys = objectType === 'bond' ? bondKeys : objectType === 'issuer' ? issuerKeys : [...bondKeys, ...issuerKeys];
  if (entityKeys.some((keyword) => normalized.includes(normalizeHeader(keyword)))) {
    return { index, header, role: 'entity_key' };
  }

  const match = buildMetricDictionary(objectType).find((item) => item.normalized === normalized);
  if (match) {
    return {
      index,
      header,
      role: 'recognized_metric',
      metricCode: match.metricCode,
      fetchInterface: match.fetchInterface,
      fieldName: match.fieldName,
    };
  }

  return { index, header, role: 'custom_metric' };
}

function padImportedRows(rows: string[][]) {
  const normalizedRows = rows.map((row) => columns.map((_, index) => row[index] ?? ''));
  while (normalizedRows.length < 55) {
    normalizedRows.push(Array.from({ length: columns.length }, () => ''));
  }
  return normalizedRows;
}

function buildImportedSheetState(
  rawRows: Array<Array<string | number | Date | boolean | null | undefined>>,
  sheetName: string,
  index: number,
  objectTypeOverride?: DetectedObjectType,
): ImportDraftSheet {
  const cleanedRows = rawRows
    .map((row) => row.map(stringifyCell))
    .filter((row) => row.some(Boolean));
  const headerRowIndex = findHeaderRow(cleanedRows);
  const objectType = objectTypeOverride ?? detectObjectType(cleanedRows, headerRowIndex);
  const rows = padImportedRows(cleanedRows.slice(headerRowIndex));
  const header = rows[0] ?? [];
  const columnProfiles = header.map((cell, columnIndex) => classifyImportedColumn(cell, columnIndex, objectType));
  const entityColumn = columnProfiles.find((profile) => profile.role === 'entity_key')?.index ?? 0;
  const outputConfigs: Record<string, OutputCellConfig> = {};
  const metricBindings: MetricBinding[] = [];

  columnProfiles.forEach((profile) => {
    if (profile.role !== 'recognized_metric' || !profile.metricCode) return;
    const metricCode = profile.metricCode;
    const metricColumn = columns[profile.index] ?? 'A';
    metricBindings.push({
      row: 1,
      cell: `${metricColumn}1`,
      displayName: profile.header,
      metricCode,
      source: profile.fetchInterface ?? 'DM_DATA',
      disambiguationStatus: 'confirmed',
    });
    rows.slice(1).forEach((row, rowIndex) => {
      if (!row.some(Boolean)) return;
      const targetCell = makeCellAddress(profile.index + 1, rowIndex + 2);
      outputConfigs[targetCell] = {
        targetCell,
        metricCell: `${metricColumn}1`,
        columnParamCell: `${metricColumn}1`,
        metricCode,
        columnParam: profile.fieldName ?? profile.header,
        status: 'ready',
        value: row[profile.index] ?? '',
        fetchInterface: profile.fetchInterface,
        fieldName: profile.fieldName ?? profile.header,
        parameterCell: makeCellAddress(entityColumn + 1, rowIndex + 2),
      };
    });
  });

  return {
    sheetId: `import_sheet_${index + 1}`,
    sheetName: sheetName || `Sheet${index + 1}`,
    rows,
    years: header.slice(1, 8),
    outputConfigs,
    metricDrafts: {},
    metricBindings,
    inputCellModes: {},
    headerColumnModes: Object.fromEntries(
      columnProfiles.slice(1, 8).map((profile) => [
        columns[profile.index] ?? 'B',
        profile.role === 'recognized_metric' ? ('system' as HeaderColumnMode) : ('custom' as HeaderColumnMode),
      ]),
    ),
    columnParamType: 'fieldCode',
    activeColumnCount: Math.min(columns.length, Math.max(8, header.filter(Boolean).length || 8)),
    detectedObjectType: objectType,
    columnProfiles,
    previewRows: rows.slice(0, 5),
    customMetricCount: columnProfiles.filter((profile) => profile.role === 'custom_metric').length,
  };
}

function rebuildImportedSheetWithObjectType(sheetState: ImportDraftSheet, objectType: DetectedObjectType) {
  const rebuilt = buildImportedSheetState(sheetState.rows, sheetState.sheetName, Number(sheetState.sheetId.replace(/\D/g, '')) || 0, objectType);
  return {
    ...rebuilt,
    sheetId: sheetState.sheetId,
    sheetName: sheetState.sheetName,
  };
}

function filterTemplates(templates: TemplateConfig[], search: string) {
  const query = search.trim().toLowerCase();
  if (!query) return templates;
  return templates.filter((template) => {
    const haystack = [
      template.templateName,
      template.description,
      objectScopeText(template.applicableObject),
      ...template.requiredParams.map((param) => `${param.name} ${param.description}`),
      ...template.optionalParams.map((param) => `${param.name} ${param.description}`),
      ...template.defaultOutputFields,
      template.hiddenFetchConfig.interfaceName,
    ]
      .join(' ')
      .toLowerCase();
    return haystack.includes(query);
  });
}

function buildTemplateWorkbook(template: TemplateConfig) {
  const headers = template.defaultOutputFields.slice(0, 7);
  const rows: string[][] = Array.from({ length: 55 }, () => Array.from({ length: 16 }, () => ''));
  const metricDrafts: Record<string, string> = {};
  const outputConfigs: Record<string, OutputCellConfig> = {};
  const slots = [
    ...template.requiredParams.map((param) => ({ ...param, required: true })),
    ...template.optionalParams.map((param) => ({ ...param, required: false })),
  ].slice(0, 12);

  rows[0] = ['参数槽位', ...headers, ...Array.from({ length: Math.max(0, 15 - headers.length) }, () => '')];
  slots.forEach((slot, index) => {
    const rowIndex = index + 1;
    const cell = `A${rowIndex + 1}`;
    rows[rowIndex][0] = `${slot.required ? '必填' : '可选'}｜${slot.name}`;
    metricDrafts[cell] = slot.placeholder;
    headers.forEach((fieldName, headerIndex) => {
      const columnIndex = headerIndex + 2;
      const targetCell = makeCellAddress(columnIndex, rowIndex + 1);
      const fieldKey = template.hiddenFetchConfig.fieldMapping[fieldName] ?? fieldName;
      outputConfigs[targetCell] = {
        targetCell,
        metricCell: cell,
        columnParamCell: `${columns[columnIndex - 1]}1`,
        metricCode: `${template.hiddenFetchConfig.interfaceName}.${fieldKey}`,
        columnParam: fieldName,
        status: 'ready',
        value: mockFetchMetricValue(`${template.hiddenFetchConfig.interfaceName}.${fieldKey}`, fieldName),
        lastRefreshAt: new Date().toISOString().slice(0, 16).replace('T', ' '),
        fetchInterface: template.hiddenFetchConfig.interfaceName,
        fieldName,
        parameterCell: cell,
      };
    });
  });

  rows[slots.length + 2][0] = '隐藏取数接口';
  rows[slots.length + 2][1] = template.hiddenFetchConfig.interfaceName;
  rows[slots.length + 3][0] = '参数映射';
  rows[slots.length + 3][1] = Object.entries(template.hiddenFetchConfig.parameterMapping)
    .map(([key, cell]) => `${key}=${cell}`)
    .join('; ');

  return { rows, headers, outputConfigs, metricDrafts };
}

function moduleStageCopy(moduleId: string, index: number) {
  const copies: Record<string, string[]> = {
    'research-view': [
      '支持研报、纪要、公告等材料进入统一收件箱。',
      '抽取多空观点、时间判断、品种偏好和风险提示。',
      '把观点背后的数据、图表和原文段落保留为证据链。',
      '沉淀为可筛选、可引用、可回填的观点卡片。',
    ],
    'credit-analysis': [
      '从主体、债券代码或持仓清单进入分析流程。',
      '汇总财务、估值、成交、评级、舆情和条款信息。',
      '识别利差走阔、评级调整、负面舆情等变化。',
      '输出主体画像、风险结论和后续跟踪动作。',
    ],
    'quant-backtest': [
      '用自然语言或表格条件描述策略规则。',
      '选择利率债、信用债、存单等资产池和调仓频率。',
      '运行历史区间回测，记录净值、换手和回撤。',
      '拆解久期、票息、利差、择券和交易贡献。',
    ],
  };

  return copies[moduleId]?.[index] ?? '这里会承接该能力的核心任务、过程状态和输出结果。';
}

function isMetricInputCell(address: string) {
  const { row, column } = parseCell(address);
  return column === 1 && row >= 2 && row <= 13;
}

function getSuggestionContext(address: string, activeColumnCount: number): SuggestionContext | null {
  const { row, column } = parseCell(address);
  if (column === 1 && row >= 2 && row <= 13) return 'metricRow';
  if (row === 1 && column >= 2 && column <= activeColumnCount) return 'headerCol';
  return null;
}

function isSelectedCell(row: number, column: number, range: string) {
  const [start, end = start] = range.split(':');
  const startCoord = parseCell(start);
  const endCoord = parseCell(end);

  return (
    row >= Math.min(startCoord.row, endCoord.row) &&
    row <= Math.max(startCoord.row, endCoord.row) &&
    column >= Math.min(startCoord.column, endCoord.column) &&
    column <= Math.max(startCoord.column, endCoord.column)
  );
}

function parseCell(cell: string) {
  const match = cell.match(/^([A-Z]+)(\d+)$/);
  if (!match) return { column: 1, row: 1 };
  const column = match[1].split('').reduce((sum, letter) => sum * 26 + letter.charCodeAt(0) - 64, 0);
  return { column, row: Number(match[2]) };
}

function getCellValue(rows: string[][], address: string) {
  const { row, column } = parseCell(address);
  return rows[row - 1]?.[column - 1] ?? '';
}

export { App };

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
  X,
} from 'lucide-react';
import type { LucideIcon } from 'lucide-react';
import { FormEvent, MouseEvent, useMemo, useState } from 'react';
import {
  baseRows,
  buildInitialOutputConfigs,
  columns,
  defaultTemplates,
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

const initialSelection: CellSelection = {
  sheetId: workbook.activeSheetId,
  sheetName: workbook.activeSheetName,
  address: 'A13',
  activeCell: 'A13',
};

const contextReferences: AiContextChip[] = [{ id: 'ref_sheet1', type: 'reference', label: '@Sheet1' }];

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

function App() {
  const [gridRows, setGridRows] = useState(baseRows);
  const [years, setYears] = useState(baseRows[0].slice(1, 8));
  const [metricBindings, setMetricBindings] = useState<MetricBinding[]>(initialMetricBindings);
  const [outputConfigs, setOutputConfigs] = useState<Record<string, OutputCellConfig>>(() =>
    buildInitialOutputConfigs(baseRows[0].slice(1, 8), initialMetricBindings),
  );
  const [selection, setSelection] = useState<CellSelection>(initialSelection);
  const [menu, setMenu] = useState<MenuState>(null);
  const [input, setInput] = useState('生猪价格');
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'assistant_intro',
      role: 'assistant',
      text: '我已绑定当前工作簿。输入区用于放指标和年份，输出区只显示刷新后的结果；你可以右键选区，也可以在这里用自然语言告诉我想新增什么指标。',
    },
  ]);
  const [confirmationCard, setConfirmationCard] = useState<AiConfirmationCard | null>(null);
  const [saveDraft, setSaveDraft] = useState<SaveTemplateDraft | null>(null);
  const [customTemplates, setCustomTemplates] = useState<TemplateConfig[]>([]);
  const [templateModalOpen, setTemplateModalOpen] = useState(false);
  const [formulaPanelMode, setFormulaPanelMode] = useState<'audit' | 'export' | null>(null);
  const [appliedNote, setAppliedNote] = useState('');
  const [activeSuggestionCell, setActiveSuggestionCell] = useState<string | null>(null);
  const [metricDrafts, setMetricDrafts] = useState<Record<string, string>>({});
  const [activeModuleId, setActiveModuleId] = useState(productModules[0].id);

  const displayRows = useMemo(() => {
    const rows = gridRows.map((row) => [...row]);
    rows[0] = ['指标', ...years];
    Object.values(outputConfigs).forEach((config) => {
      const { row, column } = parseCell(config.targetCell);
      rows[row - 1] = [...(rows[row - 1] ?? [])];
      rows[row - 1][column - 1] = renderOutputValue(config);
    });
    return rows;
  }, [gridRows, outputConfigs, years]);

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
      return `隐藏取数配置: DMDATA(${config.metricCode}, ${config.yearCell})`;
    }
    return getCellValue(displayRows, selection.activeCell);
  }, [displayRows, outputConfigs, selection.activeCell]);

  const activeModule = useMemo(
    () => productModules.find((module) => module.id === activeModuleId) ?? productModules[0],
    [activeModuleId],
  );

  const handleSelect = (address: string) => {
    setSelection({
      sheetId: workbook.activeSheetId,
      sheetName: workbook.activeSheetName,
      address,
      activeCell: address.split(':')[0],
    });
    setActiveSuggestionCell(isMetricInputCell(address) ? address : null);
  };

  const updateYear = (columnIndex: number, value: string) => {
    const cleanYear = value.replace(/[^\d]/g, '').slice(0, 4);
    const nextYears = [...years];
    nextYears[columnIndex - 2] = cleanYear;
    setYears(nextYears);

    setOutputConfigs((current) => {
      const next = { ...current };
      Object.values(next).forEach((config) => {
        if (config.yearCell === `${columns[columnIndex - 1]}1`) {
          config.year = cleanYear;
          config.status = config.metricCode && cleanYear ? 'pendingRefresh' : 'empty';
          config.value = '';
          config.lastRefreshAt = undefined;
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
          yearCell: `${columns[columnIndex - 1]}1`,
          metricCode: metric.code,
          year,
          status: year ? 'pendingRefresh' : 'empty',
          value: '',
        };
      });
      return next;
    });
    setActiveSuggestionCell(null);
    setMetricDrafts((current) => ({ ...current, [`A${row}`]: metric.label }));
  };

  const refreshOutputs = () => {
    const now = '2026-05-07 15:46';
    setOutputConfigs((current) => {
      const loading: Record<string, OutputCellConfig> = {};
      Object.entries(current).forEach(([cell, config]) => {
        loading[cell] = config.status === 'pendingRefresh' ? { ...config, status: 'loading' } : config;
      });
      setTimeout(() => {
        setOutputConfigs((latest) => {
          const refreshed: Record<string, OutputCellConfig> = {};
          Object.entries(latest).forEach(([cell, config]) => {
            if (config.status === 'loading') {
              refreshed[cell] = {
                ...config,
                status: 'ready',
                value: mockFetchMetricValue(config.metricCode, config.year),
                lastRefreshAt: now,
              };
            } else {
              refreshed[cell] = config;
            }
          });
          return refreshed;
        });
      }, 450);
      return loading;
    });
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
    const request = buildAiRequest(query, selected);
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
      sheetId: workbook.activeSheetId,
      sheetName: workbook.activeSheetName,
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

  const openContextMenu = (event: MouseEvent, address = selection.address) => {
    event.preventDefault();
    setMenu({
      x: event.clientX,
      y: event.clientY,
      selection: {
        sheetId: workbook.activeSheetId,
        sheetName: workbook.activeSheetName,
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

  return (
    <div className="terminal-shell" onClick={() => setMenu(null)}>
      <TopBar activeModuleId={activeModuleId} modules={productModules} onModuleSelect={setActiveModuleId} />
      <div className="workspace">
        <LeftRail
          activeModule={activeModule}
          onOpenTemplates={() => setTemplateModalOpen(true)}
          customTemplateCount={customTemplates.length}
        />
        {activeModule.id === 'data-assistant' ? (
          <main className="excel-pane">
            <WorkbookHeader />
            <Ribbon
              onRefresh={refreshOutputs}
              onSaveTemplate={requestSaveTemplate}
              onShowAudit={() => setFormulaPanelMode('audit')}
              onExportTable={() => setFormulaPanelMode('export')}
            />
            <FormulaBar activeCell={selection.activeCell} value={selectedFormulaValue} />
            <SpreadsheetGrid
              rows={displayRows}
              selection={selection}
              years={years}
              outputConfigs={outputConfigs}
              activeSuggestionCell={activeSuggestionCell}
              metricDrafts={metricDrafts}
              onSelect={handleSelect}
              onYearChange={updateYear}
              onMetricDraftChange={(cell, value) => setMetricDrafts((current) => ({ ...current, [cell]: value }))}
              onBindMetric={bindMetricToRow}
              onContextMenu={openContextMenu}
            />
            <SheetFooter selectedRange={selection.address} />
          </main>
        ) : (
          <ModuleWorkspace module={activeModule} />
        )}
        <AiPanel
          input={input}
          setInput={setInput}
          selectedChip={selectedChip}
          messages={messages}
          confirmationCard={confirmationCard}
          saveDraft={saveDraft}
          appliedNote={appliedNote}
          outputConfigs={outputConfigs}
          metricBindings={metricBindings}
          formulaPanelMode={formulaPanelMode}
          onSubmit={handleSubmit}
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
      </div>
      {menu && <ContextMenu x={menu.x} y={menu.y} onRun={runMenuAction} />}
      {templateModalOpen && (
        <TemplateGalleryModal
          defaultTemplates={defaultTemplates}
          customTemplates={customTemplates}
          onClose={() => setTemplateModalOpen(false)}
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
  customTemplateCount,
}: {
  activeModule: ProductModule;
  onOpenTemplates: () => void;
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
              <Plus size={14} />
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

function WorkbookHeader() {
  return (
    <section className="workbook-header">
      <div className="tabs">
        <div className="file-tab active">
          <FileSpreadsheet size={14} />
          海底捞利润表_2021-2025.xlsx
          <X size={13} />
        </div>
        <div className="file-tab">
          107_编辑在线文档与Excel.md
          <X size={13} />
        </div>
        <button className="tab-plus">+</button>
      </div>
      <div className="pathline">
        tree &gt; ... &gt; 海底捞利润表_2021-2025.xlsx
        <span>最近修改: 05月07 15:46</span>
      </div>
    </section>
  );
}

function Ribbon({
  onRefresh,
  onSaveTemplate,
  onShowAudit,
  onExportTable,
}: {
  onRefresh: () => void;
  onSaveTemplate: () => void;
  onShowAudit: () => void;
  onExportTable: () => void;
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
  activeSuggestionCell,
  metricDrafts,
  onSelect,
  onYearChange,
  onMetricDraftChange,
  onBindMetric,
  onContextMenu,
}: {
  rows: string[][];
  selection: CellSelection;
  years: string[];
  outputConfigs: Record<string, OutputCellConfig>;
  activeSuggestionCell: string | null;
  metricDrafts: Record<string, string>;
  onSelect: (address: string) => void;
  onYearChange: (columnIndex: number, value: string) => void;
  onMetricDraftChange: (cell: string, value: string) => void;
  onBindMetric: (row: number, metric: MetricDefinition) => void;
  onContextMenu: (event: MouseEvent, address?: string) => void;
}) {
  const rowCount = 55;
  const activeSuggestionRow = activeSuggestionCell ? parseCell(activeSuggestionCell).row : 0;
  const activeQuery = activeSuggestionCell ? metricDrafts[activeSuggestionCell] ?? getCellValue(rows, activeSuggestionCell) : '';
  const suggestions = activeSuggestionCell ? findMetricCandidates(activeQuery) : [];

  return (
    <section className="grid-wrap" onContextMenu={(event) => onContextMenu(event, selection.address)}>
      <table className="sheet-grid">
        <thead>
          <tr>
            <th className="corner" />
            {columns.map((column) => (
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
                {columns.map((column, columnIndex) => {
                  const address = `${column}${rowNumber}`;
                  const value = row[columnIndex] ?? '';
                  const selected = isSelectedCell(rowNumber, columnIndex + 1, selection.address);
                  const isHeader = rowNumber === 1 && columnIndex < 8;
                  const isMetricInput = columnIndex === 0 && rowNumber >= 2 && rowNumber <= 13;
                  const isYearInput = rowNumber === 1 && columnIndex >= 1 && columnIndex < 8;
                  const outputConfig = outputConfigs[address];
                  const isOutput = Boolean(outputConfig);

                  return (
                    <td
                      className={[
                        selected ? 'selected-cell' : '',
                        isHeader ? 'table-header-cell' : '',
                        isMetricInput || isYearInput ? 'input-cell' : '',
                        isYearInput ? 'year-input-cell' : '',
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
                      {isYearInput ? (
                        <input
                          value={years[columnIndex - 1] ?? ''}
                          onChange={(event) => onYearChange(columnIndex + 1, event.target.value)}
                          aria-label={`${address} 年份`}
                        />
                      ) : isMetricInput ? (
                        <input
                          className="metric-input"
                          value={metricDrafts[address] ?? value}
                          placeholder={rowNumber === 13 ? '输入或搜索指标' : ''}
                          onChange={(event) => onMetricDraftChange(address, event.target.value)}
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
      {activeSuggestionCell && (
        <div className="suggestion-popover" style={{ top: 78 + activeSuggestionRow * 22 }}>
          <div className="suggestion-title">
            <Search size={14} />
            可输入内容 / 指标联想
          </div>
          {suggestions.map((metric) => (
            <button key={metric.code} onClick={() => onBindMetric(activeSuggestionRow, metric)}>
              <strong>{metric.label}</strong>
              <span>{metric.category} · {metric.unit} · {metric.source}</span>
              <small>{metric.description}</small>
            </button>
          ))}
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
  confirmationCard,
  saveDraft,
  appliedNote,
  outputConfigs,
  metricBindings,
  formulaPanelMode,
  onSubmit,
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
  confirmationCard: AiConfirmationCard | null;
  saveDraft: SaveTemplateDraft | null;
  appliedNote: string;
  outputConfigs: Record<string, OutputCellConfig>;
  metricBindings: MetricBinding[];
  formulaPanelMode: 'audit' | 'export' | null;
  onSubmit: (event: FormEvent) => void;
  onMetricChoice: (metricCode: string) => void;
  onConfirmMetric: () => void;
  onCancelMetric: () => void;
  onTemplateNameChange: (value: string) => void;
  onConfirmSaveTemplate: () => void;
  onCancelSaveTemplate: () => void;
  onCloseFormulaPanel: () => void;
}) {
  return (
    <aside className="ai-panel">
      <header className="ai-header">
        <div>
          <Bot size={18} />
          <strong>DM AI</strong>
        </div>
        <div className="panel-icons">
          <Plus size={15} />
          <Bell size={15} />
          <PanelRightClose size={16} />
          <X size={16} />
        </div>
      </header>
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
              metric={config.metricCode}; yearCell={config.yearCell}; status={statusText(config.status)}
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

function SheetFooter({ selectedRange }: { selectedRange: string }) {
  return (
    <footer className="sheet-footer">
      <div className="sheet-tabs">
        <button className="sheet-nav">‹</button>
        <button className="sheet-nav">›</button>
        <button className="sheet-tab active">Sheet1</button>
        <button className="sheet-plus">+</button>
      </div>
      <div className="status">
        <span>就绪</span>
        <span>选区: {selectedRange}</span>
        <span>输入区 A2:A13</span>
        <span>100%</span>
      </div>
    </footer>
  );
}

function buildAiRequest(query: string, selection: CellSelection): AiRequest {
  return {
    query,
    workbook,
    sheet,
    selectedRange: selection,
    referencedContexts: query.includes('@Sheet1') ? contextReferences : [],
  };
}

function renderOutputValue(config: OutputCellConfig) {
  if (config.status === 'pendingRefresh') return '待刷新';
  if (config.status === 'loading') return '刷新中';
  if (config.status === 'failed') return '失败';
  return config.value ?? '';
}

function statusText(status: OutputStatus) {
  return {
    empty: '空',
    pendingRefresh: '待刷新',
    loading: '刷新中',
    ready: '已刷新',
    failed: '失败',
  }[status];
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

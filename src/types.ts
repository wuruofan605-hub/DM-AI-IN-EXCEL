export type OutputStatus = 'empty' | 'pendingRefresh' | 'loading' | 'ready' | 'failed';

export type MetricDefinition = {
  code: string;
  label: string;
  category: string;
  unit: string;
  source: string;
  description: string;
  aliases: string[];
};

export type MetricBinding = {
  row: number;
  cell: string;
  displayName: string;
  metricCode: string;
  source: string;
  disambiguationStatus: 'confirmed' | 'needsConfirmation';
};

export type OutputCellConfig = {
  targetCell: string;
  metricCell: string;
  yearCell: string;
  metricCode: string;
  year: string;
  status: OutputStatus;
  lastRefreshAt?: string;
  value?: string;
  fetchInterface?: string;
  fieldName?: string;
  parameterCell?: string;
};

export type TemplateObjectScope = 'issuer' | 'bond' | 'bond_set';

export type TemplateParamSlot = {
  name: string;
  description: string;
  placeholder: string;
};

export type HiddenFetchConfig = {
  interfaceName: string;
  parameterMapping: Record<string, string>;
  fieldMapping: Record<string, string>;
};

export type TemplateConfig = {
  templateId: string;
  templateName: string;
  description: string;
  inputRanges: string[];
  outputRange: string;
  metricBindings: MetricBinding[];
  refreshRules: string;
  formatPreset: string;
  type: 'default' | 'custom';
  applicableObject: TemplateObjectScope;
  requiredParams: TemplateParamSlot[];
  optionalParams: TemplateParamSlot[];
  defaultOutputFields: string[];
  hiddenFetchConfig: HiddenFetchConfig;
  previewRows: string[][];
};

export type AiContextChip = {
  id: string;
  type: 'selected' | 'reference';
  label: string;
};

export type CellSelection = {
  sheetId: string;
  sheetName: string;
  address: string;
  activeCell: string;
};

export type WorkbookContext = {
  workbookId: string;
  workbookName: string;
  activeSheetId: string;
  activeSheetName: string;
  visibleRange: string;
  usedRange: string;
};

export type SheetContext = {
  sheetId: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
};

export type AiRequest = {
  query: string;
  workbook: WorkbookContext;
  sheet: SheetContext;
  selectedRange: CellSelection;
  referencedContexts: AiContextChip[];
};

export type ChatMessage = {
  id: string;
  role: 'user' | 'assistant';
  text: string;
  request?: AiRequest;
};

export type AiConfirmationCard = {
  id: string;
  type: 'metricBinding';
  status: 'pending' | 'confirmed' | 'cancelled';
  targetCell: string;
  query: string;
  candidates: MetricDefinition[];
  selectedMetricCode: string;
};

export type SaveTemplateDraft = {
  id: string;
  status: 'pending' | 'saved' | 'cancelled';
  templateName: string;
  inputRanges: string[];
  outputRange: string;
  refreshRules: string;
  formatPreset: string;
};

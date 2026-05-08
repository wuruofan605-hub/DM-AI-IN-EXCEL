import type { MetricBinding, MetricDefinition, OutputCellConfig, TemplateConfig } from './types';

export const columns = 'ABCDEFGHIJKLMNOP'.split('');

export const workbook = {
  workbookId: 'wb_dm_macro_2021_2026',
  workbookName: '海底捞利润表_2021-2025.xlsx',
  activeSheetId: 'sheet_1',
  activeSheetName: 'Sheet1',
  visibleRange: 'A1:P55',
  usedRange: 'A1:H13',
};

export const sheet = {
  sheetId: 'sheet_1',
  sheetName: 'Sheet1',
  rowCount: 55,
  columnCount: 16,
};

export const baseRows = [
  ['指标', '2019', '2021', '2022', '2023', '2024', '2025', '2026'],
  ['GDP现价总量（万亿元）', '99.09', '114.92', '121.02', '126.06', '134.91', '140.19', ''],
  ['GDP实际增速（%）', '6.1', '8.4', '3.0', '5.2', '5.0', '5.0', ''],
  ['CPI累计同比（%）', '2.9', '0.9', '2.0', '0.2', '0.2', '0.0', ''],
  ['PPI全部工业品同比（%）', '-0.3', '8.1', '4.1', '-3.0', '-2.2', '-2.6', ''],
  ['食品烟酒CPI同比（%）', '5.5', '-0.3', '2.4', '0.3', '-0.1', '-0.7', ''],
  ['城镇调查失业率（%）', '5.2', '5.5', '5.5', '5.2', '5.1', '5.1', ''],
  ['社会消费品零售总额（亿元）', '411649', '438352', '436449', '467098', '483345', '501202', ''],
  ['餐饮收入社零增速（%）', '9.4', '18.6', '-6.3', '20.4', '3.8', '2.1', ''],
  ['人均可支配收入（元）', '30733', '35128', '36883', '39218', '41314', '43377', ''],
  ['居民收入中位数（元）', '26523', '29975', '31370', '33036', '34707', '36231', ''],
  ['', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', ''],
];

export const metricDefinitions: MetricDefinition[] = [
  {
    code: 'CN.GDP.NOMINAL',
    label: 'GDP现价总量（万亿元）',
    category: '宏观总量',
    unit: '万亿元',
    source: 'DM 宏观数据库 / 国家统计局',
    description: '按现价核算的国内生产总值，适合做年度总量横向比较。',
    aliases: ['gdp', '国内生产总值', 'GDP现价'],
  },
  {
    code: 'CN.GDP.REAL_YOY',
    label: 'GDP实际增速（%）',
    category: '宏观总量',
    unit: '%',
    source: 'DM 宏观数据库 / 国家统计局',
    description: '剔除价格因素后的 GDP 同比增速。',
    aliases: ['实际GDP', 'GDP增速', '经济增速'],
  },
  {
    code: 'CN.CPI.YOY',
    label: 'CPI累计同比（%）',
    category: '通胀',
    unit: '%',
    source: 'DM 宏观数据库 / 国家统计局',
    description: '居民消费价格指数累计同比。',
    aliases: ['cpi', '通胀', '居民消费价格'],
  },
  {
    code: 'CN.PPI.YOY',
    label: 'PPI全部工业品同比（%）',
    category: '通胀',
    unit: '%',
    source: 'DM 宏观数据库 / 国家统计局',
    description: '工业生产者出厂价格指数同比。',
    aliases: ['ppi', '工业品价格', '生产者价格'],
  },
  {
    code: 'CN.HOG.PRICE.NATIONAL',
    label: '生猪价格：全国均价（元/公斤）',
    category: '农产品价格',
    unit: '元/公斤',
    source: 'DM 农产品数据库',
    description: '全国生猪平均价格，适合做宏观消费和养殖利润跟踪。',
    aliases: ['生猪价格', '猪价', '生猪', '猪肉价格'],
  },
  {
    code: 'CN.HOG.PRICE.OUTER3',
    label: '生猪价格：外三元（元/公斤）',
    category: '农产品价格',
    unit: '元/公斤',
    source: 'DM 农产品数据库',
    description: '外三元生猪价格，市场交易跟踪常用口径。',
    aliases: ['外三元', '生猪价格', '猪价'],
  },
  {
    code: 'CN.HOG.PRICE.22CITY',
    label: '22省市生猪平均价（元/公斤）',
    category: '农产品价格',
    unit: '元/公斤',
    source: 'DM 农产品数据库',
    description: '22省市生猪平均价，适合做区域均价观察。',
    aliases: ['22省市', '生猪均价', '猪价'],
  },
];

const rowMetricCodes = [
  'CN.GDP.NOMINAL',
  'CN.GDP.REAL_YOY',
  'CN.CPI.YOY',
  'CN.PPI.YOY',
  'CN.CPI.FOOD_TOBACCO',
  'CN.URBAN.UNEMPLOYMENT',
  'CN.RETAIL.SALES',
  'CN.CATERING.RETAIL.YOY',
  'CN.INCOME.DISPOSABLE',
  'CN.INCOME.MEDIAN',
];

export const initialMetricBindings: MetricBinding[] = rowMetricCodes.map((metricCode, index) => {
  const row = index + 2;
  const definition = metricDefinitions.find((item) => item.code === metricCode);
  return {
    row,
    cell: `A${row}`,
    displayName: baseRows[index + 1][0],
    metricCode,
    source: definition?.source ?? 'DM 宏观数据库',
    disambiguationStatus: 'confirmed',
  };
});

export const defaultTemplates: TemplateConfig[] = [
  {
    templateId: 'tpl_macro_ts',
    templateName: '宏观指标时间序列',
    description: '按年份横向展开 GDP、CPI、PPI、社零等宏观指标。',
    inputRanges: ['A2:A13', 'B1:H1'],
    outputRange: 'B2:H13',
    metricBindings: initialMetricBindings,
    refreshRules: '按年份列批量刷新，输出区只显示结果。',
    formatPreset: '淡黄输入区 + 淡蓝输出区',
    type: 'default',
  },
  {
    templateId: 'tpl_company_compare',
    templateName: '公司财务指标对比',
    description: '对比公司收入、利润率、ROE、现金流等指标。',
    inputRanges: ['A2:A20', 'B1:H1'],
    outputRange: 'B2:H20',
    metricBindings: [],
    refreshRules: '按公司和期间刷新财务指标。',
    formatPreset: '金融终端紧凑表格',
    type: 'default',
  },
  {
    templateId: 'tpl_bond_monitor',
    templateName: '债券估值监控',
    description: '跟踪收益率、久期、估值净价和利差。',
    inputRanges: ['A2:A20', 'B1:H1'],
    outputRange: 'B2:H20',
    metricBindings: [],
    refreshRules: '按债券代码和估值日刷新。',
    formatPreset: '报价监控表',
    type: 'default',
  },
];

export function makeCellAddress(columnIndex: number, row: number) {
  return `${columns[columnIndex - 1]}${row}`;
}

export function findMetricCandidates(query: string) {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return metricDefinitions.slice(0, 5);

  return metricDefinitions
    .filter((metric) => {
      const haystack = [metric.label, metric.code, metric.category, ...metric.aliases].join(' ').toLowerCase();
      return haystack.includes(normalized);
    })
    .slice(0, 6);
}

export function mockFetchMetricValue(metricCode: string, year: string) {
  const numericYear = Number(year || 0);
  const seed = metricCode.split('').reduce((sum, char) => sum + char.charCodeAt(0), 0);
  if (!numericYear) return '';

  if (metricCode === 'CN.GDP.NOMINAL') return (132 + (numericYear - 2024) * 5.7).toFixed(2);
  if (metricCode === 'CN.GDP.REAL_YOY') return (4.6 + ((numericYear + seed) % 5) * 0.2).toFixed(1);
  if (metricCode === 'CN.CPI.YOY') return (0.2 + ((numericYear + seed) % 8) * 0.25).toFixed(1);
  if (metricCode === 'CN.PPI.YOY') return (-2.8 + ((numericYear + seed) % 10) * 0.35).toFixed(1);
  if (metricCode.includes('HOG.PRICE')) return (14.8 + ((numericYear + seed) % 7) * 0.65).toFixed(2);
  if (metricCode.includes('INCOME')) return String(33000 + (numericYear - 2023) * 1850 + (seed % 500));
  if (metricCode.includes('RETAIL')) return String(470000 + (numericYear - 2023) * 17000 + (seed % 2000));
  return (2 + ((numericYear + seed) % 10) * 0.7).toFixed(1);
}

export function buildInitialOutputConfigs(years: string[], bindings: MetricBinding[]): Record<string, OutputCellConfig> {
  const configs: Record<string, OutputCellConfig> = {};
  bindings.forEach((binding) => {
    years.forEach((year, index) => {
      const columnIndex = index + 2;
      const targetCell = makeCellAddress(columnIndex, binding.row);
      configs[targetCell] = {
        targetCell,
        metricCell: `A${binding.row}`,
        yearCell: `${columns[columnIndex - 1]}1`,
        metricCode: binding.metricCode,
        year,
        status: columnIndex === 8 ? 'pendingRefresh' : 'ready',
        value: columnIndex === 8 ? '' : baseRows[binding.row - 1]?.[columnIndex - 1] ?? '',
        lastRefreshAt: columnIndex === 8 ? undefined : '2026-05-07 15:46',
      };
    });
  });
  return configs;
}

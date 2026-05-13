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
    templateId: 'tpl_bond_basic_info',
    templateName: '债券基础信息查询',
    description: '输入债券代码后，拉取债券核心静态资料，适合做单券尽调和台账补全。',
    inputRanges: ['A2:A4'],
    outputRange: 'B2:H8',
    metricBindings: [],
    refreshRules: '按债券代码刷新债券静态字段，输出区只显示字段结果。',
    formatPreset: '债券信息表',
    type: 'default',
    applicableObject: 'bond',
    requiredParams: [
      { name: '债券代码', description: '交易所或银行间债券代码。', placeholder: '例如 102300123.IB' },
    ],
    optionalParams: [
      { name: '指标组', description: '可限定只看基础、发行或兑付字段。', placeholder: '基础信息' },
    ],
    defaultOutputFields: ['债券简称', '发行人', '债券类型', '起息日', '到期日', '余额', '评级'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_PROFILE',
      parameterMapping: { bondCode: 'A2', fieldSet: 'A3' },
      fieldMapping: {
        债券简称: 'bondShortName',
        发行人: 'issuerName',
        债券类型: 'bondType',
        起息日: 'valueDate',
        到期日: 'maturityDate',
        余额: 'outstandingBalance',
        评级: 'bondRating',
      },
    },
    previewRows: [
      ['债券代码', '102300123.IB', ''],
      ['债券简称', '23示例债01', '发行人'],
      ['余额', '18.50亿', '评级 AAA'],
    ],
  },
  {
    templateId: 'tpl_bond_balance',
    templateName: '债券余额查询',
    description: '按债券代码和估值日查询余额、剩余期限和托管量，适合持仓和存续规模核对。',
    inputRanges: ['A2:A6'],
    outputRange: 'B2:F8',
    metricBindings: [],
    refreshRules: '按债券代码和估值日刷新余额类字段。',
    formatPreset: '余额核对表',
    type: 'default',
    applicableObject: 'bond_set',
    requiredParams: [
      { name: '债券代码', description: '支持单只债券或一组债券代码。', placeholder: '粘贴债券代码列表' },
      { name: '估值日', description: '余额和托管量对应的日期。', placeholder: '2026-05-12' },
    ],
    optionalParams: [
      { name: '债券类型', description: '筛选信用债、利率债、存单等。', placeholder: '信用债' },
      { name: '剩余期限', description: '按剩余期限区间筛选。', placeholder: '1Y-3Y' },
    ],
    defaultOutputFields: ['债券简称', '发行人', '余额', '剩余期限', '托管量'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_BALANCE',
      parameterMapping: { bondCodes: 'A2', valuationDate: 'A3', bondType: 'A4', tenorBucket: 'A5' },
      fieldMapping: {
        债券简称: 'bondShortName',
        发行人: 'issuerName',
        余额: 'outstandingBalance',
        剩余期限: 'remainingTenor',
        托管量: 'custodyAmount',
      },
    },
    previewRows: [
      ['估值日', '2026-05-12', ''],
      ['债券简称', '余额', '剩余期限'],
      ['23示例债01', '18.50亿', '2.4Y'],
    ],
  },
  {
    templateId: 'tpl_lgfv_basic',
    templateName: '城投债基础信息拉取',
    description: '按省份或平台名单批量拉取城投债基础字段，适合区域城投债筛选。',
    inputRanges: ['A2:A7'],
    outputRange: 'B2:H10',
    metricBindings: [],
    refreshRules: '按区域、评级和债券类型生成城投债清单。',
    formatPreset: '城投筛选表',
    type: 'default',
    applicableObject: 'bond_set',
    requiredParams: [
      { name: '省份或城投平台名单', description: '输入省份、地市或平台名单。', placeholder: '江苏省 / 平台名单' },
    ],
    optionalParams: [
      { name: '评级', description: '筛选主体或债项评级。', placeholder: 'AA+及以上' },
      { name: '行政层级', description: '筛选省级、市级、区县级平台。', placeholder: '市级' },
      { name: '债券类型', description: '筛选中票、公司债、企业债等。', placeholder: '中票' },
    ],
    defaultOutputFields: ['债券简称', '发行人', '省份', '地市', '主体评级', '债项评级', '余额'],
    hiddenFetchConfig: {
      interfaceName: 'DM_LGFV_BOND_SCREEN',
      parameterMapping: { regionOrIssuers: 'A2', rating: 'A3', adminLevel: 'A4', bondType: 'A5' },
      fieldMapping: {
        债券简称: 'bondShortName',
        发行人: 'issuerName',
        省份: 'province',
        地市: 'city',
        主体评级: 'issuerRating',
        债项评级: 'bondRating',
        余额: 'outstandingBalance',
      },
    },
    previewRows: [
      ['区域', '江苏省', '评级 AA+'],
      ['债券简称', '地市', '余额'],
      ['24苏城投MTN001', '苏州', '12.00亿'],
    ],
  },
  {
    templateId: 'tpl_bond_valuation_yield',
    templateName: '债券估值与收益率监控',
    description: '跟踪单券估值净价、收益率、久期和利差变化，适合交易前后监控。',
    inputRanges: ['A2:A6'],
    outputRange: 'B2:F8',
    metricBindings: [],
    refreshRules: '按债券代码、估值日和估值源刷新估值指标。',
    formatPreset: '估值监控表',
    type: 'default',
    applicableObject: 'bond',
    requiredParams: [
      { name: '债券代码', description: '需要监控的债券代码。', placeholder: '102300123.IB' },
      { name: '估值日', description: '估值指标对应日期。', placeholder: '2026-05-12' },
    ],
    optionalParams: [
      { name: '估值源', description: '中债、中证或内部估值。', placeholder: '中债' },
      { name: '指标', description: '指定收益率、久期、利差等指标组。', placeholder: '估值+收益率' },
    ],
    defaultOutputFields: ['中债估值净价', '收益率', '久期', '利差', '估值变动'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_VALUATION',
      parameterMapping: { bondCode: 'A2', valuationDate: 'A3', source: 'A4', metrics: 'A5' },
      fieldMapping: {
        中债估值净价: 'cdbCleanPrice',
        收益率: 'yieldToMaturity',
        久期: 'modifiedDuration',
        利差: 'spread',
        估值变动: 'valuationChange',
      },
    },
    previewRows: [
      ['债券代码', '102300123.IB', '中债'],
      ['净价', '收益率', '利差'],
      ['101.23', '2.64%', '58bp'],
    ],
  },
  {
    templateId: 'tpl_rating_query',
    templateName: '主体评级与债项评级查询',
    description: '按主体或债券查询主体评级、债项评级和评级展望，适合信用跟踪。',
    inputRanges: ['A2:A6'],
    outputRange: 'B2:F8',
    metricBindings: [],
    refreshRules: '按主体名称或债券代码刷新最新评级信息。',
    formatPreset: '评级跟踪表',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '主体名称或债券代码', description: '支持发行人名称、统一社会信用代码或债券代码。', placeholder: '示例城投集团 / 102300123.IB' },
    ],
    optionalParams: [
      { name: '评级机构', description: '限定中诚信、联合、东方金诚等。', placeholder: '中诚信' },
      { name: '评级日期', description: '查询某日期前后的评级。', placeholder: '2026-05-12' },
    ],
    defaultOutputFields: ['主体评级', '评级展望', '债项评级', '评级机构', '最新变动日'],
    hiddenFetchConfig: {
      interfaceName: 'DM_CREDIT_RATING',
      parameterMapping: { issuerOrBond: 'A2', agency: 'A3', ratingDate: 'A4' },
      fieldMapping: {
        主体评级: 'issuerRating',
        评级展望: 'ratingOutlook',
        债项评级: 'bondRating',
        评级机构: 'ratingAgency',
        最新变动日: 'latestRatingChangeDate',
      },
    },
    previewRows: [
      ['主体/债券', '示例城投集团', ''],
      ['主体评级', '展望', '变动日'],
      ['AA+', '稳定', '2026-03-18'],
    ],
  },
  {
    templateId: 'tpl_outstanding_bonds',
    templateName: '存续债列表',
    description: '输入发行人后生成存续债清单，适合主体信用分析和持仓排查。',
    inputRanges: ['A2:A6'],
    outputRange: 'B2:H10',
    metricBindings: [],
    refreshRules: '按发行人和筛选条件生成存续债列表。',
    formatPreset: '存续债清单',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '发行人', description: '需要查询存续债的发行主体。', placeholder: '示例城投集团有限公司' },
    ],
    optionalParams: [
      { name: '是否城投', description: '限定城投或非城投主体。', placeholder: '是' },
      { name: '债券类型', description: '筛选中票、公司债、短融、存单等。', placeholder: '中票' },
      { name: '剩余期限', description: '按剩余期限区间筛选。', placeholder: '1Y-5Y' },
    ],
    defaultOutputFields: ['债券代码', '债券简称', '发行日期', '到期日', '余额', '票息', '评级'],
    hiddenFetchConfig: {
      interfaceName: 'DM_ISSUER_OUTSTANDING_BONDS',
      parameterMapping: { issuerName: 'A2', isLgfv: 'A3', bondType: 'A4', tenorBucket: 'A5' },
      fieldMapping: {
        债券代码: 'bondCode',
        债券简称: 'bondShortName',
        发行日期: 'issueDate',
        到期日: 'maturityDate',
        余额: 'outstandingBalance',
        票息: 'couponRate',
        评级: 'bondRating',
      },
    },
    previewRows: [
      ['发行人', '示例城投集团', ''],
      ['债券代码', '债券简称', '余额'],
      ['102300123.IB', '23示例债01', '18.50亿'],
    ],
  },
  {
    templateId: 'tpl_issuer_financial_summary',
    templateName: '主体财务摘要',
    description: '按主体和报告期拉取核心财务摘要，适合信用分析前置数据准备。',
    inputRanges: ['A2:A6'],
    outputRange: 'B2:G8',
    metricBindings: [],
    refreshRules: '按发行人、报告期和指标组刷新财务摘要。',
    formatPreset: '财务摘要表',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '发行人', description: '需要查询财务数据的主体。', placeholder: '示例城投集团有限公司' },
      { name: '报告期', description: '年报、半年报或最新一期。', placeholder: '2025A' },
    ],
    optionalParams: [
      { name: '报表口径', description: '合并或母公司口径。', placeholder: '合并' },
      { name: '指标组', description: '资产负债、利润、现金流等。', placeholder: '信用摘要' },
    ],
    defaultOutputFields: ['总资产', '总负债', '营业收入', '净利润', '资产负债率', '经营现金流'],
    hiddenFetchConfig: {
      interfaceName: 'DM_ISSUER_FINANCIALS',
      parameterMapping: { issuerName: 'A2', reportPeriod: 'A3', statementScope: 'A4', metricSet: 'A5' },
      fieldMapping: {
        总资产: 'totalAssets',
        总负债: 'totalLiabilities',
        营业收入: 'operatingRevenue',
        净利润: 'netProfit',
        资产负债率: 'debtToAssetRatio',
        经营现金流: 'operatingCashFlow',
      },
    },
    previewRows: [
      ['发行人', '示例城投集团', '2025A'],
      ['总资产', '资产负债率', '经营现金流'],
      ['1260亿', '62.4%', '38亿'],
    ],
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
  if (!numericYear && metricCode.startsWith('DM_')) return mockTemplateFieldValue(metricCode, year);
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

function mockTemplateFieldValue(metricCode: string, fieldName: string) {
  const samples: Record<string, string> = {
    债券代码: '102300123.IB',
    债券简称: '23示例债01',
    发行人: '示例城投集团有限公司',
    债券类型: '中期票据',
    起息日: '2023-04-21',
    到期日: '2028-04-21',
    余额: '18.50亿',
    评级: 'AA+',
    剩余期限: '2.94Y',
    托管量: '18.45亿',
    省份: '江苏省',
    地市: '苏州市',
    主体评级: 'AA+',
    债项评级: 'AA+',
    中债估值净价: '101.23',
    收益率: '2.64%',
    久期: '2.71',
    利差: '58bp',
    估值变动: '+0.08',
    评级展望: '稳定',
    评级机构: '中诚信国际',
    最新变动日: '2026-03-18',
    发行日期: '2023-04-19',
    票息: '3.28%',
    总资产: '1260亿',
    总负债: '786亿',
    营业收入: '94亿',
    净利润: '11.6亿',
    资产负债率: '62.4%',
    经营现金流: '38亿',
  };
  if (samples[fieldName]) return samples[fieldName];
  const fieldKey = metricCode.split('.').pop() ?? '';
  return samples[fieldKey] ?? '已取数';
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

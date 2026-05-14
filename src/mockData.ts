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
    templateId: 'tpl_issuer_blank',
    templateName: '空白主体模板',
    description: '主体取数的空白框架，第一列输入主体名称，后续字段由用户选择或让 AI 补充。',
    inputRanges: ['A2:A39'],
    outputRange: 'B1:AB39',
    metricBindings: [],
    refreshRules: '按主体名称和用户选择的指标生成隐藏取数配置。',
    formatPreset: '主体空白表',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '主体名称', description: '需要拉取信息的发行主体或企业名称。', placeholder: '山东高速集团有限公司' },
    ],
    optionalParams: [
      { name: '指标', description: '点击选择主体相关指标。', placeholder: '点击选择指标' },
    ],
    defaultOutputFields: ['点击选择指标'],
    hiddenFetchConfig: {
      interfaceName: 'DM_ISSUER_CUSTOM_FIELDS',
      parameterMapping: { issuerName: 'A2', selectedFields: 'B1:AB1' },
      fieldMapping: { 点击选择指标: 'selectedFields' },
    },
    previewRows: [
      ['主体名称', '点击选择指标', ''],
      ['山东高速集团有限公司', '', ''],
    ],
  },
  {
    templateId: 'tpl_issuer_basic_info',
    templateName: '主体基础信息',
    description: '按主体名称拉取法定代表人、实际控制人、公司性质、注册资本和经营范围等基础资料。',
    inputRanges: ['A2:A39'],
    outputRange: 'A1:AD39',
    metricBindings: [],
    refreshRules: '按主体名称刷新主体基础信息字段。',
    formatPreset: '主体基础信息表',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '主体名称', description: '发行主体或企业名称。', placeholder: '山东高速集团有限公司' },
    ],
    optionalParams: [
      { name: '字段组', description: '基础信息、工商信息或行业信息。', placeholder: '基础信息' },
    ],
    defaultOutputFields: ['法定代表人', '实际控制人', '公司性质', '成立日期', '注册资本', '注册地址', '行业分类'],
    hiddenFetchConfig: {
      interfaceName: 'DM_ISSUER_PROFILE',
      parameterMapping: { issuerName: 'A2' },
      fieldMapping: {
        法定代表人: 'legalRepresentative',
        实际控制人: 'actualController',
        公司性质: 'companyNature',
        成立日期: 'establishedDate',
        注册资本: 'registeredCapital',
        注册地址: 'registeredAddress',
        行业分类: 'industry',
      },
    },
    previewRows: [
      ['主体名称', '法定代表人', '实际控制人'],
      ['山东高速集团', '王其峰', '山东省国资委'],
      ['公司性质', '地方国有企业', '注册资本'],
    ],
  },
  {
    templateId: 'tpl_issuer_top_shareholders',
    templateName: '主体前十大股东',
    description: '输入主体名称，在输出表查看股东名称、股东性质、认缴出资额、持股比例和简介。',
    inputRanges: ['输入表!A2:A12'],
    outputRange: '输出表-主体前十大股东!A1:F10',
    metricBindings: [],
    refreshRules: '按输入表主体名称刷新前十大股东明细。',
    formatPreset: '主体股东穿透表',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '主体名称', description: '需要查询股东结构的主体名称。', placeholder: '山东高速集团有限公司' },
    ],
    optionalParams: [
      { name: '股东数量', description: '默认展示前十大股东。', placeholder: '10' },
    ],
    defaultOutputFields: ['主体名称', '股东名称', '股东性质', '认缴出资额(万元)', '持股比例(%)', '股东简介'],
    hiddenFetchConfig: {
      interfaceName: 'DM_ISSUER_TOP_SHAREHOLDERS',
      parameterMapping: { issuerName: '输入表!A2', limit: '10' },
      fieldMapping: {
        主体名称: 'issuerName',
        股东名称: 'shareholderName',
        股东性质: 'shareholderType',
        '认缴出资额(万元)': 'subscribedCapital',
        '持股比例(%)': 'shareholdingRatio',
        股东简介: 'shareholderProfile',
      },
    },
    previewRows: [
      ['主体名称', '股东名称', '持股比例'],
      ['山东高速集团', '山东省国资委', '70.37%'],
      ['山东高速集团', '山东发展投资', '19.75%'],
    ],
  },
  {
    templateId: 'tpl_credit_bank_peer',
    templateName: '授信指标-银行同业',
    description: '面向银行同业授信，按主体和报告期展示规模增长、资产负债、同业负债、贷款等指标。',
    inputRanges: ['A1:B2'],
    outputRange: 'A2:E40',
    metricBindings: [],
    refreshRules: '按主体名称和报告期横向刷新授信主要指标。',
    formatPreset: '银行同业授信指标表',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '主体名称', description: '银行或金融机构主体。', placeholder: '请输入' },
      { name: '报告期', description: '需要对比的报告期。', placeholder: '2025-12-31' },
    ],
    optionalParams: [
      { name: '指标组', description: '规模增长、资产质量、盈利能力等。', placeholder: '授信主要指标' },
    ],
    defaultOutputFields: ['总资产(亿元)', '负债总额(亿元)', '净资产(亿元)', '同业负债/总负债(%)', '贷款总额(亿元)'],
    hiddenFetchConfig: {
      interfaceName: 'DM_CREDIT_BANK_PEER_METRICS',
      parameterMapping: { issuerName: 'B1', reportPeriods: 'B2:E2' },
      fieldMapping: {
        '总资产(亿元)': 'totalAssets',
        '负债总额(亿元)': 'totalLiabilities',
        '净资产(亿元)': 'netAssets',
        '同业负债/总负债(%)': 'interbankLiabilityRatio',
        '贷款总额(亿元)': 'totalLoans',
      },
    },
    previewRows: [
      ['主体名称', '请输入', ''],
      ['报告期', '2025-12-31', '2024-12-31'],
      ['总资产(亿元)', '', ''],
    ],
  },
  {
    templateId: 'tpl_credit_securities',
    templateName: '授信指标-证券行业',
    description: '面向证券行业授信，按主体和报告期展示资产、负债、利润、同业资产和同业负债。',
    inputRanges: ['A1:B2'],
    outputRange: 'A2:E15',
    metricBindings: [],
    refreshRules: '按证券公司主体和报告期刷新授信指标。',
    formatPreset: '证券行业授信指标表',
    type: 'default',
    applicableObject: 'issuer',
    requiredParams: [
      { name: '主体名称', description: '证券公司主体名称。', placeholder: '请输入' },
      { name: '报告期', description: '需要对比的报告期。', placeholder: '2025-12-31' },
    ],
    optionalParams: [
      { name: '指标组', description: '资产负债、利润、同业指标等。', placeholder: '证券主要指标' },
    ],
    defaultOutputFields: ['资产总额(亿元)', '负债总额(亿元)', '利润总额(亿元)', '净利润(亿元)', '同业资产(亿元)', '同业负债(亿元)'],
    hiddenFetchConfig: {
      interfaceName: 'DM_CREDIT_SECURITIES_METRICS',
      parameterMapping: { issuerName: 'B1', reportPeriods: 'B2:E2' },
      fieldMapping: {
        '资产总额(亿元)': 'totalAssets',
        '负债总额(亿元)': 'totalLiabilities',
        '利润总额(亿元)': 'totalProfit',
        '净利润(亿元)': 'netProfit',
        '同业资产(亿元)': 'interbankAssets',
        '同业负债(亿元)': 'interbankLiabilities',
      },
    },
    previewRows: [
      ['主体名称', '请输入', ''],
      ['报告期', '2025-12-31', '2024-12-31'],
      ['资产总额(亿元)', '', ''],
    ],
  },
  {
    templateId: 'tpl_bond_blank',
    templateName: '空白债券模板',
    description: '债券取数的空白框架，第一列输入债券代码，后续字段由用户选择或让 AI 补充。',
    inputRanges: ['A2:A39'],
    outputRange: 'B1:AB39',
    metricBindings: [],
    refreshRules: '按债券代码和用户选择的指标生成隐藏取数配置。',
    formatPreset: '债券空白表',
    type: 'default',
    applicableObject: 'bond',
    requiredParams: [
      { name: '债券代码', description: '需要拉取信息的债券代码。', placeholder: '2380186.IB' },
    ],
    optionalParams: [
      { name: '指标', description: '点击选择债券相关指标。', placeholder: '点击选择指标' },
    ],
    defaultOutputFields: ['点击选择指标'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_CUSTOM_FIELDS',
      parameterMapping: { bondCode: 'A2', selectedFields: 'B1:AB1' },
      fieldMapping: { 点击选择指标: 'selectedFields' },
    },
    previewRows: [
      ['债券代码', '点击选择指标', ''],
      ['2380186.IB', '', ''],
    ],
  },
  {
    templateId: 'tpl_bond_basic_info',
    templateName: '债券基础信息',
    description: '按债券代码拉取债券简称、债券名称、发行人、主承销商、是否城投和关键日期。',
    inputRanges: ['A2:A39'],
    outputRange: 'A1:AB39',
    metricBindings: [],
    refreshRules: '按债券代码刷新债券基础信息字段。',
    formatPreset: '债券基础信息表',
    type: 'default',
    applicableObject: 'bond',
    requiredParams: [
      { name: '债券代码', description: '银行间或交易所债券代码。', placeholder: '2380186.IB' },
    ],
    optionalParams: [
      { name: '字段组', description: '基础信息、发行信息、上市信息等。', placeholder: '基础信息' },
    ],
    defaultOutputFields: ['债券简称', '债券名称', '发行人名称', '主承销商', '是否城投', '债券类型', '发行期限'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_PROFILE',
      parameterMapping: { bondCode: 'A2' },
      fieldMapping: {
        债券简称: 'bondShortName',
        债券名称: 'bondFullName',
        发行人名称: 'issuerName',
        主承销商: 'leadUnderwriter',
        是否城投: 'isLgfv',
        债券类型: 'bondType',
        发行期限: 'issueTerm',
      },
    },
    previewRows: [
      ['债券代码', '债券简称', '发行人'],
      ['2380186.IB', '23山东高速债02', '山东高速集团'],
      ['是否城投', '是', '企业债券'],
    ],
  },
  {
    templateId: 'tpl_bond_valuation',
    templateName: '债券估值',
    description: '按估值日期和债券代码拉取中债、中证估值收益率、净价、全价、久期和凸性。',
    inputRanges: ['A1:B3'],
    outputRange: 'A2:AC40',
    metricBindings: [],
    refreshRules: '按估值日期和债券代码刷新估值字段。',
    formatPreset: '债券估值表',
    type: 'default',
    applicableObject: 'bond',
    requiredParams: [
      { name: '估值日期', description: '估值指标对应日期。', placeholder: '2026/5/12' },
      { name: '债券代码', description: '需要查询估值的债券代码。', placeholder: '2380186.IB' },
    ],
    optionalParams: [
      { name: '估值源', description: '中债或中证估值字段。', placeholder: '中债+中证' },
    ],
    defaultOutputFields: ['债券简称', '中债估值待偿期', '中债估值收益率', '中债估值净价', '中债估值全价', '中债估值修正久期'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_VALUATION',
      parameterMapping: { valuationDate: 'B1', bondCode: 'A3' },
      fieldMapping: {
        债券简称: 'bondShortName',
        中债估值待偿期: 'cdbRemainingTerm',
        中债估值收益率: 'cdbYield',
        中债估值净价: 'cdbCleanPrice',
        中债估值全价: 'cdbFullPrice',
        中债估值修正久期: 'cdbModifiedDuration',
      },
    },
    previewRows: [
      ['估值日期', '2026/5/12', ''],
      ['债券代码', '债券简称', '收益率'],
      ['2380186.IB', '23山东高速02', '1.4272'],
    ],
  },
  {
    templateId: 'tpl_bond_calculator',
    templateName: '债券计算器',
    description: '支持行权收益率、到期收益率、净价、全价之间的价格收益率换算。',
    inputRanges: ['A1:D3'],
    outputRange: 'A2:AK40',
    metricBindings: [],
    refreshRules: '按结算日、债券代码和价格/收益率输入刷新计算器结果。',
    formatPreset: '债券价格收益率计算器',
    type: 'default',
    applicableObject: 'bond',
    requiredParams: [
      { name: '结算日', description: '交易结算日期。', placeholder: '2026/5/12' },
      { name: '债券代码', description: '需要计算的债券代码。', placeholder: '2380186.IB' },
      { name: '收益率或价格', description: '按当前计算器页签输入。', placeholder: '2.5 / 100.0686' },
    ],
    optionalParams: [
      { name: '券面总额', description: '用于计算结算金额。', placeholder: '万' },
    ],
    defaultOutputFields: ['债券简称', '债券行权收益率', '债券到期收益率', '债券净价', '债券全价', '应计利息', '结算金额'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_CALCULATOR',
      parameterMapping: { settlementDate: 'B1', bondCode: 'A3', inputYieldOrPrice: 'C3' },
      fieldMapping: {
        债券简称: 'bondShortName',
        债券行权收益率: 'exerciseYield',
        债券到期收益率: 'yieldToMaturity',
        债券净价: 'cleanPrice',
        债券全价: 'fullPrice',
        应计利息: 'accruedInterest',
        结算金额: 'settlementAmount',
      },
    },
    previewRows: [
      ['结算日', '2026/5/12', ''],
      ['债券代码', '行权收益率', '净价'],
      ['2380186.IB', '2.5', '100.0686'],
    ],
  },
  {
    templateId: 'tpl_daily_issuance',
    templateName: '每日发行信息',
    description: '按发行起止日拉取债券发行规模、票面利率、发行人、期限、市场和评级信息。',
    inputRanges: ['A1:C1'],
    outputRange: 'A2:AC40',
    metricBindings: [],
    refreshRules: '按发行起始日和结束日刷新每日发行债券清单。',
    formatPreset: '每日发行信息表',
    type: 'default',
    applicableObject: 'bond_set',
    requiredParams: [
      { name: '发行起始日', description: '发行区间开始日期。', placeholder: '2026/5/12' },
      { name: '发行结束日', description: '发行区间结束日期。', placeholder: '2026/5/12' },
    ],
    optionalParams: [
      { name: '市场', description: '银行间、交易所或全部。', placeholder: '全部' },
      { name: '主体评级', description: '按主体评级筛选。', placeholder: 'AAA' },
    ],
    defaultOutputFields: ['债券代码', '债券简称', '债券名称', '计划发行规模(亿)', '实际发行规模(亿)', '票面利率%', '发行人名称'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_DAILY_ISSUANCE',
      parameterMapping: { issueStartDate: 'B1', issueEndDate: 'C1' },
      fieldMapping: {
        债券代码: 'bondCode',
        债券简称: 'bondShortName',
        债券名称: 'bondFullName',
        '计划发行规模(亿)': 'plannedIssueAmount',
        '实际发行规模(亿)': 'actualIssueAmount',
        '票面利率%': 'couponRate',
        发行人名称: 'issuerName',
      },
    },
    previewRows: [
      ['发行起始日', '2026/5/12', '2026/5/12'],
      ['债券简称', '计划规模', '主体评级'],
      ['23山东高速债02', '3.07', 'AAA'],
    ],
  },
  {
    templateId: 'tpl_bond_regulatory_one_table',
    templateName: '债券信息(监管一表通)',
    description: '面向监管报送的一表通债券信息模板，覆盖债券全称、ISIN、期限、发行规模和募集资金用途等字段。',
    inputRanges: ['A2:A39'],
    outputRange: 'A1:AM39',
    metricBindings: [],
    refreshRules: '按债券代码刷新监管一表通字段。',
    formatPreset: '监管一表通',
    type: 'default',
    applicableObject: 'bond',
    requiredParams: [
      { name: '债券代码', description: '需要填报的一表通债券代码。', placeholder: '2380186.IB' },
    ],
    optionalParams: [
      { name: '报送字段组', description: '监管口径字段集合。', placeholder: '一表通' },
    ],
    defaultOutputFields: ['债券全称', '债券简称', 'ISIN 代码', '投资标的交易代码', '期限(年)', '债券期次', '发行规模(亿)'],
    hiddenFetchConfig: {
      interfaceName: 'DM_BOND_REGULATORY_ONE_TABLE',
      parameterMapping: { bondCode: 'A2' },
      fieldMapping: {
        债券全称: 'bondFullName',
        债券简称: 'bondShortName',
        'ISIN 代码': 'isin',
        投资标的交易代码: 'investmentTradingCode',
        '期限(年)': 'termYears',
        债券期次: 'bondTranche',
        '发行规模(亿)': 'issueAmount',
      },
    },
    previewRows: [
      ['债券代码', '债券全称', '债券简称'],
      ['2380186.IB', '2023年第二期山东高速...', '23山东高速债02'],
      ['期限(年)', '发行规模(亿)', '募集资金用途'],
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

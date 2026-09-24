import React, { useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import demoRows from './data/demoRows.json'
import demoAnalysis from './data/demoAnalysis.json'
import './styles.css'

const columns = ['账号名称', '粉丝数', '笔记标题', '正文', '发布时间', '点赞数', '收藏数', '评论数', '内容形式', '笔记链接', '样本类型']
const requiredColumns = ['账号名称', '笔记标题', '发布时间', '点赞数', '收藏数', '评论数']
const fieldAliases = {
  '笔记标题': ['标题', '笔记标题'],
  正文: ['正文', '笔记正文', '内容正文'],
  发布时间: ['发布时间', '日期', '发布日'],
  点赞数: ['点赞', '点赞数'],
  收藏数: ['收藏', '收藏数'],
  评论数: ['评论', '评论数'],
  内容形式: ['形式', '内容形式', '笔记形式'],
  笔记链接: ['链接', '笔记链接', '原文链接'],
  账号名称: ['账号', '账号名称'],
  粉丝数: ['粉丝', '粉丝数'],
  样本类型: ['样本类型'],
}
const analysisColumns = ['内容主题', '目标人群', '用户需求', '内容类型', '内容场景', '标题钩子', '内容目的']
const analysisTableColumns = ['笔记标题', ...analysisColumns]
const ANALYSIS_STORAGE_KEY = 'insurance_content_analysis_v1'
const DATASET_STORAGE_KEY = 'insurance_content_rows_v1'
// 公开生产构建（Vercel）禁用 AI 重新分析，避免暴露 /api/analyze 后端调用
const IS_PUBLIC_DEMO = import.meta.env.PROD
const COMBINATION_GROUPS = [
  { key: 'audience-need', label: '目标人群 × 用户需求', left: '目标人群', right: '用户需求', template: (left, right) => `面向【${left}】人群，围绕【${right}】需求的内容，在当前样本中表现高于整体基准。` },
  { key: 'need-hook', label: '用户需求 × 标题钩子', left: '用户需求', right: '标题钩子', template: (left, right) => `围绕【${left}】需求，采用【${right}】标题钩子的内容，在当前样本中表现高于整体基准。` },
  { key: 'audience-format', label: '目标人群 × 内容类型', left: '目标人群', right: '内容类型', template: (left, right) => `面向【${left}】人群，以【${right}】形式呈现的内容，在当前样本中表现高于整体基准。` },
]
const TAG_TAXONOMY = {
  内容主题: ['医疗险', '重疾险', '意外险', '寿险', '储蓄险', '社保医保', '家庭保险配置', '理赔', '保险基础知识', '资产规划', '其他'],
  目标人群: ['年轻职场人', '女性', '宝妈宝爸', '儿童', '父母/中老年', '中产家庭', '家庭人群', '保险小白', '通用人群'],
  用户需求: ['不会买/不会选', '避坑防骗', '省钱控预算', '完善保障', '理赔解决', '疾病投保', '家庭配置', '养老储蓄', '资产规划', '保险知识学习'],
  内容类型: ['科普知识', '购买攻略', '避坑指南', '产品测评', '产品对比', '案例故事', '经验分享', '观点讨论', '热点借势'],
  内容场景: ['首次投保', '个人配置', '夫妻配置', '儿童投保', '父母投保', '家庭配置', '疾病投保', '理赔', '养老退休', '储蓄理财', '其他'],
  标题钩子: ['数字量化', '提问', '避坑警示', '反常识', '利益点', '焦虑痛点', '身份代入', '结果展示', '经验背书', '热点'],
  内容目的: ['拉新流量', '用户教育', '建立信任', '需求激发', '产品种草', '转化获客'],
}

const seedRows = [
  { 账号名称: '保贝说险', 粉丝数: 128600, 笔记标题: '给爸妈买保险，先看懂这 3 个坑', 正文: '用真实案例拆解父母投保时最容易忽略的保障细节。', 发布时间: '2026-09-18', 点赞数: 8240, 收藏数: 5630, 评论数: 486, 内容形式: '图文', 笔记链接: 'https://www.xiaohongshu.com/' },
  { 账号名称: '小周聊保障', 粉丝数: 76300, 笔记标题: '百万医疗险到底怎么选？一张表讲清', 正文: '把免赔额、续保和外购药拆成一眼能看懂的对比。', 发布时间: '2026-09-16', 点赞数: 5310, 收藏数: 4820, 评论数: 218, 内容形式: '图文', 笔记链接: 'https://www.xiaohongshu.com/' },
  { 账号名称: '懂保汇', 粉丝数: 215400, 笔记标题: '重疾险不是越贵越好，预算这样分配', 正文: '针对不同人生阶段，给出可执行的重疾险预算框架。', 发布时间: '2026-09-14', 点赞数: 9670, 收藏数: 7210, 评论数: 693, 内容形式: '视频', 笔记链接: 'https://www.xiaohongshu.com/' },
  { 账号名称: '保单研究所', 粉丝数: 94200, 笔记标题: '买了保险却不能赔？这 5 种情况要避开', 正文: '围绕理赔争议场景，梳理投保前需要确认的关键信息。', 发布时间: '2026-09-12', 点赞数: 6840, 收藏数: 5100, 评论数: 351, 内容形式: '长图', 笔记链接: 'https://www.xiaohongshu.com/' },
  { 账号名称: '一颗糖的保险课', 粉丝数: 58700, 笔记标题: '月薪 8000，保险配置顺序怎么排？', 正文: '从医疗、意外到重疾，用收入视角讲明白保障优先级。', 发布时间: '2026-09-10', 点赞数: 4120, 收藏数: 3980, 评论数: 175, 内容形式: '图文', 笔记链接: 'https://www.xiaohongshu.com/' },
  { 账号名称: '保贝说险', 粉丝数: 128600, 笔记标题: '体检报告有异常，还能买保险吗？', 正文: '常见体检异常的核保思路与准备材料清单。', 发布时间: '2026-09-08', 点赞数: 7380, 收藏数: 4650, 评论数: 402, 内容形式: '视频', 笔记链接: 'https://www.xiaohongshu.com/' },
  { 账号名称: '懂保汇', 粉丝数: 215400, 笔记标题: '孩子的第一份保险，别从返还型开始', 正文: '儿童保障配置的基础逻辑，以及常见产品话术辨析。', 发布时间: '2026-09-05', 点赞数: 10600, 收藏数: 8040, 评论数: 759, 内容形式: '图文', 笔记链接: 'https://www.xiaohongshu.com/' },
]

const accountColors = ['coral', 'blue', 'green', 'gold', 'violet']

function formatNumber(value) {
  return new Intl.NumberFormat('zh-CN').format(Number(value) || 0)
}

function isMissing(value) {
  return value === undefined || value === null || String(value).trim() === ''
}

function displayText(value) {
  return isMissing(value) ? '—' : value
}

function normalizeSampleType(value) {
  const normalized = String(value ?? '').trim()
  if (normalized === '高表现' || normalized === '高表现笔记') return '高表现'
  if (normalized === '普通' || normalized === '普通笔记') return '普通'
  return isMissing(normalized) ? '—' : normalized
}

function formatOptionalNumber(value) {
  return isMissing(value) || value === '—' ? '—' : formatNumber(value)
}

function withAnalysisFields(row) {
  return analysisColumns.reduce((result, column) => {
    const value = row[column]
    const labels = Array.isArray(value) ? value : (isMissing(value) ? [] : [value])
    return { ...result, [column]: labels.length ? labels : ['待分析'] }
  }, { ...row })
}

function getDatasetKey(rows) {
  return JSON.stringify(rows.map(getRowKey))
}

function getRowKey(row) {
  return row['笔记链接'] && row['笔记链接'] !== '—'
    ? `link:${row['笔记链接']}`
    : `account-title:${row['账号名称'] || ''}::${row['笔记标题'] || ''}`
}

function loadInitialRows() {
  try {
    const savedRows = JSON.parse(localStorage.getItem(DATASET_STORAGE_KEY) || 'null')
    if (Array.isArray(savedRows) && savedRows.length) return loadSavedResults(savedRows.map((row) => withAnalysisFields({ ...row, '样本类型': normalizeSampleType(row['样本类型']) })))
    return loadSavedResults(demoRows.map((row) => withAnalysisFields({ ...row, '样本类型': normalizeSampleType(row['样本类型']) })), demoAnalysis)
  } catch {
    return loadSavedResults(demoRows.map((row) => withAnalysisFields({ ...row, '样本类型': normalizeSampleType(row['样本类型']) })), demoAnalysis)
  }
}

function hasAnalysisResult(row) {
  return analysisColumns.every((column) => Array.isArray(row[column]) && row[column].length > 0 && !row[column].includes('待分析'))
}

function loadSavedResults(rows, fallbackResults = null) {
  try {
    const saved = JSON.parse(localStorage.getItem(ANALYSIS_STORAGE_KEY) || '{}')
    const results = saved[getDatasetKey(rows)] || fallbackResults || {}
    return rows.map((row, index) => {
      const result = results[getRowKey(row)] || saved[getRowKey(row)]
      return result ? { ...row, ...result } : row
    })
  } catch {
    if (fallbackResults) return rows.map((row) => fallbackResults[getRowKey(row)] ? { ...row, ...fallbackResults[getRowKey(row)] } : row)
    return rows
  }
}

function saveAnalysisResult(rows, index, labels) {
  try {
    const saved = JSON.parse(localStorage.getItem(ANALYSIS_STORAGE_KEY) || '{}')
    const id = getRowKey(rows[index])
    saved[id] = { id, ...Object.fromEntries(analysisColumns.map((column) => [column, labels[column]])) }
    localStorage.setItem(ANALYSIS_STORAGE_KEY, JSON.stringify(saved))
  } catch {
    // Storage failures should not interrupt analysis.
  }
}

function saveRawDataset(rows) {
  try {
    localStorage.setItem(DATASET_STORAGE_KEY, JSON.stringify(rows))
  } catch {
    // Storage failures should not interrupt import or analysis.
  }
}

function calculateOpportunityStats(rows, dimension) {
  const validRows = rows.filter((row) => row['样本类型'] === '高表现' || row['样本类型'] === '普通')
  const baselineHighRate = validRows.length ? validRows.filter((row) => row['样本类型'] === '高表现').length / validRows.length : 0
  const tagMap = new Map()
  validRows.forEach((row) => {
    const labels = Array.isArray(row[dimension]) ? new Set(row[dimension].filter((label) => label && label !== '待分析')) : new Set()
    labels.forEach((label) => {
      const current = tagMap.get(label) || { label, rows: [], sampleCount: 0, highCount: 0 }
      current.rows.push(row)
      current.sampleCount += 1
      if (row['样本类型'] === '高表现') current.highCount += 1
      tagMap.set(label, current)
    })
  })
  const tags = [...tagMap.values()].map((tag) => {
    const highRate = tag.sampleCount ? tag.highCount / tag.sampleCount : 0
    const supplyRate = validRows.length ? tag.sampleCount / validRows.length : 0
    const performanceLift = highRate - baselineHighRate
    return { ...tag, highRate, supplyRate, performanceLift, opportunityScore: performanceLift * (1 - supplyRate) }
  })
  const sortedSupplyRates = tags.map((tag) => tag.supplyRate).sort((left, right) => left - right)
  const middle = Math.floor(sortedSupplyRates.length / 2)
  const medianSupply = sortedSupplyRates.length ? (sortedSupplyRates.length % 2 ? sortedSupplyRates[middle] : (sortedSupplyRates[middle - 1] + sortedSupplyRates[middle]) / 2) : 0
  const candidates = tags.filter((tag) => tag.sampleCount >= 3 && tag.highRate > baselineHighRate && tag.supplyRate < medianSupply).sort((left, right) => right.opportunityScore - left.opportunityScore)
  return { validRows, baselineHighRate, medianSupply, tags, candidates }
}

function clearProjectData() {
  if (!window.confirm('确认清空当前项目数据和 AI 标签吗？')) return
  localStorage.removeItem(DATASET_STORAGE_KEY)
  localStorage.removeItem(ANALYSIS_STORAGE_KEY)
  window.location.reload()
}

function getSheetInfo(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName]
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
  const headerRowIndex = matrix.findIndex((row) => row.some((value) => !isMissing(value)))
  const headers = headerRowIndex >= 0 ? matrix[headerRowIndex] : []
  return { sheetName, sheet, matrix, headerRowIndex, headers }
}

function hasCoreFields(headers) {
  const headerSet = new Set(headers.map((header) => String(header).trim()))
  return requiredColumns.every((column) => fieldAliases[column].some((alias) => headerSet.has(alias)))
}

function parseFile(file, onSuccess, onError) {
  const reader = new FileReader()
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true })
      const sheetNames = workbook.SheetNames
      console.log('[Excel Debug] Excel 中所有 Sheet 名称:', sheetNames)
      const isCsv = file.name.toLowerCase().endsWith('.csv')
      const sheetInfos = sheetNames.map((sheetName) => getSheetInfo(workbook, sheetName))
      const exactSheet = sheetInfos.find(({ sheetName }) => sheetName === '笔记原始数据')
      const sheetInfo = isCsv
        ? sheetInfos[0]
        : exactSheet || sheetInfos.find(({ headers }) => hasCoreFields(headers))
      if (!sheetInfo) throw new Error('Excel 中找不到符合条件的笔记数据 Sheet，请确认存在“笔记原始数据”或包含核心字段的工作表。')

      const { sheet, sheetName, headerRowIndex, headers } = sheetInfo
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', range: headerRowIndex })
      console.log('[Excel Debug] 当前实际读取的 Sheet 名称:', sheetName)
      console.log('[Excel Debug] 实际读取到的表头数组:', headers)
      console.log('[Excel Debug] 前3行解析后的原始数据:', rows.slice(0, 3))
      const headerLookup = new Map(headers.map((header) => [String(header).trim(), header]))
      const fieldMap = Object.fromEntries(columns.map((column) => {
        const sourceHeader = fieldAliases[column].find((alias) => headerLookup.has(alias))
        return [column, sourceHeader ? headerLookup.get(sourceHeader) : null]
      }))
      const missingColumns = requiredColumns.filter((column) => !fieldMap[column])
      if (missingColumns.length) {
        const actualHeaders = headers.length ? headers.map((header) => String(header).trim()).join('、') : '（未读取到表头）'
        throw new Error(`缺少列名：${missingColumns.join('、')}。实际读取到的列名：${actualHeaders}`)
      }
      const invalidSampleTypeCount = rows.filter((row) => !['高表现', '普通'].includes(normalizeSampleType(fieldMap['样本类型'] ? row[fieldMap['样本类型']] : ''))).length
      const validRows = rows
        .filter((row) => Object.values(row).some((value) => !isMissing(value)))
        .map((row) => withAnalysisFields({ ...Object.fromEntries(columns.map((column) => [column, fieldMap[column] ? row[fieldMap[column]] : '—'])), ...row, '样本类型': normalizeSampleType(fieldMap['样本类型'] ? row[fieldMap['样本类型']] : '') }))
      if (!validRows.length) throw new Error('文件中没有找到有效内容，请检查列名是否正确。')
      onSuccess(validRows, sheetName, invalidSampleTypeCount ? `样本类型校验提示：${invalidSampleTypeCount} 条为空或不是“高表现/普通”` : '')
    } catch (error) {
      onError(error.message || '文件解析失败，请重试。')
    }
  }
  reader.onerror = () => onError('文件读取失败，请重试。')
  reader.readAsArrayBuffer(file)
}

function App() {
  const [rows, setRows] = useState(loadInitialRows)
  const [selectedAccounts, setSelectedAccounts] = useState([])
  const [fileName, setFileName] = useState('')
  const [message, setMessage] = useState('')
  const [analysisMessage, setAnalysisMessage] = useState('')
  const [analysisRunning, setAnalysisRunning] = useState(false)
  const [failedIndexes, setFailedIndexes] = useState([])
  const [insightDimension, setInsightDimension] = useState(analysisColumns[0])
  const [isDemoDataset] = useState(() => !localStorage.getItem(DATASET_STORAGE_KEY))
  const fileInputRef = useRef(null)

  const accounts = useMemo(() => [...new Set(rows.map((row) => row['账号名称']).filter(Boolean))], [rows])
  const filteredRows = useMemo(
    () => selectedAccounts.length ? rows.filter((row) => selectedAccounts.includes(row['账号名称'])) : rows,
    [rows, selectedAccounts],
  )
  const stats = useMemo(() => ({
    accounts: accounts.length,
    notes: rows.length,
    likes: rows.reduce((sum, row) => sum + (Number(row['点赞数']) || 0), 0),
    saves: rows.reduce((sum, row) => sum + (Number(row['收藏数']) || 0), 0),
    totalSample: rows.length,
    highPerformance: rows.filter((row) => row['样本类型'] === '高表现').length,
    ordinary: rows.filter((row) => row['样本类型'] === '普通').length,
  }), [accounts.length, rows])
  const analysisCoverage = useMemo(() => rows.filter(hasAnalysisResult).length, [rows])
  const insightStats = useMemo(() => {
    const validRows = rows.filter((row) => row['样本类型'] === '高表现' || row['样本类型'] === '普通')
    const overallRate = validRows.length ? validRows.filter((row) => row['样本类型'] === '高表现').length / validRows.length : 0
    const tagMap = new Map()
    validRows.forEach((row) => {
      const labels = Array.isArray(row[insightDimension]) ? new Set(row[insightDimension].filter((label) => label && label !== '待分析')) : new Set()
      labels.forEach((label) => {
        const current = tagMap.get(label) || { label, sampleCount: 0, highCount: 0, ordinaryCount: 0 }
        current.sampleCount += 1
        if (row['样本类型'] === '高表现') current.highCount += 1
        else current.ordinaryCount += 1
        tagMap.set(label, current)
      })
    })
    const tags = [...tagMap.values()].map((tag) => ({
      ...tag,
      rate: tag.sampleCount ? tag.highCount / tag.sampleCount : 0,
      relativeRate: tag.sampleCount ? tag.highCount / tag.sampleCount - overallRate : 0,
    })).sort((left, right) => right.rate - left.rate || right.sampleCount - left.sampleCount)
    return { validRows, overallRate, tags }
  }, [rows, insightDimension])
  const opportunityStats = useMemo(() => {
    return calculateOpportunityStats(rows, insightDimension)
  }, [rows, insightDimension])
  const crossDimensionStats = useMemo(() => analysisColumns.map((dimension) => ({ dimension, ...calculateOpportunityStats(rows, dimension) })), [rows])
  const combinationStats = useMemo(() => COMBINATION_GROUPS.map((group) => ({ ...group, ...calculateCombinationStats(rows, group) })), [rows])

  const handleFile = (file) => {
    if (!file) return
    setMessage('')
    parseFile(file, (data, sheetName, validationMessage) => {
      saveRawDataset(data)
      setRows(loadSavedResults(data))
      setSelectedAccounts([])
      setFailedIndexes([])
      setFileName(file.name)
      setMessage(`已导入：${file.name}｜工作表：${sheetName}｜${data.length} 条内容${validationMessage ? `｜${validationMessage}` : ''}`)
    }, setMessage)
  }

  const toggleAccount = (account) => {
    setSelectedAccounts((current) => current.includes(account) ? current.filter((item) => item !== account) : [...current, account])
  }

  const handleAnalyze = async (retryOnly = false) => {
    if (IS_PUBLIC_DEMO || analysisRunning) return
    const targetIndexes = rows.map((row, index) => ({ row, index })).filter(({ row, index }) => {
      if (retryOnly && !failedIndexes.includes(index)) return false
      return !hasAnalysisResult(row)
    })
    if (!targetIndexes.length) {
      setAnalysisMessage(`${rows.length}篇均已完成AI分析，无需重复调用`)
      return
    }
    setAnalysisRunning(true)
    let batchSuccessCount = 0
    const nextFailedIndexes = []
    const analyzedRows = [...rows]
    const initialCompletedCount = rows.filter(hasAnalysisResult).length
    for (const [position, { index, row }] of targetIndexes.entries()) {
      const completedBeforeRequest = initialCompletedCount + batchSuccessCount
      setAnalysisMessage(`AI分析进度：${completedBeforeRequest + 1} / ${rows.length}\n已完成：${completedBeforeRequest}\n待分析：${rows.length - completedBeforeRequest}`)
      try {
        const response = await fetch('/api/analyze', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            title: row['笔记标题'],
            content: row['正文'],
            TAG_TAXONOMY,
          }),
        })
        const responseBody = await response.text()
        if (!response.ok) {
          const errorMessage = `HTTP ${response.status} ${responseBody}`
          nextFailedIndexes.push(index)
          console.error('[AI API Test] error message:', errorMessage.slice(0, 300))
          continue
        }
        const result = JSON.parse(responseBody)
        if (result.success && result.labels) {
          analyzedRows[index] = { ...analyzedRows[index], ...result.labels }
          batchSuccessCount += 1
          saveAnalysisResult(analyzedRows, index, result.labels)
          setRows([...analyzedRows])
        }
        else nextFailedIndexes.push(index)
      } catch (error) {
        nextFailedIndexes.push(index)
        console.error('[AI API Test] error message:', String(error).slice(0, 300))
      }
    }
    setFailedIndexes(nextFailedIndexes)
    setAnalysisRunning(false)
    const completedCount = analyzedRows.filter(hasAnalysisResult).length
    setAnalysisMessage(`AI分析完成\n成功：${completedCount} / ${rows.length}\n待分析：${rows.length - completedCount}\n本轮新增成功：${batchSuccessCount}\n本轮失败：${nextFailedIndexes.length}`)
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <div className="brand-mark">险</div>
          <div>
            <p className="eyebrow">CONTENT INTELLIGENCE / 01</p>
            <span className="brand-name">保险内容增长分析平台</span>
          </div>
        </div>
        <div className="status-indicator"><span /> 数据仅在本地处理</div>
      </header>

      <section className="hero">
        <div>
          <p className="section-kicker">竞品内容数据库</p>
          <h1>看见高表现内容，<em>找到下一个机会。</em></h1>
          <p className="hero-copy">基于竞品内容数据识别高表现内容与潜在选题机会</p>
        </div>
        <div className="hero-index">01 <span>/</span> CONTENT DB</div>
      </section>

      {isDemoDataset && <div className="status-indicator demo-status-indicator"><span /> 当前展示：保险赛道内容研究 Demo · {rows.length} 篇样本</div>}

      <section className="stats-grid" aria-label="基础统计">
        <StatCard label="竞品账号数" value={stats.accounts} suffix="个" />
        <StatCard label="笔记总数" value={stats.notes} suffix="篇" />
        <StatCard label="总点赞数" value={stats.likes} />
        <StatCard label="总收藏数" value={stats.saves} />
      </section>
      <section className="sample-stats" aria-label="样本类型统计">
        <StatCard label="总样本" value={stats.totalSample} suffix="篇" />
        <StatCard label="高表现" value={stats.highPerformance} suffix="篇" />
        <StatCard label="普通" value={stats.ordinary} suffix="篇" />
      </section>

      <InsightsSection insightDimension={insightDimension} setInsightDimension={setInsightDimension} insightStats={insightStats} />
      <OpportunityMatrixSection insightDimension={insightDimension} setInsightDimension={setInsightDimension} opportunityStats={opportunityStats} />
      <CrossDimensionOpportunities stats={crossDimensionStats} />
      <CombinationOpportunities stats={combinationStats} />

      <section className="upload-panel">
        <div className="upload-icon">↑</div>
        <div className="upload-copy">
          <strong>导入竞品内容数据</strong>
          <span>支持 .xlsx / .xls / .csv，数据将在浏览器本地解析</span>
          {fileName && <small>{fileName} · {message}</small>}
          {message && !fileName && <small className="error-message">{message}</small>}
        </div>
        <button className="upload-button" type="button" onClick={() => fileInputRef.current?.click()}>选择文件</button>
        <button className="clear-project-button" type="button" onClick={clearProjectData}>清空项目数据</button>
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(event) => handleFile(event.target.files?.[0])} />
      </section>

      <section className="database-section">
        <div className="section-heading">
          <div>
            <p className="section-kicker">DATABASE / NOTES</p>
            <h2>内容数据</h2>
          </div>
          <span className="result-count">显示 {filteredRows.length} 条记录</span>
        </div>
        <div className="filter-row">
          <span className="filter-label">筛选竞品账号</span>
          <div className="account-filters">
            {accounts.slice(0, 5).map((account, index) => (
              <button key={account} type="button" className={`account-chip ${selectedAccounts.includes(account) ? 'is-selected' : ''}`} onClick={() => toggleAccount(account)}>
                <span className={`chip-dot ${accountColors[index % accountColors.length]}`} />{account}
              </button>
            ))}
            {selectedAccounts.length > 0 && <button type="button" className="clear-filter" onClick={() => setSelectedAccounts([])}>清除筛选</button>}
          </div>
        </div>
        <div className="table-frame">
          <div className="table-scroll">
            <table>
              <thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
              <tbody>
                {filteredRows.map((row, index) => <TableRow key={`${row['笔记标题']}-${index}`} row={row} />)}
              </tbody>
            </table>
            {!filteredRows.length && <div className="empty-state">当前筛选条件下暂无数据</div>}
          </div>
        </div>
        <p className="table-note">共 {rows.length} 条笔记 · 点击笔记链接可在新窗口查看原文</p>
      </section>

      <section className="analysis-section">
        <div className="section-heading analysis-heading">
          <div>
            <p className="section-kicker">AI CONTENT ANALYSIS</p>
            <h2>AI内容分析</h2>
          </div>
          <div className="analysis-actions">
            {analysisMessage && <span className="analysis-message">{analysisMessage.split('\n').map((line) => <span key={line}>{line}<br /></span>)}</span>}
            {IS_PUBLIC_DEMO ? (
              <span className="analysis-message">在线 Demo 暂不开放 AI 重新分析，当前展示为已完成分析的研究样本。</span>
            ) : (
              <>
                <button className="analysis-button" type="button" disabled={analysisRunning} onClick={() => handleAnalyze(false)}>{analysisRunning ? 'AI分析中...' : '开始AI分析'}</button>
                {!analysisRunning && failedIndexes.length > 0 && <button className="retry-button" type="button" onClick={() => handleAnalyze(true)}>仅重试失败项（{failedIndexes.length}）</button>}
              </>
            )}
          </div>
        </div>
        <div className="table-frame analysis-table-frame">
          <div className="table-scroll">
            <table className="analysis-table">
              <thead><tr>{analysisTableColumns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
              <tbody>
                {filteredRows.map((row, index) => <AnalysisTableRow key={`${row['笔记标题']}-${index}`} row={row} />)}
              </tbody>
            </table>
            {!filteredRows.length && <div className="empty-state">当前筛选条件下暂无数据</div>}
          </div>
        </div>
        <p className="table-note">共 {filteredRows.length} 条待分析内容 · AI标签覆盖：{analysisCoverage} / {rows.length}</p>
      </section>
    </main>
  )
}

function StatCard({ label, value, suffix = '' }) {
  const displayValue = typeof value === 'string' ? value : formatNumber(value)
  return <article className="stat-card"><span>{label}</span><strong>{displayValue}<small>{suffix}</small></strong><i /></article>
}

function TableRow({ row }) {
  return <tr>
    <td><span className="account-name">{displayText(row['账号名称'])}</span></td>
    <td>{formatOptionalNumber(row['粉丝数'])}</td>
    <td className="title-cell">{displayText(row['笔记标题'])}</td>
    <td className="body-cell">{displayText(row['正文'])}</td>
    <td>{row['发布时间'] instanceof Date ? row['发布时间'].toLocaleDateString('zh-CN') : displayText(row['发布时间'])}</td>
    <td className="metric-cell">{formatNumber(row['点赞数'])}</td>
    <td className="metric-cell">{formatNumber(row['收藏数'])}</td>
    <td className="metric-cell">{formatNumber(row['评论数'])}</td>
    <td><span className="format-tag">{displayText(row['内容形式'])}</span></td>
    <td>{isMissing(row['笔记链接']) || row['笔记链接'] === '—' ? '—' : <a className="note-link" href={row['笔记链接']} target="_blank" rel="noreferrer">查看原文 ↗</a>}</td>
    <td><span className="format-tag">{displayText(row['样本类型'])}</span></td>
  </tr>
}

function AnalysisTableRow({ row }) {
  return <tr>
    <td className="title-cell">{displayText(row['笔记标题'])}</td>
    {analysisColumns.map((column) => <td key={column}><div className="analysis-tags">{(Array.isArray(row[column]) ? row[column] : [row[column]]).map((label, index) => <span className="analysis-tag" key={`${column}-${label}-${index}`}>{displayText(label)}</span>)}</div></td>)}
  </tr>
}

function InsightsSection({ insightDimension, setInsightDimension, insightStats }) {
  const { validRows, overallRate, tags } = insightStats
  const highCount = validRows.filter((row) => row['样本类型'] === '高表现').length
  return <section className="insights-section">
    <div className="section-heading insights-heading">
      <div><p className="section-kicker">HIGH-PERFORMANCE INSIGHTS</p><h2>高表现内容特征</h2></div>
    </div>
    <div className="insight-summary">
      <StatCard label="总样本" value={validRows.length} suffix="篇" />
      <StatCard label="高表现样本" value={highCount} suffix="篇" />
      <StatCard label="整体高表现率" value={`${(overallRate * 100).toFixed(1)}%`} />
    </div>
    <div className="insight-tabs" role="tablist">
      {analysisColumns.map((dimension) => <button key={dimension} type="button" className={`insight-tab ${insightDimension === dimension ? 'is-active' : ''}`} onClick={() => setInsightDimension(dimension)}>{dimension}</button>)}
    </div>
    <div className="insight-layout">
      <div className="insight-chart" aria-label={`${insightDimension}高表现率条形图`}>
        <div className="chart-axis"><span>标签</span><span>高表现率 0% — 100%</span></div>
        <div className="chart-body">
          <div className="benchmark-line" style={{ left: `${overallRate * 100}%` }}><span>{(overallRate * 100).toFixed(1)}%</span></div>
          {tags.map((tag) => <div className="chart-row" key={tag.label}>
            <span className="chart-label">{tag.label}{tag.sampleCount < 3 && <small>样本较少</small>}</span>
            <div className="chart-track"><div className={`chart-bar ${tag.sampleCount < 3 ? 'is-small-sample' : ''}`} style={{ width: `${tag.rate * 100}%` }} /></div>
            <strong>{(tag.rate * 100).toFixed(1)}%</strong>
          </div>)}
          {!tags.length && <div className="empty-state">当前没有可用于统计的 AI 标签和样本类型数据</div>}
        </div>
      </div>
      <div className="insight-detail table-frame">
        <table className="insight-table"><thead><tr><th>标签</th><th>样本数</th><th>高表现</th><th>普通</th><th>高表现率</th><th>相对基准</th></tr></thead><tbody>
          {tags.map((tag) => <tr key={tag.label}><td>{tag.label}{tag.sampleCount < 3 && <span className="small-sample-tag">样本较少</span>}</td><td>{tag.sampleCount}</td><td>{tag.highCount}</td><td>{tag.ordinaryCount}</td><td>{(tag.rate * 100).toFixed(1)}%</td><td className={tag.relativeRate >= 0 ? 'positive-rate' : 'negative-rate'}>{tag.relativeRate >= 0 ? '+' : ''}{(tag.relativeRate * 100).toFixed(1)}pct</td></tr>)}
        </tbody></table>
      </div>
    </div>
    <p className="insight-note">高表现率基于当前 {validRows.length} 篇研究样本计算；样本量较小的标签仅供观察，不代表稳定规律。</p>
  </section>
}

function OpportunityMatrixSection({ insightDimension, setInsightDimension, opportunityStats }) {
  const { validRows, baselineHighRate, medianSupply, tags, candidates } = opportunityStats
  const maxSupply = Math.max(...tags.map((tag) => tag.supplyRate), medianSupply, 0.01)
  const hasCandidates = candidates.length > 0
  return <section className="opportunity-section">
    <div className="section-heading insights-heading">
      <div><p className="section-kicker">CONTENT OPPORTUNITY MATRIX</p><h2>内容机会矩阵</h2><p className="opportunity-subtitle">从「内容供给」与「高表现率」两个维度识别候选内容方向</p></div>
    </div>
    <div className="insight-tabs" role="tablist">
      {analysisColumns.map((dimension) => <button key={dimension} type="button" className={`insight-tab ${insightDimension === dimension ? 'is-active' : ''}`} onClick={() => setInsightDimension(dimension)}>{dimension}</button>)}
    </div>
    <div className="opportunity-matrix" aria-label="内容机会矩阵散点图">
      <svg className="matrix-svg" viewBox="0 0 1000 520" role="img" aria-label="内容机会矩阵">
        <rect x="70" y="30" width="890" height="430" fill="#fbfcfd" />
        <rect x="70" y="30" width={890 * (medianSupply / maxSupply)} height={430 * baselineHighRate} fill="#fff8f4" />
        <rect x={70 + 890 * (medianSupply / maxSupply)} y="30" width={890 * (1 - medianSupply / maxSupply)} height={430 * baselineHighRate} fill="#f5f9fd" />
        <rect x="70" y={30 + 430 * baselineHighRate} width={890 * (medianSupply / maxSupply)} height={430 * (1 - baselineHighRate)} fill="#fafafa" />
        <rect x={70 + 890 * (medianSupply / maxSupply)} y={30 + 430 * baselineHighRate} width={890 * (1 - medianSupply / maxSupply)} height={430 * (1 - baselineHighRate)} fill="#f7f8fa" />
        <line x1="70" y1="460" x2="960" y2="460" className="matrix-axis-line" /><line x1="70" y1="30" x2="70" y2="460" className="matrix-axis-line" />
        <line x1="70" y1={460 - 430 * baselineHighRate} x2="960" y2={460 - 430 * baselineHighRate} className="matrix-reference-line" /><line x1={70 + 890 * (medianSupply / maxSupply)} y1="30" x2={70 + 890 * (medianSupply / maxSupply)} y2="460" className="matrix-reference-line" />
        <text x="82" y="52" className="matrix-quadrant-label">潜在内容机会</text><text x="948" y="52" textAnchor="end" className="matrix-quadrant-label">已验证热门方向</text><text x="82" y="448" className="matrix-quadrant-label">暂不优先</text><text x="948" y="448" textAnchor="end" className="matrix-quadrant-label">竞争充分 / 表现一般</text>
        <text x="82" y={450 - 430 * baselineHighRate} className="matrix-reference-label">整体高表现率 {(baselineHighRate * 100).toFixed(1)}%</text><text x={75 + 890 * (medianSupply / maxSupply)} y="43" className="matrix-reference-label">供给率中位数 {(medianSupply * 100).toFixed(1)}%</text>
        {[0, 20, 40, 60, 80, 100].map((value) => <text key={`y-${value}`} x="60" y={465 - value * 4.3} textAnchor="end" className="matrix-tick">{value}%</text>)}
        {[0, .25, .5, .75, 1].map((value) => <text key={`x-${value}`} x={70 + 890 * value} y="480" textAnchor="middle" className="matrix-tick">{(value * maxSupply * 100).toFixed(1)}%</text>)}
        <text x="515" y="510" textAnchor="middle" className="matrix-axis-title">内容供给率（低 → 高）</text><text x="16" y="250" textAnchor="middle" className="matrix-axis-title" transform="rotate(-90 16 250)">高表现率</text>
        {tags.map((tag, index) => { const x = 70 + 890 * (tag.supplyRate / maxSupply); const y = 460 - 430 * tag.highRate; const offset = (index % 3 - 1) * 12; return <g key={tag.label} className={tag.sampleCount < 3 ? 'matrix-svg-point is-small-sample' : 'matrix-svg-point'} transform={`translate(${x},${y})`}><title>{`${tag.label}｜样本数 ${tag.sampleCount}｜高表现 ${tag.highCount}｜普通 ${tag.ordinaryCount}｜高表现率 ${(tag.highRate * 100).toFixed(1)}%｜内容供给率 ${(tag.supplyRate * 100).toFixed(1)}%｜相对基准 ${(tag.performanceLift * 100).toFixed(1)}pct${tag.sampleCount < 3 ? '｜样本较少' : ''}`}</title><circle r="7" /><text x="12" y={offset} className="matrix-point-label">{tag.label}{tag.sampleCount < 3 ? ' · 样本较少' : ''}</text></g> })}
      </svg>
    </div>
    <div className="opportunity-candidates">
      <div className="section-heading candidate-heading"><div><p className="section-kicker">CANDIDATE DIRECTIONS</p><h3>潜在内容机会</h3></div><span className="result-count">最多展示 5 个</span></div>
      {hasCandidates ? <div className="candidate-grid">{candidates.map((candidate) => <OpportunityCard key={candidate.label} candidate={candidate} baselineHighRate={baselineHighRate} />)}</div> : <p className="empty-state opportunity-empty">当前样本中暂未发现同时满足低供给与高表现条件的标签。</p>}
    </div>
    <p className="insight-note">机会指数仅用于当前样本内部排序。机会矩阵基于当前 {validRows.length} 篇研究样本的相对表现与内容供给计算，用于发现候选测试方向。样本为研究样本而非全平台随机样本，结果不代表小红书整体内容分布或因果关系。</p>
  </section>
}

function OpportunityCard({ candidate, baselineHighRate }) {
  return <article className="opportunity-card">
    <div className="candidate-title"><h4>{candidate.label}</h4><span>潜在内容机会</span></div>
    <div className="candidate-metrics"><div><small>高表现率</small><strong>{(candidate.highRate * 100).toFixed(1)}%</strong></div><div><small>整体基准</small><strong>{(baselineHighRate * 100).toFixed(1)}%</strong></div><div><small>样本数</small><strong>{candidate.sampleCount}</strong></div><div><small>内容供给率</small><strong>{(candidate.supplyRate * 100).toFixed(1)}%</strong></div><div><small>相对基准</small><strong>{candidate.performanceLift >= 0 ? '+' : ''}{(candidate.performanceLift * 100).toFixed(1)}pct</strong></div><div><small>机会指数</small><strong>{candidate.opportunityScore.toFixed(3)}</strong></div></div>
    <div className="representative-notes"><small>代表内容</small>{candidate.representativeNotes.map((row, index) => <div key={`${row['笔记标题']}-${index}`}><span>{row['笔记链接'] && row['笔记链接'] !== '—' ? <a href={row['笔记链接']} target="_blank" rel="noreferrer">{row['笔记标题']}</a> : row['笔记标题']}</span><em>{formatNumber(row['点赞数'])} 赞 · {formatNumber(row['收藏数'])} 藏</em></div>)}</div>
  </article>
}

function CrossDimensionOpportunities({ stats }) {
  const validRows = stats[0]?.validRows || []
  const baselineHighRate = stats[0]?.baselineHighRate || 0
  const allCandidates = stats.flatMap(({ dimension, candidates }) => candidates.map((candidate) => ({ ...candidate, dimension }))).sort((left, right) => right.opportunityScore - left.opportunityScore).slice(0, 10)
  return <section className="cross-opportunity-section">
    <div className="section-heading insights-heading"><div><p className="section-kicker">CROSS-DIMENSION OPPORTUNITIES</p><h2>跨维度机会发现</h2><p className="opportunity-subtitle">从7个内容维度中统一筛选「低供给 × 高表现」候选方向</p></div></div>
    <div className="cross-summary"><StatCard label="扫描维度" value={stats.length} suffix="个" /><StatCard label="有效标签" value={stats.reduce((sum, item) => sum + item.tags.length, 0)} /><StatCard label="候选机会" value={stats.reduce((sum, item) => sum + item.candidates.length, 0)} /><StatCard label="整体高表现率" value={`${(baselineHighRate * 100).toFixed(1)}%`} /></div>
    <div className="cross-content">
      <div className="cross-ranking"><div className="section-heading candidate-heading"><div><p className="section-kicker">OPPORTUNITY RANKING</p><h3>候选机会排行榜</h3></div><span className="result-count">最多展示 10 个</span></div>{allCandidates.length ? <div className="ranking-list">{allCandidates.map((candidate, index) => <div className="ranking-row" key={`${candidate.dimension}-${candidate.label}`}><strong className="ranking-number">{String(index + 1).padStart(2, '0')}</strong><span className="ranking-dimension">{candidate.dimension}</span><b>{candidate.label}</b><span>样本 {candidate.sampleCount}</span><span>高表现 {candidate.highCount}</span><span>{(candidate.supplyRate * 100).toFixed(1)}%</span><span>{(candidate.highRate * 100).toFixed(1)}%</span><span className="positive-rate">+{(candidate.performanceLift * 100).toFixed(1)}pct</span><strong>{candidate.opportunityScore.toFixed(3)}</strong></div>)}</div> : <p className="empty-state opportunity-empty">暂无符合当前标准的候选方向</p>}</div>
      <div className="dimension-overview"><div className="section-heading candidate-heading"><div><p className="section-kicker">DIMENSION OVERVIEW</p><h3>维度概览</h3></div></div>{stats.map((item) => { const best = [...item.candidates].sort((left, right) => right.opportunityScore - left.opportunityScore)[0]; return <div className="dimension-card" key={item.dimension}><b>{item.dimension}</b><span>标签 {item.tags.length} · 候选 {item.candidates.length}</span><small>{best ? `最高机会指数：${best.label} · ${best.opportunityScore.toFixed(3)}` : '暂无候选'}</small></div> })}</div>
    </div>
    <p className="insight-note">该结果基于当前研究样本内部的相对供给与表现差异，用于生成后续内容测试假设，不代表平台整体内容供给，也不代表因果关系。</p>
  </section>
}

function calculateCombinationStats(rows, group) {
  const validRows = rows.filter((row) => row['样本类型'] === '高表现' || row['样本类型'] === '普通')
  const baselineHighRate = validRows.length ? validRows.filter((row) => row['样本类型'] === '高表现').length / validRows.length : 0
  const combinations = new Map()
  validRows.forEach((row) => {
    const leftLabels = Array.isArray(row[group.left]) ? new Set(row[group.left].filter((label) => label && label !== '待分析')) : new Set()
    const rightLabels = Array.isArray(row[group.right]) ? new Set(row[group.right].filter((label) => label && label !== '待分析')) : new Set()
    leftLabels.forEach((left) => rightLabels.forEach((right) => {
      const key = `${left}::${right}`
      const current = combinations.get(key) || { key, left, right, rows: [], sampleCount: 0, highCount: 0, normalCount: 0 }
      if (!current.rows.includes(row)) current.rows.push(row)
      current.sampleCount += 1
      if (row['样本类型'] === '高表现') current.highCount += 1
      else current.normalCount += 1
      combinations.set(key, current)
    }))
  })
  const items = [...combinations.values()].map((item) => {
    const highRate = item.sampleCount ? item.highCount / item.sampleCount : 0
    const supplyRate = validRows.length ? item.sampleCount / validRows.length : 0
    const performanceLift = highRate - baselineHighRate
    return { ...item, highRate, supplyRate, baselineHighRate, performanceLift, opportunityScore: performanceLift * (1 - supplyRate), strategy: group.template(item.left, item.right) }
  })
  const candidates = items.filter((item) => item.sampleCount >= 3 && item.highRate > baselineHighRate).sort((left, right) => right.opportunityScore - left.opportunityScore)
  return { validRows, baselineHighRate, items, candidates }
}

function CombinationOpportunities({ stats }) {
  const [activeKey, setActiveKey] = useState(stats[0]?.key)
  const active = stats.find((item) => item.key === activeKey) || stats[0]
  return <section className="combination-section">
    <div className="section-heading insights-heading"><div><p className="section-kicker">COMBINATION OPPORTUNITIES</p><h2>内容组合机会分析</h2><p className="opportunity-subtitle">从「人群 × 需求 × 表达方式」进一步识别可测试的内容组合</p></div></div>
    <div className="insight-tabs combination-tabs" role="tablist">{stats.map((item) => <button key={item.key} type="button" className={`insight-tab ${active?.key === item.key ? 'is-active' : ''}`} onClick={() => setActiveKey(item.key)}>{item.label}</button>)}</div>
    <div className="combination-table table-frame"><table><thead><tr><th>排名</th><th>组合</th><th>样本数</th><th>高表现</th><th>普通</th><th>高表现率</th><th>整体基准</th><th>相对基准</th><th>供给率</th><th>机会指数</th></tr></thead><tbody>{active?.candidates.slice(0, 10).map((item, index) => <CombinationRow key={item.key} item={item} index={index} />)}</tbody></table>{!active?.candidates.length && <div className="empty-state">暂无符合当前标准的候选方向</div>}</div>
    {active?.candidates.slice(0, 10).map((item) => <CombinationDetail key={item.key} item={item} />)}
    <p className="insight-note">组合分析用于从当前研究样本中生成内容测试假设。二维标签组合会进一步降低样本量，因此结果应优先用于 A/B 测试方向筛选，而非视为稳定规律或因果关系。</p>
  </section>
}

function CombinationRow({ item, index }) {
  return <tr><td>{String(index + 1).padStart(2, '0')}</td><td className="combination-name">{item.left} × {item.right}{item.sampleCount === 3 && <span className="small-sample-tag">小样本</span>}</td><td>{item.sampleCount}</td><td>{item.highCount}</td><td>{item.normalCount}</td><td>{(item.highRate * 100).toFixed(1)}%</td><td>{(item.baselineHighRate * 100).toFixed(1)}%</td><td className="positive-rate">+{(item.performanceLift * 100).toFixed(1)}pct</td><td>{(item.supplyRate * 100).toFixed(1)}%</td><td className="opportunity-score">{item.opportunityScore.toFixed(3)}</td></tr>
}

function CombinationDetail({ item }) {
  const [expanded, setExpanded] = useState(false)
  const notes = item.rows.filter((row) => row['样本类型'] === '高表现').sort((left, right) => ((Number(right['点赞数']) || 0) + (Number(right['收藏数']) || 0)) - ((Number(left['点赞数']) || 0) + (Number(left['收藏数']) || 0))).slice(0, 3)
  return <article className={`combination-detail ${expanded ? 'is-expanded' : ''}`}><button type="button" onClick={() => setExpanded(!expanded)}><strong>{item.left} × {item.right}</strong><span>{expanded ? '收起代表内容' : '查看代表内容'} · {item.strategy}</span></button>{expanded && <div className="combination-notes">{notes.map((row, index) => <div key={`${row['笔记标题']}-${index}`}><span>{row['笔记链接'] && row['笔记链接'] !== '—' ? <a href={row['笔记链接']} target="_blank" rel="noreferrer">{row['笔记标题']}</a> : row['笔记标题']}</span><em>{row['账号名称']} · {row['样本类型']} · {formatNumber(row['点赞数'])} 赞 · {formatNumber(row['收藏数'])} 藏</em></div>)}</div>}</article>
}

createRoot(document.getElementById('root')).render(<App />)

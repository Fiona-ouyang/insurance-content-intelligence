import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import demoRows from './data/demoRows.json'
import demoAnalysis from './data/demoAnalysis.json'
import './dashboard.css'

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
const NAV_ITEMS = [
  { key: 'overview', label: '首页概览', icon: 'home', description: '核心指标与内容表现概览' },
  { key: 'analysis', label: '内容分析', icon: 'analysis', description: '查看笔记 AI 标签分析结果' },
  { key: 'accounts', label: '账号对比', icon: 'accounts', description: '按账号筛选并查看原始笔记指标' },
  { key: 'opportunities', label: '选题发现', icon: 'opportunities', description: '查看机会矩阵与可测试方向' },
  { key: 'insights', label: '标签洞察', icon: 'insights', description: '切换七个标签维度观察表现' },
  { key: 'notes', label: '笔记列表', icon: 'notes', description: '筛选与浏览研究样本' },
  { key: 'import', label: '数据导入', icon: 'import', description: '导入 Excel 或 CSV 研究数据' },
]

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

function getEngagementScore(row) {
  return (Number(row['点赞数']) || 0) + (Number(row['收藏数']) || 0) + (Number(row['评论数']) || 0)
}

// 优先高表现样本，同类再按互动量排序，取前 3 条用于 OpportunityCard 展示
function getRepresentativeNotes(tagRows) {
  return [...tagRows]
    .sort((left, right) => {
      const leftHigh = left['样本类型'] === '高表现' ? 1 : 0
      const rightHigh = right['样本类型'] === '高表现' ? 1 : 0
      if (leftHigh !== rightHigh) return rightHigh - leftHigh
      return getEngagementScore(right) - getEngagementScore(left)
    })
    .slice(0, 3)
    .map((row) => ({ 笔记标题: row['笔记标题'], 笔记链接: row['笔记链接'], 点赞数: row['点赞数'], 收藏数: row['收藏数'] }))
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
    return { ...tag, highRate, supplyRate, performanceLift, opportunityScore: performanceLift * (1 - supplyRate), representativeNotes: getRepresentativeNotes(tag.rows) }
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
  const [activeModule, setActiveModule] = useState('overview')
  const [noteSampleFilter, setNoteSampleFilter] = useState('全部')
  const [isDemoDataset] = useState(() => !localStorage.getItem(DATASET_STORAGE_KEY))
  const fileInputRef = useRef(null)

  const accounts = useMemo(() => [...new Set(rows.map((row) => row['账号名称']).filter(Boolean))], [rows])
  const filteredRows = useMemo(
    () => selectedAccounts.length ? rows.filter((row) => selectedAccounts.includes(row['账号名称'])) : rows,
    [rows, selectedAccounts],
  )
  const stats = useMemo(() => {
    const validRows = rows.filter((row) => row['样本类型'] === '高表现' || row['样本类型'] === '普通')
    const highPerformance = validRows.filter((row) => row['样本类型'] === '高表现').length
    return {
      accounts: accounts.length,
      notes: rows.length,
      highPerformance,
      ordinary: validRows.length - highPerformance,
      highPerformanceRate: validRows.length ? highPerformance / validRows.length : 0,
    }
  }, [accounts.length, rows])
  const noteRows = useMemo(() => filteredRows.filter((row) => noteSampleFilter === '全部' || row['样本类型'] === noteSampleFilter), [filteredRows, noteSampleFilter])
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

  const activePage = NAV_ITEMS.find((item) => item.key === activeModule) || NAV_ITEMS[0]
  useEffect(() => { window.scrollTo({ top: 0 }) }, [activeModule])
  return (
    <div className="app-layout">
      <aside className="sidebar">
        <button className="sidebar-brand" type="button" onClick={() => setActiveModule('overview')} aria-label="返回首页概览">
          <span className="brand-mark"><i /><i /><i /></span>
          <span className="sidebar-brand-copy"><strong>小红书内容增长研究平台</strong></span>
        </button>
        <p className="nav-caption">工作空间</p>
        <nav className="side-nav" aria-label="主导航">
          {NAV_ITEMS.map((item) => <button key={item.key} type="button" className={`side-nav-item ${activeModule === item.key ? 'is-active' : ''}`} aria-current={activeModule === item.key ? 'page' : undefined} onClick={() => setActiveModule(item.key)}>
            <LineIcon name={item.icon} /><span>{item.label}</span>
          </button>)}
        </nav>
        <div className="sidebar-promo"><p>用数据<br />理解内容<br />让内容<br />被看见</p></div>
        <div className="sidebar-footer"><span className="local-dot" />数据仅在本地处理</div>
      </aside>

      <main className="app-shell">
        <header className="topbar">
          <div className="topbar-context"><span>工作空间</span><i>/</i><strong>{activePage.label}</strong></div>
          <div className="status-indicator"><span /> 本地研究空间</div>
        </header>

        {activeModule === 'overview' && <HomeDashboard stats={stats} isDemoDataset={isDemoDataset} rows={rows} accounts={accounts} selectedAccounts={selectedAccounts} toggleAccount={toggleAccount} clearAccounts={() => setSelectedAccounts([])} noteSampleFilter={noteSampleFilter} setNoteSampleFilter={setNoteSampleFilter} onNavigate={setActiveModule} />}
        {activeModule === 'analysis' && <AnalysisModule rows={filteredRows} analysisCoverage={analysisCoverage} totalRows={rows.length} analysisMessage={analysisMessage} analysisRunning={analysisRunning} failedIndexes={failedIndexes} handleAnalyze={handleAnalyze} />}
        {activeModule === 'accounts' && <DatabaseModule mode="accounts" rows={filteredRows} allRows={rows} accounts={accounts} selectedAccounts={selectedAccounts} toggleAccount={toggleAccount} clearAccounts={() => setSelectedAccounts([])} />}
        {activeModule === 'opportunities' && <OpportunityModule insightDimension={insightDimension} setInsightDimension={setInsightDimension} opportunityStats={opportunityStats} crossDimensionStats={crossDimensionStats} combinationStats={combinationStats} />}
        {activeModule === 'insights' && <InsightsSection insightDimension={insightDimension} setInsightDimension={setInsightDimension} insightStats={insightStats} />}
        {activeModule === 'notes' && <DatabaseModule mode="notes" rows={noteRows} allRows={rows} accounts={accounts} selectedAccounts={selectedAccounts} toggleAccount={toggleAccount} clearAccounts={() => setSelectedAccounts([])} noteSampleFilter={noteSampleFilter} setNoteSampleFilter={setNoteSampleFilter} />}
        {activeModule === 'import' && <ImportModule fileInputRef={fileInputRef} fileName={fileName} message={message} handleFile={handleFile} clearProjectData={clearProjectData} />}
      </main>
    </div>
  )
}

function LineIcon({ name }) {
  const iconPaths = {
    home: <><path d="m3.5 10 8.5-7 8.5 7" /><path d="M5.5 9v11h13V9M9 20v-6h6v6" /></>,
    analysis: <><path d="M4 19.5h16M6.5 16v-5M12 16V5M17.5 16V8" /><path d="m5 8 5-3 5 2 4-4" /></>,
    accounts: <><circle cx="9" cy="8" r="3" /><path d="M3.5 20v-1.5a5.5 5.5 0 0 1 11 0V20M16 5.5a3 3 0 0 1 0 5.8M17 14a4.5 4.5 0 0 1 3.5 4.4V20" /></>,
    opportunities: <><path d="m12 3 1.6 5.4L19 10l-5.4 1.6L12 17l-1.6-5.4L5 10l5.4-1.6L12 3Z" /><path d="m19 15 .8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8L19 15Z" /></>,
    insights: <><path d="M4 5.5h16v13H4z" /><path d="M8 14v-3M12 14V8M16 14v-5" /></>,
    notes: <><path d="M6 4h12v16H6zM9 8h6M9 12h6M9 16h4" /></>,
    import: <><path d="M12 15V3M7.5 7.5 12 3l4.5 4.5" /><path d="M4 14v6h16v-6" /></>,
    heart: <path d="M12 20s-7-4.4-7-10a4 4 0 0 1 7-2.5A4 4 0 0 1 19 10c0 5.6-7 10-7 10Z" />,
    chat: <path d="M4 5h16v11H9l-5 4V5Z" />,
    star: <path d="m12 3.5 2.5 5.2 5.7.8-4.1 4 1 5.7-5.1-2.7-5.1 2.7 1-5.7-4.1-4 5.7-.8L12 3.5Z" />,
  }
  return <svg className="nav-icon" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">{iconPaths[name]}</svg>
}

const average = (rows, field) => rows.length ? Math.round(rows.reduce((sum, row) => sum + (Number(row[field]) || 0), 0) / rows.length) : 0
const TOPIC_COLORS = ['#5b8def', '#b266d9', '#ff5c93', '#7fdcc4', '#ffb27a', '#9d8cf0', '#d3d8e6']
const RANK_COLORS = ['#ff5c93', '#5b8def', '#ffa45c', '#2fcf9f', '#9d6cf0']

function HomeDashboard({ stats, isDemoDataset, rows, accounts, selectedAccounts, toggleAccount, clearAccounts, noteSampleFilter, setNoteSampleFilter, onNavigate }) {
  return <div className="home-dashboard">
    <section className="hero">
      <div className="hero-copy-block">
        <span className="hero-badge"><LineIcon name="home" />基于真实笔记数据的内容研究</span>
        <h1>洞察内容，<span className="gradient-text">预见增长</span></h1>
        <p className="hero-copy">从真实笔记中发现内容规律，结合数据分析与用户需求，探索更有潜力的选题方向。</p>
        <div className="hero-chips"><span><LineIcon name="insights" />内容表现分析</span><span><LineIcon name="accounts" />用户需求洞察</span><span><LineIcon name="opportunities" />选题机会挖掘</span></div>
      </div>
      <div className="hero-art" aria-hidden="true"><b className="art-bubble b1">什么内容更受欢迎？</b><b className="art-bubble b2">用户关心什么？</b><b className="art-bubble b3">下一个机会是什么？</b><div className="art-card"><span>小红书</span></div><i className="art-wave" /></div>
    </section>
    {isDemoDataset && <div className="status-indicator demo-status-indicator"><span /> 当前展示 Demo 样本</div>}
    <section className="metric-grid" aria-label="核心指标">
      <OverviewMetric icon="notes" tone="blue" label="总笔记数" value={stats.notes} sub={`来自 ${stats.accounts} 个账号`} />
      <OverviewMetric icon="heart" tone="pink" label="平均点赞" value={average(rows, '点赞数')} sub={`高表现笔记：${stats.highPerformance} 篇`} />
      <OverviewMetric icon="chat" tone="indigo" label="平均评论" value={average(rows, '评论数')} sub={`高表现笔记：${stats.highPerformance} 篇`} />
      <OverviewMetric icon="star" tone="orange" label="平均收藏" value={average(rows, '收藏数')} sub={`高表现笔记：${stats.highPerformance} 篇`} />
      <OverviewMetric icon="insights" tone="violet" label="高表现笔记占比" value={`${(stats.highPerformanceRate * 100).toFixed(1)}%`} sub={`${stats.highPerformance} / ${stats.highPerformance + stats.ordinary}`} />
    </section>
    <OverviewCharts rows={rows} onNavigate={onNavigate} />
    <NotesShowcase rows={rows} accounts={accounts} selectedAccounts={selectedAccounts} toggleAccount={toggleAccount} clearAccounts={clearAccounts} noteSampleFilter={noteSampleFilter} setNoteSampleFilter={setNoteSampleFilter} onNavigate={onNavigate} />
  </div>
}

function OverviewMetric({ icon, tone, label, value, sub }) {
  return <article className="overview-metric">
    <span className={`metric-icon ${tone}`}><LineIcon name={icon} /></span>
    <span className="metric-label">{label}</span>
    <strong>{typeof value === 'string' ? value : formatNumber(value)}</strong>
    <small className="metric-sub">{sub}</small>
  </article>
}

function OverviewCharts({ rows, onNavigate }) {
  const validRows = rows.filter((row) => row['样本类型'] === '高表现' || row['样本类型'] === '普通')
  const accountStats = [...new Set(rows.map((row) => row['账号名称']).filter(Boolean))].map((name) => {
    const own = rows.filter((row) => row['账号名称'] === name)
    const ownValid = own.filter((row) => row['样本类型'] === '高表现' || row['样本类型'] === '普通')
    return { name, count: own.length, highCount: ownValid.filter((row) => row['样本类型'] === '高表现').length, validCount: ownValid.length, likes: average(own, '点赞数'), comments: average(own, '评论数'), rate: ownValid.length ? ownValid.filter((row) => row['样本类型'] === '高表现').length / ownValid.length : 0 }
  }).sort((left, right) => right.count - left.count).slice(0, 5)
  const topicTags = calculateOpportunityStats(rows, '内容主题').tags
  const topicSorted = [...topicTags].sort((left, right) => right.sampleCount - left.sampleCount)
  const topicSlices = topicSorted.slice(0, 6).map((tag) => ({ label: tag.label, count: tag.sampleCount }))
  const otherCount = topicSorted.slice(6).reduce((sum, tag) => sum + tag.sampleCount, 0)
  if (otherCount) topicSlices.push({ label: '其他', count: otherCount })
  const topTopics = topicTags.filter((tag) => tag.sampleCount >= 3).sort((left, right) => right.highRate - left.highRate || right.sampleCount - left.sampleCount).slice(0, 5)
  return <section className="overview-charts" aria-label="内容表现图表">
    <article className="chart-panel account-panel">
      <div className="panel-heading"><h2>账号表现对比</h2><div className="interaction-legend"><span><i className="dot-blue" />平均点赞</span><span><i className="dot-purple" />平均评论</span><span><i className="line-legend" />样本内高表现率</span></div></div>
      <AccountChart data={accountStats} />
      <p className="panel-foot">基于已采集样本统计，不代表账号全部笔记表现。仅展示样本量最多的前 5 个账号；样本内高表现率 = 该账号高表现笔记数 ÷ 该账号样本总数。</p>
    </article>
    <article className="chart-panel topic-donut-panel">
      <div className="panel-heading"><h2>内容主题分布</h2></div>
      {topicSlices.length ? <TopicDonut slices={topicSlices} total={rows.length} /> : <div className="empty-state">暂无内容主题标签</div>}
    </article>
    <article className="chart-panel ranking-panel">
      <div className="panel-heading"><h2>高表现内容 TOP主题</h2><button type="button" className="link-button" onClick={() => onNavigate('insights')}>查看全部 →</button></div>
      {topTopics.length ? <ol className="topic-rank-list">{topTopics.map((tag, index) => <li key={tag.label}>
        <span className="rank-badge" style={{ background: RANK_COLORS[index] }}>{index + 1}</span>
        <span className="rank-name">{tag.label}</span>
        <div className="rank-track"><i style={{ width: `${tag.highRate * 100}%`, background: RANK_COLORS[index] }} /></div>
        <strong>{(tag.highRate * 100).toFixed(1)}%</strong>
      </li>)}</ol> : <div className="empty-state">暂无足够样本的主题</div>}
      <p className="panel-foot">按高表现率排序，仅含样本数 ≥ 3 的主题</p>
    </article>
  </section>
}

function TopicDonut({ slices, total }) {
  const sum = slices.reduce((acc, slice) => acc + slice.count, 0)
  const radius = 52
  const circumference = 2 * Math.PI * radius
  let offset = 0
  return <div className="donut-layout">
    <svg className="topic-donut" viewBox="0 0 140 140" role="img" aria-label="内容主题分布环形图">
      <g transform="rotate(-90 70 70)">{slices.map((slice, index) => {
        const length = (slice.count / sum) * circumference
        const circle = <circle key={slice.label} cx="70" cy="70" r={radius} fill="none" stroke={TOPIC_COLORS[index % TOPIC_COLORS.length]} strokeWidth="24" strokeDasharray={`${length} ${circumference - length}`} strokeDashoffset={-offset}><title>{`${slice.label}：${slice.count} 篇`}</title></circle>
        offset += length
        return circle
      })}</g>
      <text x="70" y="68" textAnchor="middle" className="donut-total">{total}</text><text x="70" y="85" textAnchor="middle" className="donut-sub">篇笔记</text>
    </svg>
    <ul className="donut-legend">{slices.map((slice, index) => <li key={slice.label}><i style={{ background: TOPIC_COLORS[index % TOPIC_COLORS.length] }} /><span>{slice.label}</span><b>{slice.count} ({total ? (slice.count / total * 100).toFixed(1) : 0}%)</b></li>)}</ul>
  </div>
}

function AccountChart({ data }) {
  if (!data.length) return <div className="empty-state">暂无账号数据</div>
  const width = 460
  const height = 190
  const left = 40
  const right = 34
  const top = 10
  const bottom = 156
  const plotWidth = width - left - right
  const plotHeight = bottom - top
  const maxValue = Math.max(...data.flatMap((item) => [item.likes, item.comments]), 1)
  const rough = maxValue / 4
  const magnitude = 10 ** Math.floor(Math.log10(rough))
  const niceMax = Math.ceil(rough / magnitude) * magnitude * 4
  const slot = plotWidth / data.length
  const barWidth = Math.min(20, slot * .26)
  const points = data.map((item, index) => ({ x: left + slot * (index + .5), y: bottom - item.rate * plotHeight }))
  const linePath = points.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ')
  return <svg className="account-chart" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="各账号平均点赞、平均评论与样本内高表现率">
    <defs><linearGradient id="barBlue" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#6aa0ff" /><stop offset="1" stopColor="#5b7fe8" /></linearGradient><linearGradient id="barPurple" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#c9a0f2" /><stop offset="1" stopColor="#a67de6" /></linearGradient></defs>
    {[0, 1, 2, 3, 4].map((step) => <g key={step}><line x1={left} x2={width - right} y1={bottom - plotHeight * step / 4} y2={bottom - plotHeight * step / 4} className="chart-gridline" />
      <text x={left - 6} y={bottom - plotHeight * step / 4 + 3} textAnchor="end" className="monthly-label">{formatNumber(Math.round(niceMax * step / 4))}</text>
      <text x={width - right + 6} y={bottom - plotHeight * step / 4 + 3} className="monthly-label">{step * 25}%</text></g>)}
    {data.map((item, index) => {
      const x = points[index].x
      const likeHeight = (item.likes / niceMax) * plotHeight
      const commentHeight = (item.comments / niceMax) * plotHeight
      return <g key={item.name}><title>{`${item.name}：${item.count} 篇｜平均点赞 ${formatNumber(item.likes)}｜平均评论 ${formatNumber(item.comments)}｜样本内高表现率 ${(item.rate * 100).toFixed(1)}%（${item.highCount}/${item.validCount}）`}</title>
        <rect x={x - barWidth - 1} y={bottom - likeHeight} width={barWidth} height={likeHeight} rx="4" fill="url(#barBlue)" />
        <rect x={x + 1} y={bottom - commentHeight} width={barWidth} height={commentHeight} rx="4" fill="url(#barPurple)" />
        <text x={x} y={bottom + 15} textAnchor="middle" className="monthly-label">{item.name.length > 5 ? `${item.name.slice(0, 4)}…` : item.name}</text>
      </g>
    })}
    <path d={linePath} className="monthly-line" />
    {points.map((point, index) => <circle key={data[index].name} cx={point.x} cy={point.y} r="3.5" className="monthly-point" />)}
  </svg>
}

function NotesShowcase({ rows, accounts, selectedAccounts, toggleAccount, clearAccounts, noteSampleFilter, setNoteSampleFilter, onNavigate }) {
  const [sortKey, setSortKey] = useState('点赞数')
  const filters = ['全部', '高表现', '普通']
  const countOf = (filter) => rows.filter((row) => filter === '全部' || row['样本类型'] === filter).length
  const dateValue = (row) => { const time = new Date(row['发布时间']).getTime(); return Number.isNaN(time) ? 0 : time }
  const visible = rows
    .filter((row) => (selectedAccounts.length === 0 || selectedAccounts.includes(row['账号名称'])) && (noteSampleFilter === '全部' || row['样本类型'] === noteSampleFilter))
    .sort((left, right) => sortKey === '发布时间' ? dateValue(right) - dateValue(left) : (Number(right[sortKey]) || 0) - (Number(left[sortKey]) || 0))
  return <section className="notes-showcase">
    <div className="showcase-head">
      <h2>高表现笔记展示</h2>
      <div className="segmented-control" role="group" aria-label="按样本类型筛选笔记">{filters.map((filter) => <button key={filter} type="button" className={noteSampleFilter === filter ? 'is-active' : ''} aria-pressed={noteSampleFilter === filter} onClick={() => setNoteSampleFilter(filter)}>{filter} ({countOf(filter)})</button>)}</div>
      <div className="showcase-selects">
        <select aria-label="排序方式" value={sortKey} onChange={(event) => setSortKey(event.target.value)}><option value="点赞数">按点赞数排序</option><option value="收藏数">按收藏数排序</option><option value="评论数">按评论数排序</option><option value="发布时间">按发布时间排序</option></select>
        <select aria-label="账号筛选" value={selectedAccounts.length === 1 ? selectedAccounts[0] : ''} onChange={(event) => { clearAccounts(); if (event.target.value) toggleAccount(event.target.value) }}><option value="">全部账号</option>{accounts.map((account) => <option key={account} value={account}>{account}</option>)}</select>
      </div>
      <button type="button" className="link-button" onClick={() => onNavigate('notes')}>查看全部 →</button>
    </div>
    {visible.length ? <div className="note-card-grid">{visible.slice(0, 6).map((row, index) => <NotePreviewCard row={row} index={index} key={`${getRowKey(row)}-${index}`} />)}</div> : <div className="empty-state note-empty">当前筛选条件下没有笔记</div>}
  </section>
}

const COVER_TONES = ['tone-a', 'tone-b', 'tone-c', 'tone-d']

function NotePreviewCard({ row, index }) {
  const tags = analysisColumns.flatMap((column) => Array.isArray(row[column]) ? row[column] : [row[column]]).filter((label) => label && label !== '待分析').slice(0, 2)
  const link = row['笔记链接']
  const hasLink = !isMissing(link) && link !== '—'
  return <article className="note-preview-card">
    <div className={`text-cover ${COVER_TONES[index % COVER_TONES.length]}`}><span>{displayText(row['样本类型'])}</span><strong>{displayText(row['笔记标题'])}</strong></div>
    <div className="note-preview-body">
      <h3>{displayText(row['笔记标题'])}</h3>
      <div className="note-preview-metrics"><span><LineIcon name="heart" />{formatNumber(row['点赞数'])}</span><span><LineIcon name="chat" />{formatNumber(row['评论数'])}</span><span><LineIcon name="star" />{formatNumber(row['收藏数'])}</span></div>
      <div className="note-preview-tags">{tags.map((tag, tagIndex) => <span key={`${tag}-${tagIndex}`} className={`tag-${tagIndex}`}>{tag}</span>)}</div>
      <div className="note-preview-foot"><span>{displayText(row['账号名称'])} · {row['发布时间'] instanceof Date ? row['发布时间'].toLocaleDateString('zh-CN') : displayText(row['发布时间'])}</span>{hasLink && <a href={link} target="_blank" rel="noreferrer" aria-label="查看原文">↗</a>}</div>
    </div>
  </article>
}

function AccountFilterBar({ accounts, selectedAccounts, toggleAccount, clearAccounts }) {
  return <div className="account-filter-bar">
    <span>账号</span>
    <div className="account-filters">
      <button type="button" className={`account-chip ${selectedAccounts.length === 0 ? 'is-selected' : ''}`} onClick={clearAccounts}>全部账号</button>
      {accounts.map((account, index) => <button key={account} type="button" className={`account-chip ${selectedAccounts.includes(account) ? 'is-selected' : ''}`} aria-pressed={selectedAccounts.includes(account)} onClick={() => toggleAccount(account)}>
        <span className={`chip-dot ${accountColors[index % accountColors.length]}`} />{account}
      </button>)}
    </div>
    {selectedAccounts.length > 0 && <small>已选 {selectedAccounts.length} 个</small>}
  </div>
}

function DatabaseModule({ mode, rows, allRows, accounts, selectedAccounts, toggleAccount, clearAccounts, noteSampleFilter, setNoteSampleFilter }) {
  const isCompare = mode === 'accounts'
  return <section className="module-content database-section">
    <div className="module-heading"><div><p className="section-kicker">{isCompare ? '账号筛选与笔记对照' : '研究样本数据库'}</p><h1>{isCompare ? '账号对比' : '笔记列表'}</h1><p>{isCompare ? '选择一个或多个账号并直接对照笔记原始互动指标；当前不生成账号均值、粉丝效率或增长趋势等衍生指标。' : '按账号和样本类型筛选当前研究数据。'}</p></div><span className="result-count">显示 {rows.length} / {allRows.length} 篇</span></div>
    <AccountFilterBar accounts={accounts} selectedAccounts={selectedAccounts} toggleAccount={toggleAccount} clearAccounts={clearAccounts} />
    {!isCompare && <div className="segmented-control note-type-control" role="group" aria-label="按样本类型筛选笔记">{['全部', '高表现', '普通'].map((filter) => <button key={filter} type="button" className={noteSampleFilter === filter ? 'is-active' : ''} aria-pressed={noteSampleFilter === filter} onClick={() => setNoteSampleFilter(filter)}>{filter}</button>)}</div>}
    <div className="table-frame data-table-frame"><div className="table-scroll"><table><thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead><tbody>{rows.map((row, index) => <TableRow key={`${getRowKey(row)}-${index}`} row={row} />)}</tbody></table>{!rows.length && <div className="empty-state">当前筛选条件下暂无数据</div>}</div></div>
    <p className="table-note">可见 {rows.length} 篇 · 笔记链接可在新窗口打开原文</p>
  </section>
}

function AnalysisModule({ rows, analysisCoverage, totalRows, analysisMessage, analysisRunning, failedIndexes, handleAnalyze }) {
  return <section className="module-content analysis-section">
    <div className="module-heading analysis-module-heading"><div><p className="section-kicker">七维内容标签</p><h1>内容分析</h1><p>查看真实笔记对应的 AI 标签结果。标签字段和分析逻辑沿用当前数据。</p></div>
      <div className="analysis-actions">{analysisMessage && <span className="analysis-message">{analysisMessage.split('\n').map((line) => <span key={line}>{line}<br /></span>)}</span>}
        {IS_PUBLIC_DEMO ? <span className="analysis-message">在线 Demo 暂不开放 AI 重新分析，当前展示为已完成分析的研究样本。</span> : <><button className="analysis-button" type="button" disabled={analysisRunning} onClick={() => handleAnalyze(false)}>{analysisRunning ? 'AI分析中...' : '开始AI分析'}</button>{!analysisRunning && failedIndexes.length > 0 && <button className="retry-button" type="button" onClick={() => handleAnalyze(true)}>仅重试失败项（{failedIndexes.length}）</button>}</>}
      </div>
    </div>
    <div className="table-frame analysis-table-frame"><div className="table-scroll"><table className="analysis-table"><thead><tr>{analysisTableColumns.map((column) => <th key={column}>{column}</th>)}</tr></thead><tbody>{rows.map((row, index) => <AnalysisTableRow key={`${getRowKey(row)}-${index}`} row={row} />)}</tbody></table>{!rows.length && <div className="empty-state">当前筛选条件下暂无数据</div>}</div></div>
    <p className="table-note">当前显示 {rows.length} 篇 · 全部样本标签覆盖 {analysisCoverage} / {totalRows} 篇</p>
  </section>
}

function OpportunityModule({ insightDimension, setInsightDimension, opportunityStats, crossDimensionStats, combinationStats }) {
  return <div className="module-content opportunity-module">
    <div className="module-heading"><div><p className="section-kicker">内容方向研究</p><h1>选题发现</h1><p>基于现有七维标签与样本表现，查看机会矩阵、跨维度机会和内容组合。</p></div></div>
    <OpportunityMatrixSection insightDimension={insightDimension} setInsightDimension={setInsightDimension} opportunityStats={opportunityStats} />
    <CrossDimensionOpportunities stats={crossDimensionStats} />
    <CombinationOpportunities stats={combinationStats} />
  </div>
}

function ImportModule({ fileInputRef, fileName, message, handleFile, clearProjectData }) {
  return <section className="module-content import-module">
    <div className="module-heading"><div><p className="section-kicker">本地数据管理</p><h1>数据导入</h1><p>支持 Excel 与 CSV 文件，解析过程在浏览器本地完成。</p></div></div>
    <section className="upload-panel">
      <div className="upload-icon"><LineIcon name="import" /></div>
      <div className="upload-copy"><strong>导入笔记数据</strong><span>支持 .xlsx / .xls / .csv，必需字段：账号名称、笔记标题、发布时间、点赞数、收藏数、评论数</span>
        {fileName && <small>{fileName} · {message}</small>}{message && !fileName && <small className="error-message">{message}</small>}
      </div>
      <button className="upload-button" type="button" onClick={() => fileInputRef.current?.click()}>选择文件</button>
      <button className="clear-project-button" type="button" onClick={clearProjectData}>清空项目数据</button>
      <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(event) => handleFile(event.target.files?.[0])} />
    </section>
  </section>
}

function StatCard({ label, value, suffix = '' }) {
  const displayValue = typeof value === 'string' ? value : formatNumber(value)
  return <article className="stat-card"><span>{label}</span><strong>{displayValue}<small>{suffix}</small></strong></article>
}

function OverviewVisuals({ topicDistribution, highPerformanceRanking }) {
  const maxTopicCount = Math.max(...topicDistribution.map((item) => item.count), 1)
  return <section className="overview-visuals" aria-label="首页内容分析">
    <article className="overview-panel topic-panel">
      <div className="overview-panel-heading"><div><p className="section-kicker">内容标签</p><h2>内容主题分布</h2></div><span>按笔记标签计数</span></div>
      {topicDistribution.length ? <div className="topic-chart">{topicDistribution.map((item) => <div className="topic-row" key={item.label}>
        <span>{item.label}</span><div className="topic-track"><i style={{ width: `${(item.count / maxTopicCount) * 100}%` }} /></div><strong>{item.count}</strong>
      </div>)}</div> : <div className="empty-state">当前数据暂无已完成的内容主题标签</div>}
    </article>
    <article className="overview-panel ranking-panel">
      <div className="overview-panel-heading"><div><p className="section-kicker">高表现样本</p><h2>高表现笔记排行</h2></div><span>按总互动排序</span></div>
      {highPerformanceRanking.length ? <ol className="top-note-list">{highPerformanceRanking.map((row, index) => <li key={getRowKey(row)}>
        <span className="top-note-rank">{String(index + 1).padStart(2, '0')}</span>
        <div className="top-note-copy"><strong>{isMissing(row['笔记链接']) || row['笔记链接'] === '—' ? displayText(row['笔记标题']) : <a href={row['笔记链接']} target="_blank" rel="noreferrer">{displayText(row['笔记标题'])}</a>}</strong><small>{displayText(row['账号名称'])}</small></div>
        <span className="top-note-metric">{formatNumber(getEngagementScore(row))}<small>互动</small></span>
      </li>)}</ol> : <div className="empty-state">当前数据暂无高表现笔记</div>}
    </article>
  </section>
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
  return <section className="module-content insights-section" id="insights">
    <div className="module-heading"><div><p className="section-kicker">七维内容标签</p><h1>标签洞察</h1><p>切换七个标签维度，观察各标签的高表现率与样本分布。</p></div></div>
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
  return <section className="opportunity-section" id="opportunity">
    <div className="section-heading insights-heading">
      <div><p className="section-kicker">CONTENT OPPORTUNITY MATRIX</p><h2>内容机会矩阵</h2><p className="opportunity-subtitle">从「内容供给」与「高表现率」两个维度识别候选内容方向</p></div>
    </div>
    <div className="insight-tabs" role="tablist">
      {analysisColumns.map((dimension) => <button key={dimension} type="button" className={`insight-tab ${insightDimension === dimension ? 'is-active' : ''}`} onClick={() => setInsightDimension(dimension)}>{dimension}</button>)}
    </div>
    <MatrixChart tags={tags} candidates={candidates} baselineHighRate={baselineHighRate} medianSupply={medianSupply} maxSupply={maxSupply} />
    <p className="matrix-reading-note">阅读说明：横轴为内容供给率，纵轴为高表现率；左上象限代表低供给、高表现的候选机会，实际选题仍需结合样本量及后续测试判断。</p>
    <div className="opportunity-candidates">
      <div className="section-heading candidate-heading"><div><p className="section-kicker">CANDIDATE DIRECTIONS</p><h3>潜在内容机会</h3></div><span className="result-count">最多展示 5 个</span></div>
      {hasCandidates ? <div className="candidate-grid">{candidates.map((candidate) => <OpportunityCard key={candidate.label} candidate={candidate} baselineHighRate={baselineHighRate} />)}</div> : <p className="empty-state opportunity-empty">当前样本中暂未发现同时满足低供给与高表现条件的标签。</p>}
    </div>
    <p className="insight-note">机会指数仅用于当前样本内部排序。机会矩阵基于当前 {validRows.length} 篇研究样本的相对表现与内容供给计算，用于发现候选测试方向。样本为研究样本而非全平台随机样本，结果不代表小红书整体内容分布或因果关系。</p>
  </section>
}

const MATRIX_W = 1000
const MATRIX_H = 560
const MATRIX_PAD = { left: 76, right: 130, top: 34, bottom: 66 }

function placeMatrixLabels(points, plot) {
  const boxes = points.map((p) => ({ x1: p.x - 9, x2: p.x + 9, y1: p.y - 9, y2: p.y + 9 }))
  const placed = [
    { x1: plot.x1 + 8, x2: plot.x1 + 270, y1: plot.y1 + 14, y2: plot.y1 + 44 },
    { x1: plot.x2 - 160, x2: plot.x2 - 8, y1: plot.y1 + 14, y2: plot.y1 + 44 },
    { x1: plot.x1 + 8, x2: plot.x1 + 160, y1: plot.y2 - 34, y2: plot.y2 - 6 },
    { x1: plot.x2 - 160, x2: plot.x2 - 8, y1: plot.y2 - 34, y2: plot.y2 - 6 },
  ]
  const overlaps = (a, b) => a.x1 < b.x2 && a.x2 > b.x1 && a.y1 < b.y2 && a.y2 > b.y1
  const order = [...points].sort((a, b) => Number(b.isCandidate) - Number(a.isCandidate) || b.tag.sampleCount - a.tag.sampleCount)
  const result = new Map()
  order.forEach((p) => {
    const w = [...p.tag.label].reduce((sum, ch) => sum + (ch.charCodeAt(0) > 255 ? 15 : 8), 0) + 4
    const h = 18
    const options = [
      { anchor: 'start', x: p.x + 13, y: p.y, box: { x1: p.x + 11, x2: p.x + 15 + w, y1: p.y - h / 2, y2: p.y + h / 2 } },
      { anchor: 'end', x: p.x - 13, y: p.y, box: { x1: p.x - 15 - w, x2: p.x - 11, y1: p.y - h / 2, y2: p.y + h / 2 } },
      { anchor: 'middle', x: p.x, y: p.y - 18, box: { x1: p.x - w / 2, x2: p.x + w / 2, y1: p.y - 18 - h / 2, y2: p.y - 18 + h / 2 } },
      { anchor: 'middle', x: p.x, y: p.y + 20, box: { x1: p.x - w / 2, x2: p.x + w / 2, y1: p.y + 20 - h / 2, y2: p.y + 20 + h / 2 } },
    ]
    const fit = options.find(({ box }) => box.x1 >= plot.x1 && box.x2 <= plot.x2 && box.y1 >= plot.y1 && box.y2 <= plot.y2 && !placed.some((other) => overlaps(box, other)) && !boxes.some((dot, index) => points[index] !== p && overlaps(box, dot)))
    if (fit) { placed.push(fit.box); result.set(p.tag.label, fit) }
  })
  return result
}

function MatrixChart({ tags, candidates, baselineHighRate, medianSupply, maxSupply }) {
  const [hovered, setHovered] = useState(null)
  const plot = { x1: MATRIX_PAD.left, x2: MATRIX_W - MATRIX_PAD.right, y1: MATRIX_PAD.top, y2: MATRIX_H - MATRIX_PAD.bottom }
  const xMax = Math.max(Math.ceil(maxSupply * 10 * 1.1) / 10, 0.1)
  const yMax = Math.min(1, Math.max(Math.ceil((Math.max(...tags.map((tag) => tag.highRate), baselineHighRate, 0.1) + 0.05) * 10) / 10, 0.2))
  const sx = (value) => plot.x1 + (value / xMax) * (plot.x2 - plot.x1)
  const sy = (value) => plot.y2 - (value / yMax) * (plot.y2 - plot.y1)
  const candidateSet = new Set(candidates.map((candidate) => candidate.label))
  const points = useMemo(() => tags.map((tag) => ({ tag, x: sx(tag.supplyRate), y: sy(tag.highRate), isCandidate: candidateSet.has(tag.label) })), [tags, candidates, xMax, yMax])
  const labels = useMemo(() => placeMatrixLabels(points, plot), [points])
  const xTicks = Array.from({ length: 6 }, (_, index) => (xMax / 5) * index)
  const yTicks = Array.from({ length: 6 }, (_, index) => (yMax / 5) * index)
  const refX = sx(medianSupply)
  const refY = sy(baselineHighRate)
  const active = points.find((p) => p.tag.label === hovered)
  const hiddenCount = points.length - labels.size
  return <div className="opportunity-matrix" onMouseLeave={() => setHovered(null)}>
    <svg className="matrix-svg" viewBox={`0 0 ${MATRIX_W} ${MATRIX_H}`} role="img" aria-label="内容机会矩阵">
      <line className="matrix-axis-line" x1={plot.x1} y1={plot.y2} x2={plot.x2} y2={plot.y2} />
      <line className="matrix-axis-line" x1={plot.x1} y1={plot.y1} x2={plot.x1} y2={plot.y2} />
      {xTicks.map((value) => <text key={`x${value}`} className="matrix-tick" x={sx(value)} y={plot.y2 + 24} textAnchor="middle">{(value * 100).toFixed(0)}%</text>)}
      {yTicks.map((value) => <text key={`y${value}`} className="matrix-tick" x={plot.x1 - 10} y={sy(value) + 5} textAnchor="end">{(value * 100).toFixed(0)}%</text>)}
      <text className="matrix-axis-title" x={(plot.x1 + plot.x2) / 2} y={MATRIX_H - 14} textAnchor="middle">内容供给率（样本占比）→</text>
      <text className="matrix-axis-title" transform={`translate(18 ${(plot.y1 + plot.y2) / 2}) rotate(-90)`} textAnchor="middle">高表现率 →</text>
      <line className="matrix-reference-line" x1={refX} y1={plot.y1} x2={refX} y2={plot.y2} />
      <line className="matrix-reference-line" x1={plot.x1} y1={refY} x2={plot.x2} y2={refY} />
      <text className="matrix-reference-label" x={refX} y={plot.y1 - 10} textAnchor="middle">供给中位数 {(medianSupply * 100).toFixed(1)}%</text>
      <text className="matrix-reference-label" x={plot.x2 + 8} y={refY + 5} textAnchor="start">整体基准 {(baselineHighRate * 100).toFixed(1)}%</text>
      <text className="matrix-quadrant-label" x={plot.x1 + 12} y={plot.y1 + 36}>低供给 · 高表现（潜在机会）</text>
      <text className="matrix-quadrant-label" x={plot.x2 - 12} y={plot.y1 + 36} textAnchor="end">高供给 · 高表现</text>
      <text className="matrix-quadrant-label" x={plot.x1 + 12} y={plot.y2 - 12}>低供给 · 低表现</text>
      <text className="matrix-quadrant-label" x={plot.x2 - 12} y={plot.y2 - 12} textAnchor="end">高供给 · 低表现</text>
      {points.map((p) => {
        const label = labels.get(p.tag.label)
        return <g key={p.tag.label} className={`matrix-svg-point ${p.isCandidate ? 'is-candidate' : ''} ${p.tag.sampleCount < 3 ? 'is-small-sample' : ''} ${hovered === p.tag.label ? 'is-active' : ''}`} onMouseEnter={() => setHovered(p.tag.label)} onFocus={() => setHovered(p.tag.label)} onBlur={() => setHovered(null)} tabIndex={0}>
          <circle cx={p.x} cy={p.y} r="14" fill="transparent" />
          <circle className="matrix-dot" cx={p.x} cy={p.y} r="7" />
          {label && <text className="matrix-point-label" x={label.x} y={label.y + 5} textAnchor={label.anchor}>{p.tag.label}</text>}
        </g>
      })}
    </svg>
    {hiddenCount > 0 && <span className="matrix-hidden-hint">{hiddenCount} 个标签因位置相近已隐藏，悬停圆点查看</span>}
    {active && <div className={`matrix-tooltip ${active.y < 150 ? 'is-below' : ''}`} style={{ left: `${Math.min(88, Math.max(12, (active.x / MATRIX_W) * 100))}%`, top: `${(active.y / MATRIX_H) * 100}%` }}>
      <strong>{active.tag.label}</strong>
      <span>样本数<b>{active.tag.sampleCount}</b></span>
      <span>内容供给率<b>{(active.tag.supplyRate * 100).toFixed(1)}%</b></span>
      <span>高表现率<b>{(active.tag.highRate * 100).toFixed(1)}%</b></span>
      <span>高表现笔记<b>{active.tag.highCount} 篇</b></span>
      {active.isCandidate && <em>潜在内容机会</em>}
    </div>}
  </div>
}

function OpportunityCard({ candidate, baselineHighRate }) {
  return <article className="opportunity-card">
    <div className="candidate-title"><h4>{candidate.label}</h4><span>潜在内容机会</span></div>
    <div className="candidate-metrics"><div><small>高表现率</small><strong>{(candidate.highRate * 100).toFixed(1)}%</strong></div><div><small>整体基准</small><strong>{(baselineHighRate * 100).toFixed(1)}%</strong></div><div><small>样本数</small><strong>{candidate.sampleCount}</strong></div><div><small>内容供给率</small><strong>{(candidate.supplyRate * 100).toFixed(1)}%</strong></div><div><small>相对基准</small><strong>{candidate.performanceLift >= 0 ? '+' : ''}{(candidate.performanceLift * 100).toFixed(1)}pct</strong></div><div><small>机会指数</small><strong>{candidate.opportunityScore.toFixed(3)}</strong></div></div>
    <div className="representative-notes"><small>代表内容</small>{(candidate.representativeNotes || []).map((row, index) => <div key={`${row['笔记标题']}-${index}`}><span>{row['笔记链接'] && row['笔记链接'] !== '—' ? <a href={row['笔记链接']} target="_blank" rel="noreferrer">{row['笔记标题']}</a> : row['笔记标题']}</span><em>{formatNumber(row['点赞数'])} 赞 · {formatNumber(row['收藏数'])} 藏</em></div>)}</div>
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

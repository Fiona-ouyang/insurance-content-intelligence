import http from 'node:http'
import dotenv from 'dotenv'

dotenv.config()

const port = Number(process.env.API_PORT) || 3001
const OPENROUTER_MODEL = 'deepseek/deepseek-v3.2'
const OPENROUTER_URL = 'https://openrouter.ai/api/v1/chat/completions'
const ANALYSIS_FIELDS = ['内容主题', '目标人群', '用户需求', '内容类型', '内容场景', '标题钩子', '内容目的']
const MAX_ANALYSIS_ATTEMPTS = 2

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' })
  response.end(JSON.stringify(payload))
}

function readJson(request) {
  return new Promise((resolve, reject) => {
    let body = ''
    request.on('data', (chunk) => { body += chunk })
    request.on('end', () => {
      try {
        resolve(JSON.parse(body || '{}'))
      } catch {
        reject(new Error('Invalid JSON'))
      }
    })
    request.on('error', reject)
  })
}

function validateLabels(labels, tagTaxonomy) {
  if (!labels || typeof labels !== 'object' || Array.isArray(labels)) return '模型结果必须是对象。'
  for (const field of ANALYSIS_FIELDS) {
    if (!Array.isArray(labels[field])) return `${field} 必须是数组。`
    if (labels[field].length < 1 || labels[field].length > 2) return `${field} 必须包含 1-2 个标签。`
    if (labels[field].some((label) => typeof label !== 'string' || !tagTaxonomy[field]?.includes(label))) {
      return `${field} 包含词典外标签。`
    }
  }
  return ''
}

function extractJson(text) {
  const direct = text.trim()
  try {
    return JSON.parse(direct)
  } catch {}

  const fenced = direct.match(/```(?:json)?\s*([\s\S]*?)\s*```/i)
  if (fenced) {
    try {
      return JSON.parse(fenced[1])
    } catch {}
  }

  const start = direct.indexOf('{')
  if (start < 0) throw new Error('模型没有返回有效 JSON。')
  let depth = 0
  let inString = false
  let escaped = false
  for (let index = start; index < direct.length; index += 1) {
    const character = direct[index]
    if (inString) {
      if (escaped) escaped = false
      else if (character === '\\') escaped = true
      else if (character === '"') inString = false
      continue
    }
    if (character === '"') inString = true
    else if (character === '{') depth += 1
    else if (character === '}') {
      depth -= 1
      if (depth === 0) {
        try {
          return JSON.parse(direct.slice(start, index + 1))
        } catch {
          break
        }
      }
    }
  }
  throw new Error('模型没有返回有效 JSON。')
}

function getAnalysisSchema(tagTaxonomy) {
  return {
    type: 'object',
    properties: Object.fromEntries(ANALYSIS_FIELDS.map((field) => [field, {
      type: 'array',
      minItems: 1,
      maxItems: 2,
      items: { type: 'string', enum: tagTaxonomy[field] },
    }])),
    required: ANALYSIS_FIELDS,
    additionalProperties: false,
  }
}

async function analyzeWithModel(model, prompt, apiKey, tagTaxonomy) {
  const response = await fetch(OPENROUTER_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      model,
      messages: [{ role: 'user', content: prompt }],
      response_format: {
        type: 'json_schema',
        json_schema: {
          name: 'insurance_content_analysis',
          strict: true,
          schema: getAnalysisSchema(tagTaxonomy),
        },
      },
      provider: { require_parameters: true },
      temperature: 0,
      max_tokens: 500,
    }),
  })
  const responseText = await response.text()
  if (!response.ok) {
    const error = new Error(`OpenRouter HTTP ${response.status}: ${responseText.slice(0, 300)}`)
    error.statusCode = response.status
    throw error
  }

  let payload
  try {
    payload = JSON.parse(responseText)
  } catch {
    throw new Error('OpenRouter 返回了无效响应。')
  }
  const modelContent = payload.choices?.[0]?.message?.content
  if (typeof modelContent !== 'string') throw new Error('模型响应中没有找到文本内容。')

  const labels = extractJson(modelContent)
  const validationError = validateLabels(labels, tagTaxonomy)
  if (validationError) throw new Error(`模型结果校验失败：${validationError}`)
  return labels
}

async function analyzeWithOpenRouter(title, content, tagTaxonomy) {
  const apiKey = process.env.OPENROUTER_API_KEY
  if (!apiKey) throw new Error('服务端未配置 OPENROUTER_API_KEY。')

  const prompt = `你是一名保险新媒体内容分析师。根据笔记标题和正文，从给定 TAG_TAXONOMY 中为每个维度选择 1–2 个最匹配标签。禁止输出词典外标签。只返回 JSON，不解释，不输出 Markdown。

笔记标题：${title}
正文：${content}
TAG_TAXONOMY：${JSON.stringify(tagTaxonomy)}

必须返回包含以下全部字段的 JSON 对象，每个字段必须是包含 1–2 个字符串标签的数组：${ANALYSIS_FIELDS.join('、')}`

  const labels = await analyzeWithModel(OPENROUTER_MODEL, prompt, apiKey, tagTaxonomy)
  return { labels, model: OPENROUTER_MODEL }
}

const server = http.createServer(async (request, response) => {
  if (request.method !== 'POST' || request.url !== '/api/analyze') {
    sendJson(response, 404, { success: false, message: 'Not found' })
    return
  }

  try {
    const { title, content, TAG_TAXONOMY } = await readJson(request)
    if (typeof title !== 'string' || typeof content !== 'string' || !TAG_TAXONOMY || typeof TAG_TAXONOMY !== 'object') {
      sendJson(response, 400, { success: false, message: 'Invalid analyze payload' })
      return
    }
    let lastError
    for (let attempt = 1; attempt <= MAX_ANALYSIS_ATTEMPTS; attempt += 1) {
      try {
        const result = await analyzeWithOpenRouter(title, content, TAG_TAXONOMY)
        sendJson(response, 200, { success: true, labels: result.labels, model: result.model, attempt })
        return
      } catch (error) {
        lastError = error
        if (attempt < MAX_ANALYSIS_ATTEMPTS) continue
      }
    }
    throw lastError
  } catch (error) {
    const message = error instanceof Error ? error.message : 'AI 分析失败。'
    const statusCode = message.includes('未配置') ? 503 : 502
    sendJson(response, statusCode, { success: false, message })
  }
})

server.listen(port, () => {
  console.log(`Local AI API listening at http://localhost:${port}`)
})
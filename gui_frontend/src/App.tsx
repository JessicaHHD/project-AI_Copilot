import { useState, useEffect, useCallback, useRef } from 'react'

// ─── Data ───────────────────────────────────────────────────────────────────

type StageStatus = 'completed' | 'need_user' | 'not_started' | 'error'

interface Stage {
  id: string
  label: string
  status: StageStatus
  description: string
}

interface FileEntry {
  id: string
  label: string
  status: 'found' | 'available' | 'missing' | 'empty'
  path: string
  display_path?: string
  action: 'open' | 'copy'
}

interface DashboardContext {
  batch_name: string
  environment: string
  batch_date: string
  output_dir: string
  log_dir: string
  current_stage: string
  next_step: string
  manifest_status: string
  manifest_path: string
}

interface PoolOverview {
  available: boolean
  message: string
  source_file: string
  last_submitted_count: number | null
  success_count: number | null
  failed_count: number | null
  resubmittable_count: number | null
  exit_count: number | null
  failure_reasons: Array<{ reason: string; count: number }>
}

interface LogStatus {
  latest_log: { path: string; display_path: string; message: string; updated_at: string }
  current_exception: string
  debug_dir: string
  log_dir: string
  suspicious_dir: string
}

interface DashboardData {
  context: DashboardContext
  pool_overview: PoolOverview
  stages: Stage[]
  key_files: FileEntry[]
  log_status: LogStatus
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function statusColor(s: StageStatus) {
  return {
    completed: '#22c55e',
    need_user: '#f59e0b',
    not_started: '#9ca3af',
    error: '#ef4444',
  }[s]
}

function fileStatusBadge(s: FileEntry['status']) {
  const map: Record<FileEntry['status'], { label: string; bg: string; text: string }> = {
    found: { label: '已找到', bg: '#dcfce7', text: '#15803d' },
    available: { label: '可操作', bg: '#dbeafe', text: '#1d4ed8' },
    missing: { label: '暂未生成', bg: '#f3f4f6', text: '#6b7280' },
    empty: { label: '暂无', bg: '#f3f4f6', text: '#6b7280' },
  }
  return map[s]
}

async function apiGetDashboard(): Promise<DashboardData> {
  const response = await fetch('/api/dashboard')
  if (!response.ok) throw new Error(`看板接口返回 ${response.status}`)
  return response.json()
}

async function apiPost<T>(url: string, payload: Record<string, unknown> = {}): Promise<T> {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  const data = await response.json()
  if (!response.ok || data.ok === false) throw new Error(data.message || `接口返回 ${response.status}`)
  return data
}

function formatCount(value: number | null) {
  return typeof value === 'number' ? value.toLocaleString() : '—'
}

function shortPath(file: FileEntry) {
  const value = file.display_path || file.path || '暂未生成'
  return value.length > 38 ? '…' + value.slice(-35) : value
}

// ─── Toast ──────────────────────────────────────────────────────────────────

interface ToastState { msg: string; key: number }

function Toast({ toast }: { toast: ToastState | null }) {
  if (!toast) return null
  return (
    <div
      key={toast.key}
      style={{
        position: 'fixed',
        bottom: 28,
        right: 28,
        zIndex: 9999,
        background: '#1e293b',
        color: '#f1f5f9',
        padding: '10px 18px',
        borderRadius: 8,
        fontSize: 13,
        fontFamily: 'var(--font-sans)',
        boxShadow: '0 4px 16px rgba(0,0,0,0.18)',
        animation: 'fadeSlideIn 0.2s ease',
      }}
    >
      {toast.msg}
    </div>
  )
}

// ─── Failure Modal ───────────────────────────────────────────────────────────

function FailureModal({ onClose, onOpen }: { onClose: () => void; onOpen: (msg: string) => void }) {
  return (
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 1000,
        background: 'rgba(15,23,42,0.45)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}
      onClick={onClose}
    >
      <div
        style={{
          background: '#fff', borderRadius: 10, padding: '28px 32px',
          width: 420, boxShadow: '0 8px 32px rgba(0,0,0,0.16)',
        }}
        onClick={e => e.stopPropagation()}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
          <span style={{ fontSize: 15, fontWeight: 600, color: '#1e293b' }}>失败原因 Top 3</span>
          <button onClick={onClose} style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: 18, color: '#9ca3af', lineHeight: 1 }}>×</button>
        </div>
        <p style={{ fontSize: 13, color: '#64748b', lineHeight: 1.7, margin: 0 }}>
          当前工作台只展示失败原因汇总，不在页面展示完整 SKU 明细。需要核对时，请通过关键文件区打开整理表。
        </p>
        <div style={{ marginTop: 20 }}>
          <button
            onClick={() => { onOpen('已打开提报情况整理表'); onClose() }}
            style={{
              width: '100%', padding: '9px 0', background: '#1a6cf6', color: '#fff',
              border: 'none', borderRadius: 6, fontSize: 13, fontWeight: 500, cursor: 'pointer',
            }}
          >
            我知道了
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Stage workspace panels ──────────────────────────────────────────────────

// ─── Upload Modal ────────────────────────────────────────────────────────────

interface UploadedFile { name: string; size: string }

function UploadModal({ onClose, onDone }: { onClose: () => void; onDone: (f: UploadedFile) => void }) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)
  const [file, setFile] = useState<UploadedFile | null>(null)

  const handleFile = (f: File) => {
    const kb = (f.size / 1024).toFixed(1)
    setFile({ name: f.name, size: kb + ' KB' })
  }

  return (
    <div
      style={{ position: 'fixed', inset: 0, zIndex: 1000, background: 'rgba(15,23,42,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
      onClick={onClose}
    >
      <div
        style={{ background: '#fff', borderRadius: 10, padding: '28px 32px', width: 480, boxShadow: '0 8px 32px rgba(0,0,0,0.16)' }}
        onClick={e => e.stopPropagation()}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20 }}>
          <span style={{ fontSize: 15, fontWeight: 600, color: '#1e293b' }}>上传文件</span>
          <button onClick={onClose} style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: 18, color: '#9ca3af' }}>×</button>
        </div>

        {/* Drop zone */}
        <div
          onDragOver={e => { e.preventDefault(); setDragging(true) }}
          onDragLeave={() => setDragging(false)}
          onDrop={e => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) handleFile(f) }}
          onClick={() => inputRef.current?.click()}
          style={{
            border: `2px dashed ${dragging ? '#1a6cf6' : '#cbd5e1'}`,
            borderRadius: 8,
            padding: '32px 24px',
            textAlign: 'center',
            cursor: 'pointer',
            background: dragging ? '#eff6ff' : '#f8fafc',
            transition: 'border-color 0.15s, background 0.15s',
            marginBottom: 16,
          }}
        >
          <div style={{ fontSize: 28, marginBottom: 8, color: dragging ? '#1a6cf6' : '#94a3b8' }}>📂</div>
          <div style={{ fontSize: 13, color: '#475569', fontWeight: 500 }}>拖拽文件到此处，或点击选择文件</div>
          <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 4 }}>支持 .xlsx · .csv · .json · .txt</div>
          <input ref={inputRef} type="file" style={{ display: 'none' }} onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f) }} />
        </div>

        {/* File list */}
        <div style={{ minHeight: 48, marginBottom: 20 }}>
          {file ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 7 }}>
              <span style={{ fontSize: 18 }}>📄</span>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: '#1e293b' }}>{file.name}</div>
                <div style={{ fontSize: 11, color: '#64748b', fontFamily: 'var(--font-mono)' }}>{file.size}</div>
              </div>
              <button onClick={() => setFile(null)} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#9ca3af', fontSize: 16 }}>×</button>
            </div>
          ) : (
            <div style={{ fontSize: 12, color: '#cbd5e1', textAlign: 'center', paddingTop: 12 }}>暂无已选文件</div>
          )}
        </div>

        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
          <button onClick={onClose} style={{ padding: '8px 18px', border: '1px solid #d1d5db', borderRadius: 6, background: '#fff', color: '#475569', fontSize: 13, cursor: 'pointer', fontFamily: 'var(--font-sans)' }}>取消</button>
          <button
            disabled={!file}
            onClick={() => { if (file) { onDone(file); onClose() } }}
            style={{ padding: '8px 18px', border: 'none', borderRadius: 6, background: file ? '#1a6cf6' : '#e2e8f0', color: file ? '#fff' : '#9ca3af', fontSize: 13, cursor: file ? 'pointer' : 'not-allowed', fontFamily: 'var(--font-sans)', fontWeight: 500 }}
          >确认上传</button>
        </div>
      </div>
    </div>
  )
}

// ─── Batch Config Panel (slide-in drawer style) ───────────────────────────────

interface BatchConfig { month: string; batch: string; env: string; outputRoot: string }

function BatchConfigModal({ initial, onClose, onSave }: {
  initial: BatchConfig
  onClose: () => void
  onSave: (c: BatchConfig) => void
}) {
  const [form, setForm] = useState<BatchConfig>(initial)
  const set = (k: keyof BatchConfig) => (v: string) => setForm(f => ({ ...f, [k]: v }))

  const months = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']
  const batches = ['第1批提报', '第2批提报', '第3批提报']
  const envs = ['测试环境', '生产环境']

  return (
    <div
      style={{ position: 'fixed', inset: 0, zIndex: 1000, background: 'rgba(15,23,42,0.45)', display: 'flex', alignItems: 'flex-start', justifyContent: 'flex-end' }}
      onClick={onClose}
    >
      <div
        style={{
          background: '#fff', width: 400, height: '100%', padding: '28px 28px',
          boxShadow: '-4px 0 32px rgba(0,0,0,0.12)',
          display: 'flex', flexDirection: 'column', gap: 0,
          animation: 'slideInRight 0.22s ease',
        }}
        onClick={e => e.stopPropagation()}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24 }}>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, color: '#0f172a' }}>批次配置</div>
            <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 2 }}>选择提报月份、批次和运行环境</div>
          </div>
          <button onClick={onClose} style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: 20, color: '#9ca3af' }}>×</button>
        </div>

        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 20 }}>
          <ConfigField label="提报月份">
            <SegmentControl options={months} value={form.month} onChange={set('month')} cols={4} />
          </ConfigField>

          <ConfigField label="提报批次">
            <SegmentControl options={batches} value={form.batch} onChange={set('batch')} cols={3} />
          </ConfigField>

          <ConfigField label="运行环境">
            <SegmentControl options={envs} value={form.env} onChange={set('env')} cols={2} />
          </ConfigField>

          <ConfigField label="输出根目录">
            <div style={{ display: 'flex', gap: 6 }}>
              <input
                value={form.outputRoot}
                onChange={e => set('outputRoot')(e.target.value)}
                style={{
                  flex: 1, padding: '7px 10px', border: '1px solid #d1d5db', borderRadius: 6,
                  fontSize: 13, fontFamily: 'var(--font-mono)', color: '#334155', outline: 'none',
                }}
              />
              <button style={{ padding: '7px 12px', border: '1px solid #d1d5db', borderRadius: 6, background: '#f8fafc', fontSize: 12, cursor: 'pointer', color: '#475569', fontFamily: 'var(--font-sans)' }}>
                浏览
              </button>
            </div>
          </ConfigField>

          <div style={{ padding: '14px 14px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 10 }}>配置预览</div>
            <div style={{ fontSize: 13, color: '#475569', lineHeight: 2, fontFamily: 'var(--font-mono)' }}>
              <span style={{ color: '#64748b' }}>批次：</span>{form.month}{form.batch}<br />
              <span style={{ color: '#64748b' }}>环境：</span>{form.env}<br />
              <span style={{ color: '#64748b' }}>输出：</span>{form.outputRoot}
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', gap: 8, paddingTop: 20, borderTop: '1px solid #f1f5f9' }}>
          <button onClick={onClose} style={{ flex: 1, padding: '9px 0', border: '1px solid #d1d5db', borderRadius: 6, background: '#fff', color: '#475569', fontSize: 13, cursor: 'pointer', fontFamily: 'var(--font-sans)' }}>取消</button>
          <button
            onClick={() => { onSave(form); onClose() }}
            style={{ flex: 2, padding: '9px 0', border: 'none', borderRadius: 6, background: '#1a6cf6', color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', fontFamily: 'var(--font-sans)' }}
          >保存配置</button>
        </div>
      </div>
    </div>
  )
}

function ConfigField({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <div style={{ fontSize: 12, fontWeight: 600, color: '#64748b', marginBottom: 8 }}>{label}</div>
      {children}
    </div>
  )
}

function SegmentControl({ options, value, onChange, cols }: { options: string[]; value: string; onChange: (v: string) => void; cols: number }) {
  return (
    <div style={{ display: 'grid', gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 6 }}>
      {options.map(opt => (
        <button
          key={opt}
          onClick={() => onChange(opt)}
          style={{
            padding: '7px 4px', fontSize: 12, borderRadius: 6, cursor: 'pointer',
            border: `1px solid ${value === opt ? '#1a6cf6' : '#e2e8f0'}`,
            background: value === opt ? '#eff6ff' : '#fff',
            color: value === opt ? '#1d4ed8' : '#475569',
            fontWeight: value === opt ? 600 : 400,
            fontFamily: 'var(--font-sans)',
            transition: 'border-color 0.12s, background 0.12s',
          }}
        >{opt}</button>
      ))}
    </div>
  )
}

// ─── Prepare Panel ────────────────────────────────────────────────────────────

function PreparePanel({ showToast }: { showToast: (m: string) => void }) {
  const [uploadOpen, setUploadOpen] = useState(false)
  const [configOpen, setConfigOpen] = useState(false)
  const [uploadedFile, setUploadedFile] = useState<UploadedFile | null>(null)
  const [batchConfig, setBatchConfig] = useState<BatchConfig>({ month: '8月', batch: '第1批提报', env: '测试环境', outputRoot: 'data/output' })

  return (
    <div>
      <StageBadge label="已完成" color="#22c55e" bg="#dcfce7" />
      <p style={descStyle}>批次清单已识别，配置加载完成。可在此调整提报批次和运行环境，或上传所需的输入文件后进入下一阶段。</p>

      {/* Current config summary */}
      <SectionTitle>当前批次配置</SectionTitle>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginBottom: 4 }}>
        {[
          { label: '提报批次', value: batchConfig.month + batchConfig.batch },
          { label: '运行环境', value: batchConfig.env },
          { label: '输出目录', value: batchConfig.outputRoot },
          { label: '批次清单', value: 'manifest.json · 已找到' },
        ].map(({ label, value }) => (
          <div key={label} style={{ padding: '9px 12px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 7 }}>
            <div style={{ fontSize: 11, color: '#94a3b8', marginBottom: 3 }}>{label}</div>
            <div style={{ fontSize: 13, fontFamily: 'var(--font-mono)', color: '#1e293b' }}>{value}</div>
          </div>
        ))}
      </div>

      {/* Uploaded file */}
      <SectionTitle style={{ marginTop: 20 }}>输入文件</SectionTitle>
      {uploadedFile ? (
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 12px', background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 7, marginBottom: 4 }}>
          <span style={{ fontSize: 16 }}>📄</span>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 13, fontWeight: 500, color: '#1e293b' }}>{uploadedFile.name}</div>
            <div style={{ fontSize: 11, color: '#64748b', fontFamily: 'var(--font-mono)' }}>{uploadedFile.size}</div>
          </div>
          <span style={{ fontSize: 11, background: '#dcfce7', color: '#15803d', padding: '2px 8px', borderRadius: 20, fontWeight: 600 }}>已上传</span>
          <button onClick={() => setUploadedFile(null)} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#9ca3af', fontSize: 16 }}>×</button>
        </div>
      ) : (
        <div style={{ fontSize: 13, color: '#94a3b8', padding: '10px 0', fontStyle: 'italic' }}>尚未上传文件，点击下方按钮上传</div>
      )}

      <SectionTitle style={{ marginTop: 20 }}>确认项</SectionTitle>
      <CheckRow label="批次清单" value="已找到 · manifest.json" ok />
      <CheckRow label="输出目录" value={batchConfig.outputRoot} ok />
      <CheckRow label="当前环境" value={batchConfig.env} ok />

      <ActionRow>
        <ActionBtn primary onClick={() => setConfigOpen(true)}>配置提报批次</ActionBtn>
        <ActionBtn onClick={() => setUploadOpen(true)}>上传输入文件</ActionBtn>
        <ActionBtn onClick={() => showToast('已打开输出目录')}>打开输出目录</ActionBtn>
      </ActionRow>

      {uploadOpen && (
        <UploadModal
          onClose={() => setUploadOpen(false)}
          onDone={f => { setUploadedFile(f); showToast(`已上传 ${f.name}`) }}
        />
      )}
      {configOpen && (
        <BatchConfigModal
          initial={batchConfig}
          onClose={() => setConfigOpen(false)}
          onSave={c => { setBatchConfig(c); showToast(`批次已更新：${c.month}${c.batch}`) }}
        />
      )}
    </div>
  )
}

function PricingPanel({ showToast }: { showToast: (m: string) => void }) {
  return (
    <div>
      <StageBadge label="已完成" color="#22c55e" bg="#dcfce7" />
      <p style={descStyle}>筛品查价已完成，新人价整合表已生成。本批次共查价 <strong>12,340</strong> 个 SKU，其中查价失败 <strong>1,240</strong> 个，已写入失败 SKU 表。</p>
      <SectionTitle>输出文件</SectionTitle>
      <CheckRow label="新人价整合表" value="data/output/最终结果/…_筛品查价表.xlsx" ok />
      <CheckRow label="查价失败 SKU" value="data/output/最终结果/查价失败SKU.xlsx" ok />
      <SectionTitle>下一步</SectionTitle>
      <NextStep text="进入「业务确认」阶段，复制 ERP 文本并发送确认邮件。" />
      <ActionRow>
        <ActionBtn onClick={() => showToast('已打开新人价整合表')}>打开整合表</ActionBtn>
        <ActionBtn onClick={() => showToast('已打开查价失败 SKU 表')}>查看失败 SKU</ActionBtn>
      </ActionRow>
    </div>
  )
}

function ConfirmPanel({ showToast }: { showToast: (m: string) => void }) {
  return (
    <div>
      <StageBadge label="需人工处理" color="#d97706" bg="#fef3c7" />
      <p style={descStyle}>业务需要确认哪些 SKU 参与报名。普通复提无异常时可跳过；<strong>首次提报或价格异常时必须确认</strong>。</p>
      <SectionTitle>本阶段输入材料</SectionTitle>
      <CheckRow label="新人价整合表" value="已生成 · data/output/最终结果/…" ok />
      <CheckRow label="ERP 合并文本" value="可复制" ok />
      <SectionTitle>操作</SectionTitle>
      <ActionRow>
        <ActionBtn primary onClick={() => showToast('已复制 ERP 合并文本')}>复制 ERP 合并文本</ActionBtn>
        <ActionBtn onClick={() => showToast('已打开新人价整合表')}>打开新人价整合表</ActionBtn>
        <ActionBtn disabled>上传业务确认表（待收表）</ActionBtn>
      </ActionRow>
      <NextStep text="收到业务确认表后，工具将识别参与报名、不报名和未确认 SKU，并自动生成提报文件。" />
    </div>
  )
}

function SubmitPanel() {
  return (
    <div>
      <StageBadge label="未开始" color="#9ca3af" bg="#f3f4f6" />
      <p style={descStyle}>等待业务确认阶段完成后，将自动生成提报 PART 文件，并通过后台脚本执行批量提报。</p>
      <div style={{
        marginTop: 16, padding: '14px 16px', background: '#fff7ed',
        border: '1px solid #fed7aa', borderRadius: 8, fontSize: 13, color: '#92400e',
        display: 'flex', gap: 10, alignItems: 'flex-start',
      }}>
        <span style={{ fontSize: 16 }}>⚠️</span>
        <span><strong>高风险操作提示：</strong>真实提报为高风险动作，当前 MVP 仅展示状态，不在 GUI 中执行。请通过命令行脚本完成后台提报。</span>
      </div>
      <SectionTitle style={{ marginTop: 20 }}>前置条件</SectionTitle>
      <CheckRow label="业务确认表" value="待上传" warn />
      <CheckRow label="提报 PART 文件" value="暂未生成" warn />
    </div>
  )
}

function MergePanel() {
  return (
    <div>
      <StageBadge label="未开始" color="#9ca3af" bg="#f3f4f6" />
      <p style={descStyle}>提报完成后，脚本将自动下载审核结果，识别通过、拒绝和未审核 SKU，并生成提报情况整理表。</p>
      <SectionTitle>前置条件</SectionTitle>
      <CheckRow label="后台提报" value="未完成" warn />
      <CheckRow label="提报情况整理表" value="暂未生成" warn />
    </div>
  )
}

// ─── Small reused components ─────────────────────────────────────────────────

const descStyle: React.CSSProperties = { fontSize: 13, color: '#475569', lineHeight: 1.7, marginBottom: 16 }
const codeStyle: React.CSSProperties = { fontFamily: 'var(--font-mono)', fontSize: 12, background: '#f1f5f9', padding: '1px 5px', borderRadius: 3 }

function StageBadge({ label, color, bg }: { label: string; color: string; bg: string }) {
  return (
    <span style={{
      display: 'inline-block', padding: '3px 10px', borderRadius: 20,
      fontSize: 12, fontWeight: 600, color, background: bg, marginBottom: 12,
    }}>{label}</span>
  )
}

function SectionTitle({ children, style }: { children: React.ReactNode; style?: React.CSSProperties }) {
  return <div style={{ fontSize: 12, fontWeight: 600, color: '#94a3b8', letterSpacing: '0.06em', textTransform: 'uppercase', marginBottom: 8, marginTop: 20, ...style }}>{children}</div>
}

function CheckRow({ label, value, ok, warn }: { label: string; value: string; ok?: boolean; warn?: boolean }) {
  const icon = ok ? '✓' : warn ? '–' : '✕'
  const color = ok ? '#22c55e' : warn ? '#9ca3af' : '#ef4444'
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '7px 0', borderBottom: '1px solid #f1f5f9', fontSize: 13 }}>
      <span style={{ color, fontWeight: 700, width: 16, textAlign: 'center', fontSize: 14 }}>{icon}</span>
      <span style={{ color: '#334155', flex: '0 0 120px' }}>{label}</span>
      <span style={{ color: '#64748b', fontFamily: 'var(--font-mono)', fontSize: 12 }}>{value}</span>
    </div>
  )
}

function NextStep({ text }: { text: string }) {
  return (
    <div style={{
      marginTop: 16, padding: '11px 14px', background: '#eff6ff',
      border: '1px solid #bfdbfe', borderRadius: 7, fontSize: 13, color: '#1e40af',
      display: 'flex', gap: 8,
    }}>
      <span>→</span><span>{text}</span>
    </div>
  )
}

function ActionRow({ children }: { children: React.ReactNode }) {
  return <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 18 }}>{children}</div>
}

function ActionBtn({ children, onClick, primary, disabled }: {
  children: React.ReactNode; onClick?: () => void; primary?: boolean; disabled?: boolean
}) {
  return (
    <button
      onClick={disabled ? undefined : onClick}
      style={{
        padding: '7px 14px',
        background: disabled ? '#f1f5f9' : primary ? '#1a6cf6' : '#fff',
        color: disabled ? '#9ca3af' : primary ? '#fff' : '#334155',
        border: `1px solid ${disabled ? '#e2e8f0' : primary ? '#1a6cf6' : '#d1d5db'}`,
        borderRadius: 6, fontSize: 13, fontWeight: 500, cursor: disabled ? 'not-allowed' : 'pointer',
        fontFamily: 'var(--font-sans)',
        transition: 'background 0.15s, border-color 0.15s',
      }}
    >{children}</button>
  )
}

// ─── Main App ────────────────────────────────────────────────────────────────

export default function App() {
  const [activeStage, setActiveStage] = useState('confirm')
  const [dashboard, setDashboard] = useState<DashboardData | null>(null)
  const [loading, setLoading] = useState(true)
  const [apiError, setApiError] = useState('')
  const [toast, setToast] = useState<ToastState | null>(null)
  const [toastKey, setToastKey] = useState(0)
  const [modalOpen, setModalOpen] = useState(false)
  const [refreshing, setRefreshing] = useState(false)

  const showToast = useCallback((msg: string) => {
    const key = toastKey + 1
    setToastKey(key)
    setToast({ msg, key })
  }, [toastKey])

  useEffect(() => {
    if (!toast) return
    const t = setTimeout(() => setToast(null), 2500)
    return () => clearTimeout(t)
  }, [toast])

  const fetchDashboard = useCallback(async () => {
    setApiError('')
    try {
      const data = await apiGetDashboard()
      setDashboard(data)
    } catch (err) {
      setApiError(err instanceof Error ? err.message : '看板接口请求失败')
    } finally {
      setLoading(false)
      setRefreshing(false)
    }
  }, [])

  useEffect(() => { void fetchDashboard() }, [fetchDashboard])

  const stages = dashboard?.stages ?? []
  const keyFiles = dashboard?.key_files ?? []
  const context = dashboard?.context
  const pool = dashboard?.pool_overview
  const logStatus = dashboard?.log_status
  const currentStage = stages.find(s => s.id === activeStage) ?? stages[0] ?? { id: 'prepare', label: '准备任务', status: 'not_started' as StageStatus, description: '正在读取看板数据。' }

  useEffect(() => {
    if (!stages.length || stages.some(stage => stage.id === activeStage)) return
    setActiveStage(stages.find(stage => stage.status === 'need_user')?.id ?? stages[0].id)
  }, [activeStage, stages])

  const handleRefresh = () => { setRefreshing(true); void fetchDashboard().then(() => showToast('状态已刷新')) }

  const openPath = async (path: string, label = '路径') => {
    if (!path) { showToast(`${label}暂未生成`); return }
    try {
      await apiPost('/api/open-path', { path })
      showToast(`已打开 ${label}`)
    } catch (err) {
      showToast(err instanceof Error ? err.message : `${label}打开失败`)
    }
  }

  const startCli = async () => {
    try {
      await apiPost('/api/start-cli')
      showToast('已启动命令行入口')
    } catch (err) {
      showToast(err instanceof Error ? err.message : '命令行入口启动失败')
    }
  }

  const handleFileAction = async (file: FileEntry) => {
    if (file.action === 'copy') {
      if (!file.path) { showToast('ERP 文本暂未生成'); return }
      await navigator.clipboard.writeText(file.path)
      showToast(`已复制 ${file.label}`)
      return
    }
    await openPath(file.path, file.label)
  }

  return (
    <div style={{ minHeight: '100vh', background: '#f0f4f8', fontFamily: 'var(--font-sans)', fontSize: 14 }}>
      <style>{`
        @keyframes fadeSlideIn { from { opacity:0; transform:translateY(8px) } to { opacity:1; transform:translateY(0) } }
        @keyframes slideInRight { from { transform:translateX(40px); opacity:0 } to { transform:translateX(0); opacity:1 } }
        button:hover:not(:disabled) { filter: brightness(0.96); }
      `}</style>

      {/* ── Top Bar ── */}
      <div style={{
        background: '#fff', borderBottom: '1px solid #e2e8f0',
        padding: '0 28px', display: 'flex', alignItems: 'center', gap: 20, height: 56,
        position: 'sticky', top: 0, zIndex: 100,
      }}>
        <span style={{ fontSize: 16, fontWeight: 700, color: '#0f172a', whiteSpace: 'nowrap' }}>新人价自动化工作台</span>
        <div style={{ width: 1, height: 20, background: '#e2e8f0' }} />
        <TopBadge label={context?.batch_name || '读取批次中'} color="#1e40af" bg="#dbeafe" />
        <TopBadge label={context?.environment || '本地环境'} color="#92400e" bg="#fef3c7" />
        <div style={{ height: 20, width: 1, background: '#e2e8f0' }} />
        <span style={{ fontSize: 13, color: '#475569' }}>
          <span style={{ color: '#64748b' }}>当前阶段：</span>{context?.current_stage || currentStage.label}
        </span>
        <span style={{ fontSize: 13, color: '#1a6cf6', display: 'flex', alignItems: 'center', gap: 5 }}>
          <span>→</span>{context?.next_step || '正在读取下一步建议'}
        </span>
        <div style={{ flex: 1 }} />
        <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
          <span style={{ fontSize: 12, color: '#22c55e', display: 'flex', alignItems: 'center', gap: 4 }}>
            <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#22c55e', display: 'inline-block' }} />
            批次清单{context?.manifest_status || '读取中'}
          </span>
          <TopBtn onClick={handleRefresh} disabled={refreshing}>{refreshing ? '刷新中…' : '刷新状态'}</TopBtn>
          <TopBtn onClick={() => void openPath(context?.output_dir || '', '输出目录')}>打开输出目录</TopBtn>
          <TopBtn onClick={() => void openPath(context?.log_dir || '', '日志目录')}>打开日志目录</TopBtn>
          <TopBtn onClick={() => void startCli()}>启动命令行入口</TopBtn>
        </div>
      </div>

      {(loading || apiError) && (
        <div style={{ padding: '10px 28px', background: apiError ? '#fef2f2' : '#eff6ff', color: apiError ? '#b91c1c' : '#1d4ed8', borderBottom: '1px solid #e2e8f0', fontSize: 13 }}>
          {apiError ? `API 连接失败：${apiError}` : '正在读取批次看板数据…'}
        </div>
      )}

      {/* ── Pool Overview ── */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e2e8f0', padding: '16px 28px' }}>
        <div style={{ display: 'flex', gap: 0, alignItems: 'stretch' }}>
          {/* KPI tiles */}
          <div style={{ display: 'flex', gap: 0, flex: 1, borderRight: '1px solid #e2e8f0', paddingRight: 24, marginRight: 24 }}>
            {[
              { label: '上轮提报 SKU', value: formatCount(pool?.last_submitted_count ?? null), color: '#1e293b' },
              { label: '提报成功', value: formatCount(pool?.success_count ?? null), color: '#22c55e' },
              { label: '提报失败', value: formatCount(pool?.failed_count ?? null), color: '#ef4444' },
              { label: '可复提', value: formatCount(pool?.resubmittable_count ?? null), color: '#1a6cf6' },
              { label: '不报名 / 退出', value: formatCount(pool?.exit_count ?? null), color: '#9ca3af' },
            ].map((kpi, i) => (
              <div key={i} style={{
                flex: 1, padding: '10px 16px', borderRight: i < 4 ? '1px solid #f1f5f9' : 'none',
              }}>
                <div style={{ fontSize: 12, color: '#94a3b8', marginBottom: 5 }}>{kpi.label}</div>
                <div style={{ fontSize: 22, fontWeight: 700, color: kpi.color, fontFamily: 'var(--font-mono)', letterSpacing: '-0.02em' }}>{kpi.value}</div>
              </div>
            ))}
          </div>
          {/* Failure top 3 */}
          <div style={{ width: 280, paddingRight: 24, borderRight: '1px solid #e2e8f0', marginRight: 24 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 8 }}>失败原因 Top 3</div>
            {(pool?.failure_reasons.length ? pool.failure_reasons : [{ reason: pool?.message || '暂无可用数据', count: 0 }]).map((r, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '4px 0', fontSize: 13 }}>
                <span style={{ color: '#475569' }}>{r.reason}</span>
                <span style={{ fontFamily: 'var(--font-mono)', color: r.count ? '#ef4444' : '#94a3b8', fontWeight: 500 }}>{r.count ? r.count.toLocaleString() : '—'}</span>
              </div>
            ))}
          </div>
          {/* Actions */}
          <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: 6 }}>
            <PoolBtn onClick={() => void openPath(keyFiles.find(file => file.id === 'status_merge')?.path || '', '提报情况整理表')}>打开整理表</PoolBtn>
            <PoolBtn onClick={() => void openPath(context?.output_dir || '', '输出目录')}>打开输出目录</PoolBtn>
            <PoolBtn onClick={() => void startCli()}>启动命令行入口</PoolBtn>
          </div>
        </div>
      </div>

      {/* ── Three-column main area ── */}
      <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr 280px', gap: 0, flex: 1, padding: '16px 28px', minHeight: 'calc(100vh - 56px - 90px - 60px)' }}>

        {/* Left: Stage Nav */}
        <div style={{ paddingRight: 12 }}>
          <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
            <div style={{ padding: '12px 14px', borderBottom: '1px solid #f1f5f9', fontSize: 11, fontWeight: 600, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.06em' }}>流程阶段</div>
            {stages.map((stage, i) => {
              const isActive = stage.id === activeStage
              const color = statusColor(stage.status)
              return (
                <button
                  key={stage.id}
                  onClick={() => setActiveStage(stage.id)}
                  style={{
                    width: '100%', textAlign: 'left', border: 'none',
                    background: isActive ? '#eff6ff' : 'transparent',
                    borderLeft: isActive ? '3px solid #1a6cf6' : '3px solid transparent',
                    borderBottom: i < stages.length - 1 ? '1px solid #f8fafc' : 'none',
                    padding: '12px 14px 12px 12px',
                    cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 10,
                    transition: 'background 0.15s',
                    fontFamily: 'var(--font-sans)',
                  }}
                >
                  <StageIcon status={stage.status} />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: isActive ? 600 : 400, color: isActive ? '#1e40af' : '#334155' }}>{stage.label}</div>
                    <div style={{ fontSize: 11, color, marginTop: 2 }}>{stageStatusLabel(stage.status)}</div>
                  </div>
                  {isActive && <span style={{ color: '#1a6cf6', fontSize: 12 }}>›</span>}
                </button>
              )
            })}
          </div>
        </div>

        {/* Center: Workspace */}
        <div style={{ padding: '0 12px' }}>
          <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', padding: '20px 24px', minHeight: 360 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 18 }}>
              <StageIcon status={currentStage.status} size={20} />
              <span style={{ fontSize: 16, fontWeight: 700, color: '#0f172a' }}>{currentStage.label}</span>
            </div>
            <StageDetail
              stage={currentStage}
              context={context}
              erpFile={keyFiles.find(file => file.id === 'erp_text')}
              onOpen={openPath}
              onCopy={handleFileAction}
              onStartCli={startCli}
            />
          </div>
        </div>

        {/* Right: Key Files */}
        <div style={{ paddingLeft: 12 }}>
          <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
            <div style={{ padding: '12px 14px', borderBottom: '1px solid #f1f5f9', fontSize: 11, fontWeight: 600, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.06em' }}>关键文件</div>
            {keyFiles.map((f, i) => {
              const badge = fileStatusBadge(f.status)
              const canAct = f.status === 'found' || f.status === 'available'
              return (
                <div key={f.id} style={{
                  padding: '11px 14px', borderBottom: i < keyFiles.length - 1 ? '1px solid #f8fafc' : 'none',
                  display: 'flex', flexDirection: 'column', gap: 4,
                }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: 13, color: '#334155', fontWeight: 500 }}>{f.label}</span>
                    <span style={{
                      fontSize: 11, fontWeight: 600, padding: '2px 7px', borderRadius: 20,
                      background: badge.bg, color: badge.text,
                    }}>{badge.label}</span>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
                    <span style={{
                      fontFamily: 'var(--font-mono)', fontSize: 11, color: '#94a3b8',
                      overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1,
                    }}>{shortPath(f)}</span>
                    {canAct && (
                      <button
                        onClick={() => void handleFileAction(f)}
                        style={{
                          fontSize: 11, padding: '3px 9px', border: '1px solid #d1d5db',
                          borderRadius: 5, background: '#f8fafc', color: '#475569',
                          cursor: 'pointer', whiteSpace: 'nowrap', fontFamily: 'var(--font-sans)',
                        }}
                      >{f.action === 'copy' ? '复制' : '打开'}</button>
                    )}
                  </div>
                </div>
              )
            })}
          </div>
        </div>
      </div>

      {/* ── Footer: Log + Exception ── */}
      <div style={{
        background: '#1e293b', borderTop: '1px solid #334155',
        padding: '10px 28px', display: 'flex', alignItems: 'center', gap: 24,
        position: 'sticky', bottom: 0,
      }}>
        <div style={{ flex: 1, display: 'flex', gap: 24, alignItems: 'center', overflow: 'hidden' }}>
          <span style={{ fontSize: 11, fontFamily: 'var(--font-mono)', color: '#94a3b8', whiteSpace: 'nowrap' }}>
            <span style={{ color: '#64748b', marginRight: 6 }}>最近日志</span>
            {logStatus?.latest_log.updated_at ? `${logStatus.latest_log.updated_at} · ` : ''}{logStatus?.latest_log.message || '等待日志数据'}
          </span>
          <span style={{
            fontSize: 11, padding: '2px 8px', borderRadius: 20,
            background: '#14532d', color: '#86efac', fontWeight: 500,
          }}>{logStatus?.current_exception || '暂无异常'}</span>
        </div>
        <div style={{ display: 'flex', gap: 6 }}>
          <FooterBtn onClick={() => void openPath(logStatus?.log_dir || '', '日志目录')}>打开日志目录</FooterBtn>
          <FooterBtn onClick={() => void openPath(logStatus?.debug_dir || '', '调试目录')}>调试目录</FooterBtn>
          <FooterBtn onClick={() => void openPath(logStatus?.suspicious_dir || '', '疑似误下载目录')}>疑似误下载目录</FooterBtn>
        </div>
      </div>

      <Toast toast={toast} />
      {modalOpen && <FailureModal onClose={() => setModalOpen(false)} onOpen={showToast} />}
    </div>
  )
}

function StageDetail({ stage, context, erpFile, onOpen, onCopy, onStartCli }: {
  stage: Stage
  context?: DashboardContext
  erpFile?: FileEntry
  onOpen: (path: string, label?: string) => Promise<void>
  onCopy: (file: FileEntry) => Promise<void>
  onStartCli: () => Promise<void>
}) {
  return (
    <div>
      <p style={descStyle}>{stage.description}</p>
      <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '14px 16px', marginTop: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 8 }}>产品说明</div>
        <div style={{ fontSize: 13, color: '#64748b', lineHeight: 1.7 }}>
          当前版本以批次看板为核心，只读取配置、输出文件和日志；完整流程执行仍通过现有命令行工具完成。
        </div>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginTop: 16 }}>
        <CheckRow label={'批次'} value={context?.batch_name || '读取中'} ok={!!context?.batch_name} />
        <CheckRow label={'批次日期'} value={context?.batch_date || '读取中'} ok={!!context?.batch_date} />
        <CheckRow label={'输出目录'} value={context?.output_dir ? '已配置' : '读取中'} ok={!!context?.output_dir} />
        <CheckRow label={'日志目录'} value={context?.log_dir ? '已配置' : '读取中'} ok={!!context?.log_dir} />
      </div>
      <ActionRow>
        <ActionBtn primary onClick={() => void onOpen(context?.output_dir || '', '输出目录')}>打开输出目录</ActionBtn>
        <ActionBtn onClick={() => void onOpen(context?.log_dir || '', '日志目录')}>打开日志目录</ActionBtn>
        <ActionBtn onClick={() => erpFile && void onCopy(erpFile)} disabled={!erpFile || erpFile.status !== 'available'}>复制 ERP 文本</ActionBtn>
        <ActionBtn onClick={() => void onStartCli()}>启动命令行入口</ActionBtn>
      </ActionRow>
      <NextStep text={context?.next_step || '等待看板接口返回下一步建议。'} />
    </div>
  )
}

// ─── Tiny shared UI ──────────────────────────────────────────────────────────

function TopBadge({ label, color, bg }: { label: string; color: string; bg: string }) {
  return (
    <span style={{ fontSize: 12, fontWeight: 600, padding: '3px 10px', borderRadius: 20, background: bg, color }}>{label}</span>
  )
}

function TopBtn({ children, onClick, disabled }: { children: React.ReactNode; onClick?: () => void; disabled?: boolean }) {
  return (
    <button onClick={disabled ? undefined : onClick} style={{
      fontSize: 12, padding: '5px 12px', border: '1px solid #d1d5db',
      borderRadius: 6, background: '#f8fafc', color: '#334155',
      cursor: disabled ? 'not-allowed' : 'pointer', fontFamily: 'var(--font-sans)',
      opacity: disabled ? 0.6 : 1,
    }}>{children}</button>
  )
}

function PoolBtn({ children, onClick }: { children: React.ReactNode; onClick?: () => void }) {
  return (
    <button onClick={onClick} style={{
      fontSize: 12, padding: '6px 12px', border: '1px solid #d1d5db',
      borderRadius: 6, background: '#f8fafc', color: '#334155',
      cursor: 'pointer', whiteSpace: 'nowrap', fontFamily: 'var(--font-sans)',
    }}>{children}</button>
  )
}

function FooterBtn({ children, onClick }: { children: React.ReactNode; onClick?: () => void }) {
  return (
    <button onClick={onClick} style={{
      fontSize: 11, padding: '5px 11px', border: '1px solid #475569',
      borderRadius: 5, background: 'transparent', color: '#94a3b8',
      cursor: 'pointer', fontFamily: 'var(--font-sans)',
    }}>{children}</button>
  )
}

function StageIcon({ status, size = 16 }: { status: StageStatus; size?: number }) {
  const map: Record<StageStatus, { symbol: string; color: string; bg: string }> = {
    completed: { symbol: '✓', color: '#22c55e', bg: '#dcfce7' },
    need_user: { symbol: '!', color: '#d97706', bg: '#fef3c7' },
    not_started: { symbol: '·', color: '#9ca3af', bg: '#f3f4f6' },
    error: { symbol: '✕', color: '#ef4444', bg: '#fee2e2' },
  }
  const m = map[status]
  return (
    <span style={{
      width: size, height: size, borderRadius: '50%',
      background: m.bg, color: m.color,
      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
      fontSize: size * 0.65, fontWeight: 700, flexShrink: 0,
    }}>{m.symbol}</span>
  )
}

function stageStatusLabel(s: StageStatus) {
  return { completed: '已完成', need_user: '需人工', not_started: '未开始', error: '异常' }[s]
}

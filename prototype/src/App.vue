<script setup lang="ts">
import { computed, ref, type Component } from 'vue'
import {
  ArrowRight,
  ArrowUpRight,
  Bell,
  BookOpen,
  Bookmark,
  Bot,
  Check,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  CircleHelp,
  Clock3,
  CloudUpload,
  Cpu,
  Database,
  FileSearch,
  FileText,
  Files,
  FolderOpen,
  Grid2X2,
  LayoutDashboard,
  List,
  Menu,
  MessageSquareText,
  MoreHorizontal,
  PanelLeftClose,
  Plus,
  Quote,
  RefreshCw,
  Search,
  SendHorizontal,
  Settings2,
  ShieldCheck,
  SlidersHorizontal,
  Sparkles,
  Tag,
  Trash2,
  UploadCloud,
  UserRound,
  X,
} from '@lucide/vue'

type ViewId = 'home' | 'documents' | 'search' | 'chat' | 'settings'
type FileType = 'PDF' | 'DOCX' | 'TXT'
type DetailTab = 'summary' | 'content' | 'meta'
type SearchMode = 'hybrid' | 'keyword' | 'semantic'

interface DocumentItem {
  id: string
  title: string
  type: FileType
  category: string
  tags: string[]
  updated: string
  size: string
  pages: number
  status: '已索引' | '解析中'
  tone: string
  summary: string
}

interface SearchResult {
  docId: string
  before: string
  hit: string
  after: string
  score: number
  page: number
  kind: string
}

interface ChatMessage {
  id: number
  role: 'assistant' | 'user'
  text: string
  sources?: string[]
}

interface NavItem {
  id: ViewId
  label: string
  icon: Component
  badge?: string
}

const validViews: ViewId[] = ['home', 'documents', 'search', 'chat', 'settings']
const urlParams = new URLSearchParams(window.location.search)
const requestedView = urlParams.get('view') as ViewId | null
const activeView = ref<ViewId>(requestedView && validViews.includes(requestedView) ? requestedView : 'home')
const mobileMenuOpen = ref(false)
const uploadOpen = ref(false)
const selectedDoc = ref<DocumentItem | null>(null)
const detailTab = ref<DetailTab>('summary')
const viewMode = ref<'grid' | 'table'>('grid')
const activeCategory = ref('全部文档')
const docSearch = ref('')
const globalSearch = ref('')
const searchQuery = ref('机器学习模型的泛化能力怎么评估？')
const searchMode = ref<SearchMode>('hybrid')
const hasSearched = ref(true)
const chatInput = ref('')
const toast = ref('')
const uploadFiles = ref([
  { name: 'Transformer 学习笔记.pdf', size: '2.4 MB', progress: 100 },
  { name: '毕业论文调研提纲.docx', size: '86 KB', progress: 68 },
])
const selectedSourceIds = ref(['d1', 'd2', 'd4'])

const navigation: NavItem[] = [
  { id: 'home', label: '工作台', icon: LayoutDashboard },
  { id: 'documents', label: '文档库', icon: Files },
  { id: 'search', label: '智能检索', icon: FileSearch },
  { id: 'chat', label: '文档问答', icon: MessageSquareText, badge: 'AI' },
  { id: 'settings', label: '系统设置', icon: Settings2 },
]

const viewMeta: Record<ViewId, { title: string; label: string }> = {
  home: { title: '让每一份资料都能被重新找到', label: '工作台' },
  documents: { title: '文档库', label: '知识空间' },
  search: { title: '智能检索', label: '知识发现' },
  chat: { title: '文档问答', label: 'AI 助手' },
  settings: { title: '系统设置', label: '偏好与连接' },
}

const documents = ref<DocumentItem[]>([
  {
    id: 'd1',
    title: 'Transformer 架构学习笔记',
    type: 'PDF',
    category: '课程学习',
    tags: ['深度学习', '注意力机制'],
    updated: '今天 09:42',
    size: '2.4 MB',
    pages: 32,
    status: '已索引',
    tone: 'sage',
    summary:
      '文档梳理了 Transformer 的编码器—解码器结构、自注意力机制和多头注意力，并对比了循环网络的序列建模方式。核心价值在于并行计算与长距离依赖建模。',
  },
  {
    id: 'd2',
    title: '机器学习模型评估方法综述',
    type: 'PDF',
    category: '论文资料',
    tags: ['机器学习', '模型评估', '泛化'],
    updated: '昨天 18:20',
    size: '6.8 MB',
    pages: 54,
    status: '已索引',
    tone: 'amber',
    summary:
      '综述围绕训练集、验证集和测试集的划分展开，介绍了交叉验证、偏差—方差权衡以及常见评估指标，适合作为模型泛化能力研究的入口材料。',
  },
  {
    id: 'd3',
    title: '毕业设计系统需求说明',
    type: 'DOCX',
    category: '项目开发',
    tags: ['毕业设计', '需求分析'],
    updated: '11 月 22 日',
    size: '186 KB',
    pages: 18,
    status: '已索引',
    tone: 'blue',
    summary:
      '文档定义了个人文档智能整理系统的用户、文档、分类、标签、解析任务、摘要与问答等核心数据结构，并明确了文档上传及检索流程。',
  },
  {
    id: 'd4',
    title: 'RAG 检索增强生成技术整理',
    type: 'TXT',
    category: '论文资料',
    tags: ['RAG', '向量检索', '大模型'],
    updated: '11 月 21 日',
    size: '42 KB',
    pages: 24,
    status: '已索引',
    tone: 'rose',
    summary:
      '资料介绍了 RAG 的文档分块、向量化、召回、重排和答案生成流程，并强调引用溯源与检索失败时的降级处理。',
  },
  {
    id: 'd5',
    title: 'PostgreSQL 全文检索实践',
    type: 'DOCX',
    category: '项目开发',
    tags: ['PostgreSQL', '全文检索'],
    updated: '11 月 20 日',
    size: '524 KB',
    pages: 28,
    status: '已索引',
    tone: 'teal',
    summary:
      '文档说明了中文预分词、tsvector 存储、查询排序以及与 pgvector 向量检索组合使用的方法。',
  },
  {
    id: 'd6',
    title: '知识库产品界面竞品分析',
    type: 'PDF',
    category: '待整理',
    tags: ['界面设计', '竞品'],
    updated: '11 月 19 日',
    size: '3.1 MB',
    pages: 12,
    status: '解析中',
    tone: 'violet',
    summary: '正在提取文档正文和页面结构。',
  },
  {
    id: 'd7',
    title: 'FastAPI 异步任务与异常处理',
    type: 'TXT',
    category: '课程学习',
    tags: ['FastAPI', '异常处理'],
    updated: '11 月 18 日',
    size: '58 KB',
    pages: 15,
    status: '已索引',
    tone: 'slate',
    summary:
      '整理 FastAPI 中后台任务、超时控制、重试与错误响应设计，可作为大模型 API 异常处理模块的参考。',
  },
])

const requestedDocId = urlParams.get('doc')
if (requestedDocId) {
  selectedDoc.value = documents.value.find((doc) => doc.id === requestedDocId) ?? null
}
if (urlParams.get('modal') === 'upload') {
  uploadOpen.value = true
}

const searchResults: SearchResult[] = [
  {
    docId: 'd2',
    before: '评估模型泛化能力时，最常用的方法是使用独立测试集，并通过',
    hit: '交叉验证',
    after: '降低单次数据划分带来的偶然性。对于数据量有限的场景，可以采用 k 折交叉验证。',
    score: 96,
    page: 12,
    kind: '语义命中',
  },
  {
    docId: 'd2',
    before: '模型在训练集上的表现不能直接代表泛化能力，需要关注',
    hit: '偏差与方差',
    after: '之间的权衡，并结合验证曲线判断是否出现过拟合。',
    score: 92,
    page: 18,
    kind: '关键词命中',
  },
  {
    docId: 'd1',
    before: '使用注意力机制时，可通过控制模型复杂度与训练轮次改善',
    hit: '泛化表现',
    after: '，并用独立验证集持续观察性能变化。',
    score: 84,
    page: 27,
    kind: '语义关联',
  },
]

const messages = ref<ChatMessage[]>([
  {
    id: 1,
    role: 'assistant',
    text: '你好，我已经载入 3 篇与机器学习评估相关的资料。你可以直接提问，我会基于文档内容回答，并标明引用来源。',
    sources: ['d2'],
  },
  {
    id: 2,
    role: 'user',
    text: '为什么训练集准确率高，测试集表现却差很多？',
  },
  {
    id: 3,
    role: 'assistant',
    text: '这通常说明模型出现了过拟合：它记住了训练数据中的噪声和局部特征，却没有学到可迁移的规律。建议先检查训练集与测试集分布是否一致，再比较训练—验证曲线。如果训练损失持续下降而验证损失开始上升，可以尝试增加数据、使用正则化或降低模型复杂度，并通过交叉验证重新评估泛化能力。',
    sources: ['d2', 'd1'],
  },
])

const categories = ['全部文档', '课程学习', '论文资料', '项目开发', '待整理']
const suggestedPrompts = [
  '比较文档中提到的模型评估方法',
  '总结 RAG 的主要处理流程',
  '找出系统需求里的核心数据对象',
]

const filteredDocuments = computed(() => {
  const query = docSearch.value.trim().toLowerCase()
  return documents.value.filter((doc) => {
    const inCategory = activeCategory.value === '全部文档' || doc.category === activeCategory.value
    const inQuery =
      !query ||
      doc.title.toLowerCase().includes(query) ||
      doc.tags.some((tag) => tag.toLowerCase().includes(query))
    return inCategory && inQuery
  })
})

const selectedSourceDocs = computed(() =>
  documents.value.filter((doc) => selectedSourceIds.value.includes(doc.id)),
)

function go(view: ViewId) {
  activeView.value = view
  mobileMenuOpen.value = false
}

function getDoc(id: string) {
  return documents.value.find((doc) => doc.id === id)
}

function openDocument(doc: DocumentItem) {
  selectedDoc.value = doc
  detailTab.value = 'summary'
}

function closeDocument() {
  selectedDoc.value = null
}

function submitGlobalSearch() {
  const value = globalSearch.value.trim()
  if (!value) return
  searchQuery.value = value
  hasSearched.value = true
  go('search')
}

function runSearch(prompt?: string) {
  if (prompt) searchQuery.value = prompt
  hasSearched.value = Boolean(searchQuery.value.trim())
  go('search')
}

function usePrompt(prompt: string) {
  chatInput.value = prompt
  go('chat')
}

function toggleSource(id: string) {
  selectedSourceIds.value = selectedSourceIds.value.includes(id)
    ? selectedSourceIds.value.filter((item) => item !== id)
    : [...selectedSourceIds.value, id]
}

function showToast(message: string) {
  toast.value = message
  window.setTimeout(() => {
    toast.value = ''
  }, 2400)
}

function addMockFile() {
  if (uploadFiles.value.some((file) => file.name === '大模型 API 调用记录.txt')) return
  uploadFiles.value.push({ name: '大模型 API 调用记录.txt', size: '32 KB', progress: 24 })
}

function submitUpload() {
  uploadOpen.value = false
  go('documents')
  showToast('原型演示：文件已加入解析队列')
}

function removeUploadFile(index: number) {
  uploadFiles.value.splice(index, 1)
}

function sendChat(prompt?: string) {
  const text = (prompt ?? chatInput.value).trim()
  if (!text) return

  messages.value.push({ id: Date.now(), role: 'user', text })
  chatInput.value = ''

  const reply =
    text.includes('RAG') || text.includes('流程')
      ? 'RAG 的主流程可以概括为：文档解析与分块、向量化、问题改写、混合召回、证据筛选、答案生成和引用定位。在当前系统中，PostgreSQL 同时承担结构化查询与 pgvector 向量检索，检索结果再交给大模型生成带引用的回答。'
      : text.includes('评估') || text.includes('泛化')
        ? '文档中的核心结论是使用独立测试集和交叉验证评估泛化能力，同时观察偏差—方差权衡。若训练集与验证集指标差距持续扩大，应优先排查过拟合。'
        : '我会先在你的文档范围内检索与问题最相关的内容，再根据证据组织答案。当前回答仅用于演示页面交互，不会调用真实大模型。'

  window.setTimeout(() => {
    messages.value.push({
      id: Date.now() + 1,
      role: 'assistant',
      text: reply,
      sources: text.includes('RAG') ? ['d4', 'd5'] : ['d2', 'd1'],
    })
  }, 500)
}

function setDetailTab(tab: DetailTab) {
  detailTab.value = tab
}
</script>

<template>
  <div class="app-shell">
    <div v-if="mobileMenuOpen" class="mobile-overlay" @click="mobileMenuOpen = false"></div>

    <aside class="sidebar" :class="{ 'is-open': mobileMenuOpen }">
      <div class="brand">
        <div class="brand-mark">
          <BookOpen :size="24" stroke-width="1.8" />
        </div>
        <div class="brand-copy">
          <strong>知页</strong>
          <span>个人知识空间</span>
        </div>
        <button class="icon-button sidebar-close" aria-label="关闭菜单" @click="mobileMenuOpen = false">
          <PanelLeftClose :size="19" />
        </button>
      </div>

      <button class="primary-upload" @click="uploadOpen = true">
        <Plus :size="18" />
        添加资料
      </button>

      <nav class="main-nav" aria-label="主导航">
        <span class="nav-caption">知识空间</span>
        <button
          v-for="item in navigation.slice(0, 4)"
          :key="item.id"
          class="nav-item"
          :class="{ active: activeView === item.id }"
          @click="go(item.id)"
        >
          <component :is="item.icon" :size="19" stroke-width="1.8" />
          <span>{{ item.label }}</span>
          <span v-if="item.badge" class="nav-badge">{{ item.badge }}</span>
        </button>
      </nav>

      <div class="sidebar-spacer"></div>

      <div class="index-card">
        <div class="index-card-icon"><Sparkles :size="17" /></div>
        <strong>知识整理完成</strong>
        <p>18 份资料已完成解析与索引</p>
        <div class="index-progress"><span></span></div>
        <button @click="uploadOpen = true">
          继续添加资料
          <ArrowRight :size="15" />
        </button>
      </div>

      <nav class="secondary-nav" aria-label="设置导航">
        <button
          class="nav-item"
          :class="{ active: activeView === 'settings' }"
          @click="go('settings')"
        >
          <Settings2 :size="19" stroke-width="1.8" />
          <span>系统设置</span>
        </button>
      </nav>

      <div class="profile">
        <div class="avatar">邵</div>
        <div>
          <strong>个人空间</strong>
          <span>本地知识库</span>
        </div>
        <ChevronDown :size="16" />
      </div>
    </aside>

    <div class="workspace">
      <header class="topbar">
        <div class="topbar-left">
          <button class="icon-button mobile-menu" aria-label="打开菜单" @click="mobileMenuOpen = true">
            <Menu :size="21" />
          </button>
          <div class="breadcrumb">
            <span>知页</span>
            <ChevronRight :size="14" />
            <strong>{{ viewMeta[activeView].label }}</strong>
          </div>
        </div>

        <div class="topbar-actions">
          <label class="global-search">
            <Search :size="17" />
            <input
              v-model="globalSearch"
              type="search"
              placeholder="搜索文档、观点或问题…"
              aria-label="全局搜索"
              @keyup.enter="submitGlobalSearch"
            />
            <kbd>⌘ K</kbd>
          </label>
          <button class="icon-button" aria-label="通知">
            <Bell :size="19" />
            <span class="notification-dot"></span>
          </button>
          <div class="avatar avatar-small">邵</div>
        </div>
      </header>

      <main class="main-content">
        <section v-if="activeView === 'home'" class="view home-view">
          <div class="welcome-row">
            <div>
              <p class="eyebrow">MONDAY · 11 月 24 日</p>
              <h1>{{ viewMeta.home.title }}</h1>
              <p class="page-lead">让资料从“收藏过”变成“随时找得到、讲得清楚”。</p>
            </div>
            <button class="button button-primary" @click="uploadOpen = true">
              <UploadCloud :size="18" />
              上传新资料
            </button>
          </div>

          <div class="hero-search">
            <div class="hero-copy">
              <span class="hero-kicker"><Sparkles :size="15" /> AI 知识检索</span>
              <h2>从一个问题开始</h2>
              <p>关键词与语义能力协同工作，直接找到散落在资料里的答案。</p>
            </div>
            <label class="hero-input">
              <Search :size="20" />
              <input
                v-model="globalSearch"
                type="search"
                placeholder="例如：Transformer 为什么适合长文本？"
                @keyup.enter="submitGlobalSearch"
              />
              <button type="button" @click="submitGlobalSearch">开始检索</button>
            </label>
            <div class="prompt-chips">
              <button v-for="prompt in suggestedPrompts" :key="prompt" @click="runSearch(prompt)">
                {{ prompt }}
              </button>
            </div>
          </div>

          <div class="stats-grid">
            <article class="stat-card">
              <div class="stat-icon mint"><Files :size="20" /></div>
              <div><span>全部资料</span><strong>24</strong></div>
              <small>本周新增 6 份</small>
            </article>
            <article class="stat-card">
              <div class="stat-icon amber"><Tag :size="20" /></div>
              <div><span>知识标签</span><strong>38</strong></div>
              <small>覆盖 9 个学习主题</small>
            </article>
            <article class="stat-card">
              <div class="stat-icon blue"><Quote :size="20" /></div>
              <div><span>已生成摘要</span><strong>18</strong></div>
              <small>节省约 3.2 小时阅读</small>
            </article>
            <article class="stat-card">
              <div class="stat-icon rose"><MessageSquareText :size="20" /></div>
              <div><span>本周问答</span><strong>42</strong></div>
              <small>来自 12 份文档</small>
            </article>
          </div>

          <div class="dashboard-grid">
            <section class="panel recent-panel">
              <div class="section-heading">
                <div>
                  <span class="eyebrow">RECENTLY OPENED</span>
                  <h2>最近文档</h2>
                </div>
                <button class="text-button" @click="go('documents')">
                  查看全部 <ArrowUpRight :size="15" />
                </button>
              </div>
              <div class="recent-list">
                <button
                  v-for="doc in documents.slice(0, 4)"
                  :key="doc.id"
                  class="recent-row"
                  @click="openDocument(doc)"
                >
                  <div class="file-icon" :class="doc.tone">{{ doc.type }}</div>
                  <div class="recent-copy">
                    <strong>{{ doc.title }}</strong>
                    <span>{{ doc.category }} · {{ doc.pages }} 页 · {{ doc.size }}</span>
                  </div>
                  <div class="recent-time">{{ doc.updated }}</div>
                  <ChevronRight :size="17" />
                </button>
              </div>
            </section>

            <aside class="panel insight-panel">
              <div class="section-heading">
                <div>
                  <span class="eyebrow">KNOWLEDGE PULSE</span>
                  <h2>资料整理建议</h2>
                </div>
                <Sparkles :size="19" class="sparkle-icon" />
              </div>
              <div class="insight-feature">
                <span class="mini-label">待处理</span>
                <strong>2 份资料尚未添加标签</strong>
                <p>为“知识库产品竞品分析”补充标签，可以提高后续检索的召回质量。</p>
                <button @click="go('documents')">现在整理 <ArrowRight :size="15" /></button>
              </div>
              <div class="topic-list">
                <div><span>机器学习</span><strong>8</strong></div>
                <div><span>毕业设计</span><strong>6</strong></div>
                <div><span>大模型应用</span><strong>5</strong></div>
                <div><span>后端开发</span><strong>4</strong></div>
              </div>
            </aside>
          </div>
        </section>

        <section v-else-if="activeView === 'documents'" class="view">
          <div class="page-heading">
            <div>
              <p class="eyebrow">LIBRARY</p>
              <h1>文档库</h1>
              <p class="page-lead">统一管理、整理和回溯你的全部学习资料。</p>
            </div>
            <button class="button button-primary" @click="uploadOpen = true">
              <Plus :size="18" /> 添加资料
            </button>
          </div>

          <div class="category-strip" role="tablist" aria-label="文档分类">
            <button
              v-for="category in categories"
              :key="category"
              :class="{ active: activeCategory === category }"
              @click="activeCategory = category"
            >
              <FolderOpen v-if="category === '全部文档'" :size="16" />
              <span>{{ category }}</span>
              <small v-if="category !== '全部文档'">
                {{ documents.filter((doc) => doc.category === category).length }}
              </small>
            </button>
          </div>

          <section class="library-toolbar">
            <label class="inline-search">
              <Search :size="17" />
              <input v-model="docSearch" type="search" placeholder="在文档库中筛选…" />
            </label>
            <div class="toolbar-right">
              <button class="button button-soft"><SlidersHorizontal :size="16" /> 筛选</button>
              <button class="button button-soft"><Tag :size="16" /> 标签</button>
              <div class="view-switch" aria-label="视图切换">
                <button :class="{ active: viewMode === 'grid' }" aria-label="网格视图" @click="viewMode = 'grid'">
                  <Grid2X2 :size="17" />
                </button>
                <button :class="{ active: viewMode === 'table' }" aria-label="列表视图" @click="viewMode = 'table'">
                  <List :size="17" />
                </button>
              </div>
            </div>
          </section>

          <div v-if="viewMode === 'grid'" class="document-grid">
            <article
              v-for="doc in filteredDocuments"
              :key="doc.id"
              class="document-card"
              tabindex="0"
              @click="openDocument(doc)"
              @keyup.enter="openDocument(doc)"
            >
              <div class="document-preview" :class="doc.tone">
                <div class="preview-topline">
                  <span>{{ doc.type }}</span>
                  <span :class="['status-dot', doc.status === '已索引' ? 'ready' : 'pending']">
                    {{ doc.status }}
                  </span>
                </div>
                <div class="preview-paper">
                  <span class="preview-title">{{ doc.title }}</span>
                  <i></i><i></i><i class="short"></i>
                  <i></i><i class="medium"></i>
                </div>
                <button class="preview-more" aria-label="更多操作" @click.stop>
                  <MoreHorizontal :size="18" />
                </button>
              </div>
              <div class="document-card-body">
                <h3>{{ doc.title }}</h3>
                <div class="tag-row">
                  <span v-for="tag in doc.tags" :key="tag" class="tag">{{ tag }}</span>
                </div>
                <div class="document-meta">
                  <span><Clock3 :size="14" /> {{ doc.updated }}</span>
                  <span>{{ doc.pages }} 页 · {{ doc.size }}</span>
                </div>
              </div>
            </article>
          </div>

          <div v-else class="table-card">
            <table>
              <thead>
                <tr>
                  <th>文档</th>
                  <th>分类</th>
                  <th>标签</th>
                  <th>状态</th>
                  <th>更新时间</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="doc in filteredDocuments" :key="doc.id" @click="openDocument(doc)">
                  <td>
                    <div class="table-doc">
                      <span class="file-icon small" :class="doc.tone">{{ doc.type }}</span>
                      <strong>{{ doc.title }}</strong>
                    </div>
                  </td>
                  <td>{{ doc.category }}</td>
                  <td>{{ doc.tags.join('、') }}</td>
                  <td><span class="status-label" :class="{ pending: doc.status !== '已索引' }">{{ doc.status }}</span></td>
                  <td>{{ doc.updated }}</td>
                  <td><ChevronRight :size="17" /></td>
                </tr>
              </tbody>
            </table>
          </div>

          <div v-if="filteredDocuments.length === 0" class="empty-state">
            <FileSearch :size="26" />
            <strong>没有找到匹配的文档</strong>
            <span>尝试清空筛选条件或更换关键词。</span>
          </div>
        </section>

        <section v-else-if="activeView === 'search'" class="view">
          <div class="page-heading compact-heading">
            <div>
              <p class="eyebrow">SEARCH ACROSS YOUR KNOWLEDGE</p>
              <h1>智能检索</h1>
              <p class="page-lead">同时理解你输入的关键词与问题意图。</p>
            </div>
          </div>

          <div class="search-workbench">
            <label class="search-large">
              <Search :size="21" />
              <input
                v-model="searchQuery"
                type="search"
                placeholder="搜索文档内容，或直接描述你想找的内容…"
                @keyup.enter="runSearch()"
              />
              <button type="button" @click="runSearch()">检索</button>
            </label>
            <div class="search-options">
              <div class="segmented">
                <button :class="{ active: searchMode === 'hybrid' }" @click="searchMode = 'hybrid'">混合检索</button>
                <button :class="{ active: searchMode === 'keyword' }" @click="searchMode = 'keyword'">关键词</button>
                <button :class="{ active: searchMode === 'semantic' }" @click="searchMode = 'semantic'">语义</button>
              </div>
              <div class="search-filters">
                <button><FolderOpen :size="15" /> 全部资料 <ChevronDown :size="14" /></button>
                <button><Tag :size="15" /> 全部标签 <ChevronDown :size="14" /></button>
              </div>
            </div>
          </div>

          <div v-if="hasSearched" class="search-layout">
            <section class="search-results">
              <div class="results-heading">
                <div>
                  <span>找到 3 条相关内容</span>
                  <small>耗时 0.18 秒 · {{ searchMode === 'hybrid' ? 'RRF 融合排序' : searchMode === 'keyword' ? '全文匹配' : '向量相似度' }}</small>
                </div>
                <button><SlidersHorizontal :size="16" /> 相关度</button>
              </div>

              <article
                v-for="result in searchResults"
                :key="`${result.docId}-${result.page}`"
                class="result-card"
                @click="getDoc(result.docId) && openDocument(getDoc(result.docId)!)"
              >
                <div class="result-topline">
                  <span class="result-file-type">{{ getDoc(result.docId)?.type }}</span>
                  <span class="result-kind">{{ result.kind }}</span>
                  <span class="result-score">{{ result.score }}% 相关</span>
                </div>
                <h3>{{ getDoc(result.docId)?.title }}</h3>
                <p>
                  {{ result.before }}<mark>{{ result.hit }}</mark>{{ result.after }}
                </p>
                <div class="result-footer">
                  <span>第 {{ result.page }} 页</span>
                  <span>{{ getDoc(result.docId)?.category }}</span>
                  <span v-for="tag in getDoc(result.docId)?.tags.slice(0, 2)" :key="tag">{{ tag }}</span>
                </div>
              </article>
            </section>

            <aside class="ai-answer-panel">
              <div class="ai-answer-head">
                <div class="ai-orb"><Sparkles :size="18" /></div>
                <div>
                  <span class="mini-label">跨文档洞察</span>
                  <h2>基于 3 条证据的回答</h2>
                </div>
              </div>
              <p class="ai-answer-text">
                评估模型泛化能力时，应使用未参与训练的测试数据，并通过交叉验证降低数据划分偶然性。同时需要观察偏差与方差之间的权衡；若训练集表现很好而验证集明显下降，通常提示过拟合。
              </p>
              <div class="source-stack">
                <button @click="getDoc('d2') && openDocument(getDoc('d2')!)">
                  <span>01</span>
                  <div><strong>机器学习模型评估方法综述</strong><small>第 12、18 页</small></div>
                  <ArrowUpRight :size="15" />
                </button>
                <button @click="getDoc('d1') && openDocument(getDoc('d1')!)">
                  <span>02</span>
                  <div><strong>Transformer 架构学习笔记</strong><small>第 27 页</small></div>
                  <ArrowUpRight :size="15" />
                </button>
              </div>
              <button class="continue-chat" @click="go('chat')">
                <Bot :size="17" /> 继续追问
              </button>
            </aside>
          </div>

          <div v-else class="search-empty">
            <div class="search-empty-icon"><Search :size="28" /></div>
            <h2>输入问题，或从一个关键词开始</h2>
            <p>系统会同时搜索文档标题、正文与语义相近的段落。</p>
          </div>
        </section>

        <section v-else-if="activeView === 'chat'" class="view chat-view">
          <div class="page-heading chat-page-heading">
            <div>
              <p class="eyebrow">ASK YOUR DOCUMENTS</p>
              <h1>文档问答</h1>
              <p class="page-lead">答案只来自所选资料，并保留可点击的引用来源。</p>
            </div>
            <button class="button button-soft"><RefreshCw :size="16" /> 新建对话</button>
          </div>

          <div class="chat-shell">
            <aside class="source-picker">
              <div class="source-picker-head">
                <div>
                  <span class="eyebrow">CONTEXT</span>
                  <h2>回答范围</h2>
                </div>
                <span>{{ selectedSourceIds.length }}</span>
              </div>
              <p>选择文档后，回答将严格限制在所选知识范围内。</p>
              <div class="source-options">
                <button
                  v-for="doc in documents.slice(0, 5)"
                  :key="doc.id"
                  :class="{ checked: selectedSourceIds.includes(doc.id) }"
                  @click="toggleSource(doc.id)"
                >
                  <span class="source-check">
                    <Check v-if="selectedSourceIds.includes(doc.id)" :size="13" />
                  </span>
                  <span class="mini-file" :class="doc.tone">{{ doc.type }}</span>
                  <span class="source-name">{{ doc.title }}</span>
                </button>
              </div>
              <button class="source-add" @click="uploadOpen = true"><Plus :size="15" /> 添加文档</button>
            </aside>

            <section class="conversation">
              <div class="conversation-head">
                <div>
                  <span class="presence-dot"></span>
                  <div><strong>知页助手</strong><small>基于 {{ selectedSourceIds.length }} 篇资料</small></div>
                </div>
                <button aria-label="更多操作"><MoreHorizontal :size="19" /></button>
              </div>

              <div class="message-stream" aria-live="polite">
                <div v-for="message in messages" :key="message.id" class="message-row" :class="message.role">
                  <div v-if="message.role === 'assistant'" class="assistant-avatar"><Sparkles :size="16" /></div>
                  <div class="message-body">
                    <div class="message-bubble">
                      <p>{{ message.text }}</p>
                      <div v-if="message.sources?.length" class="message-sources">
                        <button
                          v-for="sourceId in message.sources"
                          :key="sourceId"
                          @click="getDoc(sourceId) && openDocument(getDoc(sourceId)!)"
                        >
                          <FileText :size="13" />
                          {{ getDoc(sourceId)?.title }}
                        </button>
                      </div>
                    </div>
                    <small v-if="message.role === 'assistant'" class="message-note">
                      AI 生成内容可能有误，请结合引用核对。
                    </small>
                  </div>
                  <div v-if="message.role === 'user'" class="user-avatar"><UserRound :size="16" /></div>
                </div>
              </div>

              <div class="chat-composer">
                <div class="chat-suggestions">
                  <button @click="sendChat('用表格比较这些资料的评估方法')">比较评估方法</button>
                  <button @click="sendChat('总结 RAG 的主要处理流程')">总结 RAG 流程</button>
                </div>
                <div class="composer-box">
                  <textarea
                    v-model="chatInput"
                    rows="2"
                    placeholder="向你的文档提问…"
                    @keydown.enter.exact.prevent="sendChat()"
                  ></textarea>
                  <div class="composer-footer">
                    <span><ShieldCheck :size="14" /> 仅使用已选文档回答</span>
                    <button aria-label="发送消息" :disabled="!chatInput.trim()" @click="sendChat()">
                      <SendHorizontal :size="18" />
                    </button>
                  </div>
                </div>
              </div>
            </section>
          </div>
        </section>

        <section v-else class="view settings-view">
          <div class="page-heading">
            <div>
              <p class="eyebrow">SYSTEM</p>
              <h1>系统设置</h1>
              <p class="page-lead">模型、索引与系统状态的演示页面。</p>
            </div>
            <button class="button button-primary" @click="showToast('原型演示：设置已保存')">
              <Check :size="17" /> 保存设置
            </button>
          </div>

          <div class="settings-grid">
            <section class="settings-card wide-card">
              <div class="settings-heading">
                <div class="settings-icon"><Cpu :size="19" /></div>
                <div><h2>大模型连接</h2><p>用于摘要生成和文档问答。</p></div>
                <span class="connection-status"><i></i> 连接正常</span>
              </div>
              <div class="form-grid">
                <label><span>服务类型</span><select><option>OpenAI 兼容接口</option></select></label>
                <label><span>模型名称</span><input value="gpt-4.1-mini" /></label>
                <label class="full-field"><span>API 地址</span><input value="https://api.example.com/v1" /></label>
                <label class="full-field"><span>API Key</span><input type="password" value="sk-prototype-demo-key" /></label>
              </div>
              <div class="settings-actions">
                <button class="button button-soft"><ShieldCheck :size="16" /> 测试连接</button>
                <small>密钥仅保存在环境变量中，不写入代码。</small>
              </div>
            </section>

            <section class="settings-card">
              <div class="settings-heading">
                <div class="settings-icon amber"><Database :size="19" /></div>
                <div><h2>向量索引</h2><p>PostgreSQL + pgvector</p></div>
              </div>
              <div class="status-list">
                <div><span>索引状态</span><strong class="success-text">健康</strong></div>
                <div><span>向量记录</span><strong>12,840</strong></div>
                <div><span>嵌入模型</span><strong>BGE-small-zh</strong></div>
                <div><span>最后重建</span><strong>11 月 22 日</strong></div>
              </div>
              <button class="button button-soft full-button"><RefreshCw :size="16" /> 重建索引</button>
            </section>

            <section class="settings-card">
              <div class="settings-heading">
                <div class="settings-icon blue"><CircleHelp :size="19" /></div>
                <div><h2>异常降级</h2><p>大模型不可用时的处理策略。</p></div>
              </div>
              <label class="toggle-row">
                <span><strong>保留基础检索</strong><small>模型异常时仍可搜索文档</small></span>
                <input type="checkbox" checked />
              </label>
              <label class="toggle-row">
                <span><strong>自动重试</strong><small>仅对超时与限流错误重试</small></span>
                <input type="checkbox" checked />
              </label>
              <label class="toggle-row">
                <span><strong>显示调用详情</strong><small>展示耗时、错误类型和模型版本</small></span>
                <input type="checkbox" />
              </label>
            </section>

            <section class="settings-card wide-card evaluation-card">
              <div class="settings-heading">
                <div class="settings-icon rose"><Bookmark :size="19" /></div>
                <div><h2>AI 结果评价</h2><p>固定测试集上的最近一次评估结果。</p></div>
                <button class="text-button">查看测试集 <ArrowUpRight :size="15" /></button>
              </div>
              <div class="evaluation-grid">
                <div><span>摘要事实一致性</span><strong>4.3 <small>/ 5</small></strong></div>
                <div><span>问答答案正确性</span><strong>4.1 <small>/ 5</small></strong></div>
                <div><span>引用准确率</span><strong>92<small>%</small></strong></div>
                <div><span>合理拒答率</span><strong>88<small>%</small></strong></div>
              </div>
            </section>
          </div>
        </section>
      </main>
    </div>

    <transition name="fade">
      <div v-if="uploadOpen" class="modal-backdrop" @click.self="uploadOpen = false">
        <section class="upload-modal" role="dialog" aria-modal="true" aria-labelledby="upload-title">
          <div class="modal-head">
            <div>
              <span class="eyebrow">ADD TO LIBRARY</span>
              <h2 id="upload-title">添加学习资料</h2>
              <p>支持 PDF、DOCX、TXT，上传后将自动解析并建立索引。</p>
            </div>
            <button class="icon-button" aria-label="关闭" @click="uploadOpen = false"><X :size="20" /></button>
          </div>

          <div class="drop-zone" @click="addMockFile" @dragover.prevent @drop.prevent="addMockFile">
            <div><CloudUpload :size="28" /></div>
            <strong>拖放文件到这里，或点击选择</strong>
            <span>单个文件不超过 50 MB</span>
          </div>

          <div class="upload-list">
            <div v-for="(file, index) in uploadFiles" :key="file.name" class="upload-file">
              <div class="file-icon small">TXT</div>
              <div class="upload-file-copy">
                <strong>{{ file.name }}</strong>
                <span>{{ file.size }} · {{ file.progress === 100 ? '已完成解析' : '正在解析' }}</span>
                <div class="upload-progress"><i :style="{ width: `${file.progress}%` }"></i></div>
              </div>
              <CheckCircle2 v-if="file.progress === 100" :size="18" class="upload-done" />
              <button v-else aria-label="移除文件" @click="removeUploadFile(index)"><Trash2 :size="17" /></button>
            </div>
          </div>

          <div class="upload-options">
            <label><span>分类</span><select><option>待整理</option><option>课程学习</option><option>论文资料</option></select></label>
            <label><span>标签</span><input placeholder="用逗号分隔，例如：机器学习, 课程" /></label>
          </div>

          <div class="modal-footer">
            <span><ShieldCheck :size="15" /> 文件仅保存在本地知识库</span>
            <div>
              <button class="button button-ghost" @click="uploadOpen = false">取消</button>
              <button class="button button-primary" @click="submitUpload">
                开始解析 <ArrowRight :size="16" />
              </button>
            </div>
          </div>
        </section>
      </div>
    </transition>

    <transition name="slide">
      <div v-if="selectedDoc" class="detail-shell">
        <div class="detail-overlay" @click="closeDocument"></div>
        <aside class="detail-drawer" aria-label="文档详情">
          <div class="detail-head">
            <div class="detail-file-block">
              <div class="file-icon large" :class="selectedDoc.tone">{{ selectedDoc.type }}</div>
              <div>
                <span class="mini-label">{{ selectedDoc.category }}</span>
                <h2>{{ selectedDoc.title }}</h2>
                <p>{{ selectedDoc.pages }} 页 · {{ selectedDoc.size }} · {{ selectedDoc.updated }}</p>
              </div>
            </div>
            <button class="icon-button" aria-label="关闭详情" @click="closeDocument"><X :size="20" /></button>
          </div>

          <div class="detail-toolbar">
            <button class="button button-soft"><Bookmark :size="16" /> 收藏</button>
            <button class="button button-soft" @click="showToast('原型演示：已将摘要复制到剪贴板')">
              <Quote :size="16" /> 复制摘要
            </button>
            <button class="icon-button"><MoreHorizontal :size="19" /></button>
          </div>

          <div class="detail-tabs">
            <button :class="{ active: detailTab === 'summary' }" @click="setDetailTab('summary')">AI 摘要</button>
            <button :class="{ active: detailTab === 'content' }" @click="setDetailTab('content')">原文片段</button>
            <button :class="{ active: detailTab === 'meta' }" @click="setDetailTab('meta')">元数据</button>
          </div>

          <div v-if="detailTab === 'summary'" class="detail-content">
            <div class="summary-banner">
              <div class="ai-orb small-orb"><Sparkles :size="16" /></div>
              <div>
                <strong>AI 内容摘要</strong>
                <span>由大模型生成 · 可追溯至原文</span>
              </div>
            </div>
            <p class="summary-text">{{ selectedDoc.summary }}</p>
            <div class="key-points">
              <h3>核心要点</h3>
              <div><span>01</span><p>明确文档主题与适用场景。</p></div>
              <div><span>02</span><p>提取关键概念、方法或结论。</p></div>
              <div><span>03</span><p>保留页码和原文位置，便于核验。</p></div>
            </div>
            <div class="tag-row detail-tags">
              <span v-for="tag in selectedDoc.tags" :key="tag" class="tag">{{ tag }}</span>
            </div>
            <button class="button button-soft full-button" @click="go('chat'); closeDocument()">
              <MessageSquareText :size="16" /> 基于此文档提问
            </button>
          </div>

          <div v-else-if="detailTab === 'content'" class="detail-content">
            <div class="document-page">
              <span class="page-number">12 / {{ selectedDoc.pages }}</span>
              <h3>模型泛化能力与评估</h3>
              <p>
                模型在训练数据上的表现并不能直接代表它的泛化能力。评估时需要使用未参与训练的独立数据，并通过
                <mark>交叉验证</mark> 降低单次划分带来的偶然性。
              </p>
              <p>
                当训练集性能持续提升而验证集性能开始下降时，通常意味着模型发生了过拟合。此时可以引入正则化、增加样本或降低模型复杂度。
              </p>
              <div class="quote-block">
                “模型评估的目标不是追求训练集上的最高分数，而是估计模型面对未知数据时的稳定表现。”
              </div>
            </div>
          </div>

          <div v-else class="detail-content">
            <div class="metadata-list">
              <div><span>文件格式</span><strong>{{ selectedDoc.type }}</strong></div>
              <div><span>所属分类</span><strong>{{ selectedDoc.category }}</strong></div>
              <div><span>文件大小</span><strong>{{ selectedDoc.size }}</strong></div>
              <div><span>页数</span><strong>{{ selectedDoc.pages }} 页</strong></div>
              <div><span>索引状态</span><strong class="success-text">{{ selectedDoc.status }}</strong></div>
              <div><span>嵌入模型</span><strong>BGE-small-zh</strong></div>
              <div><span>分块数量</span><strong>46</strong></div>
              <div><span>更新时间</span><strong>{{ selectedDoc.updated }}</strong></div>
            </div>
          </div>
        </aside>
      </div>
    </transition>

    <transition name="toast">
      <div v-if="toast" class="toast" role="status">
        <CheckCircle2 :size="17" />
        {{ toast }}
      </div>
    </transition>
  </div>
</template>

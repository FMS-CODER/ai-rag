# 页面原型

个人文档智能整理与语义检索系统的前端页面原型，使用 Vue 3 + TypeScript + Vite 构建。

## 页面

- 工作台：概览、快捷检索、最近文档和整理进度。
- 文档库：分类筛选、标签、网格/表格切换和文档详情抽屉。
- 智能检索：关键词、语义和混合检索的结果展示。
- 文档问答：文档范围选择、对话、引用来源和模拟回答。
- 系统设置：大模型、向量索引与异常降级状态。

上传弹窗、文档摘要、分块详情和提示消息均为本地交互演示，不连接后端、PostgreSQL 或大模型 API。

## 运行

```bash
cd prototype
npm install
npm run dev
```

浏览器打开终端显示的本地地址。

## 预览指定页面

- `/?view=home`
- `/?view=documents`
- `/?view=search`
- `/?view=chat`
- `/?view=settings`

细节页可直接预览：

- `/?view=documents&doc=d2`：打开文档详情抽屉。
- `/?view=home&modal=upload`：打开上传弹窗。

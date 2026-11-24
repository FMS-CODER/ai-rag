# 个人文档智能整理与语义检索系统

毕业设计项目。系统面向个人学习资料管理，规划支持 PDF、DOCX、TXT 文档的上传与整理、全文/语义检索、自动摘要和基于文档内容的问答。

## 当前内容

- `docs/`：需求、场景、系统架构与技术选型文档。
- `prototype/`：Vue 3 + TypeScript + Vite 页面原型，当前不连接后端、数据库或大模型 API。

完整方案采用 Vue 3 + FastAPI + PostgreSQL（pgvector）+ BGE + OpenAI 兼容大模型 API。

## 页面原型

```bash
cd prototype
npm install
npm run dev
```

浏览器打开终端显示的本地地址。页面原型包括工作台、文档库、智能检索、文档问答、系统设置、文档详情和上传弹窗。

## 文档

- [系统需求与架构](docs/系统需求与架构.md)
- [功能开发清单](docs/requirements/功能清单.md)
- [技术栈决策](docs/decisions/000-tech-stack.md)
- [学习笔记](docs/learning/README.md)

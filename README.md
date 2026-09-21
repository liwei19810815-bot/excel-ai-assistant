# Excel AI 助手

在 Excel 任务窗格里与可自配置的大模型对话，让 AI 直接读写、格式化、分析当前工作簿。

面向**公司内网部署**设计：模型走内网网关，数据不出内网；客户端用 **Excel 工具箱的一键安装包**安装，之后功能升级完全无感。

## 快速开始（开发）

```bash
npm install
npm run certs   # 生成本机受信任的开发证书，仅需一次
npm start       # 启动 Vite 并把插件侧载进桌面版 Excel
```

Excel 会自动打开一个新工作簿，「开始」选项卡上出现「AI 对话」按钮。

首次使用需在设置页填写接口地址与模型名称，点「测试连接」验证 —— 它会实际发一次带工具的请求，确认该模型**支持 function calling**（不支持的模型无法操作 Excel）。

## 常用命令

```bash
npm run dev        # 只启动 Vite，用浏览器调 UI（Excel 工具会报错，属正常）
npm run build      # 生产构建，产物在 dist/
npm test           # 运行单元测试
npm run typecheck  # 类型检查
npm run validate   # 校验 manifest.xml
```

## 架构

```
Excel Ribbon ──▶ 任务窗格 (React)
                   │
                   ├─ Agent 循环 ......... src/agent/loop.ts
                   │    消息 → LLM(带工具) → 工具调用 → 结果回灌 → 循环
                   │
                   ├─ LLM 适配层 ......... src/llm/providers/
                   │    OpenAI 兼容 / Anthropic，统一成内部流式接口
                   │
                   ├─ 工具注册表 ......... src/tools/
                   │    按 policy 分级：read / mutate:content / mutate:structure
                   │
                   └─ Excel 桥接 ......... src/excel/
                        串行队列 · 工作簿蓝图 · 快照撤销 · 序列化
```

### 几个关键设计

**工具分级决定安全策略。** `mutate:structure`（删行、删表、`run_script`）默认强制弹确认；`mutate:content`（写值、改格式）自动建快照可一键撤销；`read` 直接执行。

**自动上下文注入。** 每轮对话自动带上工作簿蓝图（表结构、表头、行列数）和当前选区，模型不需要先调工具去"问"。蓝图只含结构不含数据，5 万行的工作簿也只占几百 token。

**必须自建撤销。** Office.js 的写入**不进入 Excel 原生撤销栈**，Ctrl+Z 无效。所以写操作前会快照受影响区域，撤销通过对话面板上的按钮触发。UI 底部常驻提示这一点。

**写入必须串行。** Office.js 的 RequestContext 不是线程安全的，并发 `Excel.run` 会导致随机丢写入。所有调用统一经 `src/excel/coordinator.ts` 排队。

**run_script 是兜底，不是主力。** 结构化工具覆盖不到的长尾需求才用它，且每次执行前完整展示源码让用户确认 —— 这个确认是唯一真正的安全边界，静态检查只是纵深防御。

## 已实现的工具

| 工具 | 策略 | 说明 |
|---|---|---|
| `get_workbook_overview` | read | 工作簿结构蓝图 |
| `read_range` | read | 读区域，支持 markdown/csv/detailed，超限自动采样 |
| `get_selection` | read | 当前选区及内容 |
| `write_cells` | mutate:content | 批量写值/公式，写后读回校验 |
| `format_cells` | mutate:content | 字体、填充、数字格式、对齐、边框、列宽行高 |
| `create_chart` | mutate:content | 柱状/折线/饼图等 9 种图表 |
| `modify_structure` | mutate:structure | 增删行列、增删改名移动工作表 |
| `run_script` | mutate:structure | 自定义 Office.js 代码兜底 |

## 部署

见 [docs/DEPLOY.md](docs/DEPLOY.md)。核心要点：

- 静态资源放内网一台机器，**用域名 + 已被客户端信任的证书**
  （域内 PKI 下发或公网证书）。**不要用自签 + 让安装包装根证书**——
  往用户的受信任根存储塞根证书会降低他整台机器的防护等级
- 客户端用 **Excel 工具箱的一键安装包**安装，不需要管理员权限。
  `install/` 下那两个旧 bat **已作废**（会装根证书、且注册的清单里没有
  `?u=`，白名单分流会整个失效），它们现在会拒绝执行
- 以后升级只更新服务器文件，**用户完全无感**
- Nginx 缓存头必须配对（入口 HTML `no-cache`），否则热更新失效
- 白名单分流与可见性开关见 `server/`，协议见
  `Excel Macro/docs/AI接入与白名单.md`

## 平台兼容

manifest 只声明 `ExcelApi 1.1` 基线，高版本能力（如透视表需 1.8）在运行时探测并降级，保证 Mac / Web / 旧版桌面 Excel 也能装上使用。

Office.js 无法调用 VBA / 宏、Power Query、Power Pivot、OLAP 透视表 —— 这些需要阶段三的本地 sidecar 伴生进程，当前版本尚未实现。

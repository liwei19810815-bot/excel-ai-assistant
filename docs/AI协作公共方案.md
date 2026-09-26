# Claude 与 Codex 自动协作公共方案

本方案适用于所有同时使用 Claude Code 和 Codex 的项目。

## 标准流程

```text
Claude 修改代码并提交
        ↓
创建 review/request.json
        ↓
运行 review:run，自动调用 Codex
        ↓
Codex 只读检查并写入 review/result.json
        ↓
Claude 自动读取结果
        ↓
通过后推送；失败后修复并重新复审
```

## 文件协议

`review/request.json` 保存请求：

```json
{
  "requestId": "唯一 ID",
  "commit": "待复审提交 SHA",
  "createdAt": "ISO 时间",
  "repo": "仓库路径",
  "instructions": "验收要求",
  "attempt": 1
}
```

`review/result.json` 保存结果，`status` 只能是 `running`、`passed`、`failed` 或 `error`。结果必须包含相同的 `requestId` 和 `commit`。

## Claude 使用方式

```powershell
npm run review:run -- --commit HEAD
```

后台拆分执行时：

```powershell
npm run review:request -- --commit HEAD
npm run review:dispatch -- --request-id <request-id>
npm run review:poll -- --request-id <request-id>
```

## Codex 规则

Codex 只做检查，不修改工作区，不创建提交，不推送。发现问题必须写明文件、行号、原因和修复建议。无法运行命令或模型不可用时，结果为 `error`，不得伪报通过。

## 模型回退

默认回退链为：

```text
gpt-5.6-sol → gpt-5.5
```

可通过 `CODEX_REVIEW_MODELS` 调整。禁止固定使用当前账号不支持的模型。


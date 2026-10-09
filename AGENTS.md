# Claude 与 Codex 自动协作

本仓库的 Codex 复审协议在 [docs/AI协作公共方案.md](docs/AI协作公共方案.md) 和 [review/README.md](review/README.md)。

代码修改提交后，运行：

```powershell
npm run review:run -- --commit HEAD
```

只有 `review/result.json` 中当前 requestId 对应的 `status=passed` 才能推送；`failed`、`error` 或超时必须停止推送并继续处理。


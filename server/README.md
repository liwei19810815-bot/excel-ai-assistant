# 网关

任务窗格的静态托管 + 白名单分流的配置接口。

**它不代理大模型。** 模型地址是这里回给任务窗格的，由任务窗格直接去调。

## 起步

```bash
cp server/managed.json.example   server/managed.json
cp server/whitelist.txt.example  server/whitelist.txt

npm run build          # 先出 dist/
node server/gateway.mjs
```

三个配置文件**不在版本库里**（里面会有内网地址、可能有 API Key），
仓库里只保留 `.example` 模板。

## 配置文件

| 文件 | 作用 | 改完要重启吗 |
|---|---|---|
| `whitelist.txt` | 谁用公司配好的模型，一行一个账号，`#` 是注释 | **不用** |
| `managed.json` | 命中白名单时下发的模型配置 | **不用** |

都是每次请求现读的。改一行就要重启服务的话，实际结果往往是"懒得改"。

## 接口

```
GET /api/ai-config?u=<Windows 登录名>
```

```jsonc
// 命中白名单
{ "mode": "managed", "baseUrl": "...", "model": "...",
  "apiKey": "...", "enableRunScript": false, "managedBy": "IT 运维" }

// 未命中
{ "mode": "byok" }
```

## ⚠ 安全边界

**`?u=` 来自客户端，用户可以改。这是分流，不是鉴权。**

前提是白名单里那个模型本身不怕被多用几个人（没有敏感数据访问权、成本可控）。
一旦涉及成本分摊或能读到敏感数据，必须换成网关侧的 Windows 集成认证
（Kerberos/Negotiate），由服务器从票据里解析域账号。

协议细节见 `Excel Macro/docs/AI接入与白名单.md`。

## 生产部署

**不要裸跑 `gateway.mjs`，它没有 TLS。** 放在 nginx 后面由 nginx 终止 https，
配置样例见 `nginx.conf.sample`。

Office.js 要求任务窗格必须是 https，且**证书必须已被客户端信任**
（域内 PKI 下发或公网证书）。证书不受信任时任务窗格是**空白且不报错**的。

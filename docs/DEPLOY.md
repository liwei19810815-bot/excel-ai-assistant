# 内网部署指南（方案 A−）

面向：**有内网台式机，但没有 AD 域管理员权限、没有域名管理权限**的环境。

目标效果：客户端用 **Excel 工具箱的一键安装包**装一次，重开 Excel 即可用；之后每次功能升级用户**完全无感**。

---

## 一、为什么是这个方案

Office 加载项的网页资源是**从 URL 实时加载的**，不是装在用户机器上的。所以只要把资源放在内网一台机器上，以后升级只需更新那台机器的文件，所有用户下次打开 Excel 自动就是新版。

而 Office 的侧载与信任设置全部位于 `HKEY_CURRENT_USER` —— **不需要管理员权限**。

| 环节 | 做法 | 需要管理员 |
|---|---|---|
| 放静态资源 | 内网机器跑 Nginx | 否 |
| HTTPS 地址 | 域名 + **已被客户端信任的证书** | 证书由 IT 统一签发/下发 |
| 让 Excel 认识插件 | 写 `HKCU` 注册表 | 否 |

> **安装程序不再安装任何证书。**
>
> 早先的做法是用自签证书，再由安装脚本 `certutil -user -addstore Root`
> 把根证书装进当前用户的受信任根存储。那确实不需要管理员权限，
> 但它**降低的是用户整台机器的防护等级**——那个根证书能为**任意域**
> 签发被这台机器信任的证书，影响远不止这一个加载项。
> 装个 Excel 插件不该有这种副作用。
>
> 现在的前提是：**网关证书必须已经被客户端信任**，来源二选一
> —— 域内 PKI（根证书由组策略统一下发）或公网证书。
> 只能用自签的话，请让 IT 用组策略下发根证书，而不是让安装包代劳。
>
> ⚠ 证书不受信任时，任务窗格是**空白且不报错**的，用户只会说"AI 打不开"。
> 部署后第一件事：在一台**普通用户的机器**上用浏览器打开网关地址，
> 不弹证书警告才算过。

---

## 二、服务端：一次性配置

### 1. 构建产物

```bash
npm run build
```

产物在 `dist/`。其中 JS/CSS 带内容哈希，`version.json` 由构建脚本生成。

### 2. Nginx 配置（缓存头是关键）

**这一步配错，热更新就会失效** —— WebView2 会把旧版 HTML 缓存住，你部署了新版用户却看不到。

> 完整样例（含 `/api/` 反代、证书与缓存策略）见 `server/nginx.conf.sample`，
> 这里只摘出缓存头这一段。

```nginx
server {
    listen 443 ssl;
    server_name ai.corp.example.com;   # 用域名，不要用 IP（见下一节）

    ssl_certificate     /etc/nginx/certs/ai.corp.example.com.crt;
    ssl_certificate_key /etc/nginx/certs/ai.corp.example.com.key;

    root /var/www/excel-ai/dist;

    # 带哈希的静态资源：内容变了文件名就变，可以永久缓存
    location /assets/ {
        add_header Cache-Control "public, max-age=31536000, immutable";
    }

    # 入口 HTML 与版本文件：必须每次校验，否则拿不到新版
    location ~ \.(html|json)$ {
        add_header Cache-Control "no-cache";
    }

    location / {
        try_files $uri $uri/ =404;
    }
}
```

以后升级只需：`npm run build` → 把 `dist/` 覆盖到 `/var/www/excel-ai/dist`。**不需要通知用户做任何事。**

### 3. 证书

**用域名，不要用 IP。** Office 加载项对 IP 证书的支持很差，而且 IP 一变
所有客户端都要重装。给网关起一个内网域名（如 `ai.corp.example.com`）。

证书来源二选一：

- **域内 PKI 签发**（推荐）：根证书通常已由组策略下发到所有客户端，
  签一张服务器证书即可，客户端天然信任。
- **公网证书**：内网域名也能用 DNS 校验方式签发。

证书和私钥给 nginx 用，配置样例见 `server/nginx.conf.sample`。

> **不要用"自签证书 + 安装包装根证书"那一套。** 那是本项目早先的做法，已废弃：
> 往用户的受信任根存储里塞根证书，等于让它能为**任意域**签发被这台机器
> 信任的证书，影响远不止这一个加载项。真要用自签，也该由 IT 用组策略
> 统一下发根证书，而不是让安装包替每个用户做这个决定。

**验证**（这一步别省）：在一台**普通用户的机器**上用浏览器打开
`https://你的网关地址/`，不弹任何证书警告才算过。
证书不受信任时任务窗格是**空白且不报错**的，排查起来毫无线索。

### 4. manifest 交给工具箱安装包生成

**不要再手工改一份公共 manifest 放共享目录。**
白名单分流依赖 `SourceLocation` 里的 `?u=<用户名>`，而那个参数是
**安装时按当前登录名生成**的。公共 manifest 里没有它，
结果是所有人都落到"自己配模型"，IT 配好的模型对谁都不生效，
**而且不会报任何错**。

正确做法：把 `install/ai/manifest.template.xml`（在 `Excel Macro` 仓库）里的
`<Id>` 换成你自己的 GUID，然后用工具箱的打包脚本，把网关地址传进去：

```powershell
powershell -ExecutionPolicy Bypass -File build\pack.ps1 -Gateway "https://ai.corp.example.com"
```

模板里的 `{{GATEWAY}}` 和 `{{USER}}` 由安装程序在用户机器上替换。

---

## 三、客户端：用工具箱一键安装包

推荐走 **Excel 工具箱的一键安装包**（`Excel Macro` 仓库的 `build\pack.ps1 -Gateway ...`），
它会按当前用户名生成专属 manifest 并完成注册；白名单分流依赖那个
`?u=<用户名>` 参数，手工装的 manifest 没有它，所有人都会落到"自己配"。

运行一键安装包 → 重开 Excel → 出现「AI 对话」按钮。

脚本只做两件事，**都不需要管理员权限**：
1. 按 `%USERNAME%` 生成 manifest，写到 `%LOCALAPPDATA%\ExcelToolbox\ai\`
2. 写 `HKCU\...\Wef\Developer` 告诉 Excel 去哪里找它

**不装任何证书**（原因见上面第一节）。

---

## 四、模型接入

内网模型服务（vLLM / Xinference / One-API / Ollama）只要提供 **OpenAI 兼容端点**即可，插件不关心背后是什么模型。

**让用户零配置**：编辑 `src/store/settings.ts` 里 `PRESETS` 的 `intranet` 项，把 `baseUrl` 和 `model` 改成公司实际地址，重新构建。用户打开即用，连 API Key 都不用填。

**CORS**：浏览器直连要求模型服务返回跨域头。vLLM 用 `--allowed-origins '["*"]`，One-API 在设置里开。若服务端改不了，需加一层 Nginx 反代注入 CORS 头。

**务必用「测试连接」验证** —— 它不只测连通性，还会真的发一次带工具的请求，确认该模型**支持 function calling**。不支持工具调用的模型无法操作 Excel，等实际对话时才发现就太晚了。

---

## 五、升级与维护

| 改动内容 | 用户要做什么 |
|---|---|
| 功能、工具、UI、提示词、模型配置 —— **占迭代量 90%** | **什么都不用做** ✅ |
| 改 `manifest.xml`（增删 Ribbon 按钮、改标签、换地址） | 重跑一次**工具箱的一键安装包** |

为把第二种情况压到最少，`manifest.xml` 已**预留了两个 Ribbon 按钮**并一次性声明了 `ReadWriteDocument` 权限。新功能尽量挂在任务窗格内部，不新增 Ribbon 按钮。

---

## 六、已知限制与排查

**1. 这是微软定位的「测试机制」，不是官方生产部署路径**
共享文件夹目录与 `Wef\Developer` 键功能完全可用、内网广泛采用，但微软不提供支持。需知情接受。

**2. 配置文件漫游环境可能失效**
FSLogix / Citrix UPM / VMware DEM 会重定向 `HKCU` 写入，可能导致注册表键"写了但 Excel 读不到"。**上线前务必在真实客户端环境实测一台。**

**3. 插件的改动无法用 Ctrl+Z 撤销**
Office.js 的写入不进入 Excel 原生撤销栈。这是平台行为，不是 bug。插件自建了快照机制，撤销要用对话面板里每个操作卡片上的「撤销」按钮 —— UI 底部已常驻提示这一点。

**4. 按钮没出现**
- 确认 Excel 已完全退出再重开（任务栏托盘里可能还有残留进程）
- 检查注册表：`reg query "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"`
- ⚠ **不要整个删掉** `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。
  那是**所有 Office.js 加载项共用**的缓存目录，删掉会连带清空别的加载项。
  （本项目早先的卸载脚本就是这么干的，已判为错误做法并作废。）
  一般情况下「完全退出 Excel 再重开」就够了。

**5. 面板空白**
这是**最常见也最没有线索**的一种故障：证书不受信任时，Office 只是不加载，
**不报任何错**。排查顺序：
- 在**出问题的那台机器**上用浏览器打开 `https://你的网关地址/`，
  弹证书警告就说明证书没被信任 —— 找 IT 确认根证书有没有下发到这台机器
- 浏览器能正常打开的话，再看 manifest 里的地址是否和网关地址完全一致
  （含协议、端口、大小写）

**5b. AI 能用，但白名单不生效（所有人都要自己配模型）**
- 打开 `%LOCALAPPDATA%\ExcelToolbox\ai\manifest.xml`，看 `SourceLocation`
  里有没有 `?u=<用户名>`。没有就是 manifest 不是安装程序生成的
- 直接访问 `https://你的网关地址/api/ai-config?u=<某个白名单账号>`，
  确认返回的是 `{"mode":"managed",...}` 而不是 `{"mode":"byok"}`
- 服务端看 `whitelist.txt` 里有没有那个账号（比对时不区分大小写）

**6. 用户看到的还是旧版**
几乎总是缓存头没配对。检查 Nginx 是否对 `.html` 返回了 `Cache-Control: no-cache`。
应急手段：**完全退出 Excel 再重开**。

⚠ 网上常见的"删掉整个 `Wef\` 目录"**不要照做**——那是所有 Office.js 加载项
共用的缓存，会波及别的加载项。

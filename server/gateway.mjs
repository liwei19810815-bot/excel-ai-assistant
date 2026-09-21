/**
 * 网关参考实现（最小可用）。
 *
 * 它只干两件事：
 *   1. 托管任务窗格的静态文件（dist/）
 *   2. 回答 GET /api/ai-config?u=<用户名> —— 这个人用公司配好的模型，还是自己配
 *
 * 【它不代理大模型】。模型地址是这里回给任务窗格的，由任务窗格直接去调。
 * 想让流量都过网关的话另说，那是另一个需求。
 *
 * ⚠ 安全边界：`?u=` 来自客户端，用户可以编辑本机 manifest 改成别人的名字。
 *   **这是分流，不是鉴权。** 前提是白名单里那个模型本身不怕被多用几个人。
 *   一旦涉及成本分摊或能读到敏感数据，必须换成 Windows 集成认证
 *   （Kerberos/Negotiate），由服务器从票据里解析域账号。
 *   详见 Excel Macro 仓库的 docs/AI接入与白名单.md。
 *
 * 用法：
 *   node server/gateway.mjs            # 默认 8080，静态目录 dist
 *   PORT=9000 STATIC_DIR=dist node server/gateway.mjs
 *
 * 【生产上不要直接裸跑这个】。它没有 TLS。
 * 正确做法是放在 nginx 后面，由 nginx 终止 https —— 见 server/nginx.conf.sample。
 * Office.js 要求任务窗格必须是 https，且**证书必须已被客户端信任**。
 */

import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { existsSync, readFileSync } from 'node:fs';
import { extname, join, normalize, resolve, sep } from 'node:path';

const PORT = Number(process.env.PORT ?? 8080);
const STATIC_DIR = resolve(process.env.STATIC_DIR ?? 'dist');
const WHITELIST_FILE = resolve(process.env.WHITELIST_FILE ?? 'server/whitelist.txt');
const MANAGED_FILE = resolve(process.env.MANAGED_FILE ?? 'server/managed.json');
const FEATURE_FILE = resolve(process.env.FEATURE_FILE ?? 'server/feature.json');

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

/**
 * 白名单：一行一个账号，`#` 开头是注释。
 *
 * 【每次请求都重新读】。改完名单不用重启服务——运维改一行就要重启，
 * 实际结果往往是"懒得改"。文件很小，这点开销无所谓。
 */
function loadWhitelist() {
  if (!existsSync(WHITELIST_FILE)) return [];
  try {
    return readFileSync(WHITELIST_FILE, 'utf8')
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l && !l.startsWith('#'))
      .map((l) => l.toLowerCase());
  } catch {
    // 【读不了名单时当作"谁都不在白名单"】，不是"谁都在"。
    // 配置出问题的时候，少给权限永远比多给权限安全。
    return [];
  }
}

function loadManagedConfig() {
  if (!existsSync(MANAGED_FILE)) return null;
  try {
    const c = JSON.parse(readFileSync(MANAGED_FILE, 'utf8'));
    if (!c || typeof c.baseUrl !== 'string' || typeof c.model !== 'string') return null;
    return c;
  } catch {
    return null;
  }
}

/**
 * 功能可见性开关。和白名单是两件事：
 *   白名单  = 这个人用【公司配好的模型】还是自己配
 *   可见性  = 这个人【能不能看到、能不能用】AI 功能
 * 所以对不在白名单的人也要返回可见性。
 *
 *   0 不可见     —— 安装时就不注册这个加载项（唯一能真正"看不见"的办法）
 *   1 可见可使用 —— 默认
 *   2 可见但置灰 —— 按钮在，点开告诉用户已停用
 *
 * 读不出配置时【默认 1】：这是个治理开关，不是安全闸。
 * 配置文件坏掉就让所有人用不了 AI，代价比误开大得多。
 */
function loadVisibility(user) {
  if (!existsSync(FEATURE_FILE)) return 1;
  try {
    const f = JSON.parse(readFileSync(FEATURE_FILE, 'utf8'));

    // 先看有没有针对这个人的单独设置
    const ov = f.visibilityOverrides;
    if (ov && typeof ov === 'object' && user) {
      const hit = Object.keys(ov).find((k) => k.toLowerCase() === user.toLowerCase());
      if (hit && isVisibility(ov[hit])) return ov[hit];
    }

    return isVisibility(f.visibility) ? f.visibility : 1;
  } catch {
    return 1;
  }
}

function isVisibility(v) {
  return v === 0 || v === 1 || v === 2;
}

function sendJson(res, status, body) {
  const s = JSON.stringify(body);
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    // 【绝不缓存】。把某个人移出白名单之后，缓存会让他继续用着公司的模型。
    'cache-control': 'no-store',
  });
  res.end(s);
}

function handleAiConfig(url, res) {
  const user = (url.searchParams.get('u') ?? '').trim();
  const visibility = loadVisibility(user);

  // 没带身份 = 不认识 = 自己配。不要报错，任务窗格会安静退回 byok。
  if (!user) return sendJson(res, 200, { visibility, mode: 'byok' });

  const list = loadWhitelist();
  if (!list.includes(user.toLowerCase())) {
    return sendJson(res, 200, { visibility, mode: 'byok' });
  }

  const managed = loadManagedConfig();
  if (!managed) {
    // 在白名单里、但服务端根本没配模型——这是运维的配置错误。
    // 【不要假装 managed】：回 byok 让用户自己配，至少他能用；
    // 同时在服务端日志里喊出来，免得没人发现。
    console.error(`[ai-config] ${user} 在白名单里，但 ${MANAGED_FILE} 缺失或无效`);
    return sendJson(res, 200, { visibility, mode: 'byok' });
  }

  sendJson(res, 200, {
    visibility,
    mode: 'managed',
    baseUrl: managed.baseUrl,
    model: managed.model,
    ...(managed.apiKey ? { apiKey: managed.apiKey } : {}),
    ...(typeof managed.enableRunScript === 'boolean'
      ? { enableRunScript: managed.enableRunScript }
      : {}),
    managedBy: managed.managedBy ?? 'IT 运维',
  });
}

async function serveStatic(pathname, res) {
  // 【坏掉的百分号编码会让 decodeURIComponent 抛异常】。
  // 在异步函数里抛出去就是个没人接的 rejection，进程日志里出现一堆
  // 看不懂的报错。畸形 URL 直接回 400。
  let decoded;
  try {
    decoded = decodeURIComponent(pathname);
  } catch {
    res.writeHead(400).end('bad request');
    return;
  }

  // 【必须防目录穿越】。不挡的话 /../../etc/passwd 这类请求能读到
  // 静态目录之外的文件。
  //
  // 【前缀比对必须带上分隔符】。只写 full.startsWith(STATIC_DIR) 时，
  // 静态目录是 .../dist 的话，.../dist2/secret 也会通过——
  // 这是很典型的前缀碰撞漏洞。
  const rel = normalize(decoded).replace(/^([/\\])+/, '');
  const full = resolve(join(STATIC_DIR, rel));
  const rootWithSep = STATIC_DIR.endsWith(sep) ? STATIC_DIR : STATIC_DIR + sep;
  if (full !== STATIC_DIR && !full.startsWith(rootWithSep)) {
    res.writeHead(403).end('forbidden');
    return;
  }

  const target = existsSync(full) && !extname(full) ? join(full, 'index.html') : full;
  try {
    const buf = await readFile(target);
    res.writeHead(200, {
      'content-type': MIME[extname(target)] ?? 'application/octet-stream',
    });
    res.end(buf);
  } catch {
    res.writeHead(404).end('not found');
  }
}

const server = createServer((req, res) => {
  const url = new URL(req.url ?? '/', `http://${req.headers.host ?? 'localhost'}`);

  if (url.pathname === '/api/ai-config') {
    if (req.method !== 'GET') return res.writeHead(405).end();
    return handleAiConfig(url, res);
  }

  if (req.method !== 'GET') return res.writeHead(405).end();
  void serveStatic(url.pathname === '/' ? '/index.html' : url.pathname, res);
});

server.listen(PORT, () => {
  console.log(`网关已启动：http://localhost:${PORT}`);
  console.log(`  静态目录：${STATIC_DIR}`);
  console.log(`  白名单　：${WHITELIST_FILE}（${loadWhitelist().length} 个账号）`);
  console.log(`  下发配置：${MANAGED_FILE}（${loadManagedConfig() ? '已就绪' : '缺失'}）`);
  const v = loadVisibility('');
  console.log(`  可见性　：${v}（${['不可见', '可见可使用', '可见但置灰'][v]}）`);
  console.log('');
  console.log('生产部署请放在 nginx 后面由它终止 https，见 server/nginx.conf.sample');
});

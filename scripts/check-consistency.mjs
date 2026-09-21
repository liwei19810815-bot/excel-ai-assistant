/**
 * 一致性静态检查。
 *
 * 守的是【不会让任何单元测试变红的失败】——这类问题在本项目里反复出现，
 * 而且每次都是靠人工评审才发现的：
 *
 *   1. 作废的安装脚本被人"顺手修好"又投入使用
 *   2. 文档里重新出现已经被否掉的做法（装根证书、删公共缓存、双击旧 bat）
 *   3. dist/ 不是当前源码构建的，部署上去是旧版且没有任何征兆
 *   4. 含内网地址或 API Key 的真实配置被提交进库
 *
 * tsc 和 vitest 都发现不了这四类，所以单独做一个检查。
 *
 * 用法：node scripts/check-consistency.mjs
 * 退出码 0 = 通过。
 */

import { readFileSync, existsSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { sourceFingerprint } from './source-fingerprint.mjs';

const problems = [];
const notes = [];

function fail(msg) {
  problems.push(msg);
}

function read(p) {
  return existsSync(p) ? readFileSync(p, 'utf8') : '';
}

//----------------------------------------------------------------------------
// 1. 作废的安装脚本不能含有任何会改动系统状态的命令
//
// 【为什么盯这个】：旧脚本会往用户受信任根存储装根证书，
// 还会 rd /s /q 整个 Wef 缓存目录（那是所有 Office.js 加载项共用的）。
// 它们已被改成"拒绝执行"，但将来很可能有人觉得"修一下还能用"。
//----------------------------------------------------------------------------
const DEPRECATED_BATS = ['install/安装.bat', 'install/卸载.bat'];

// 只看可执行行：REM 注释和 echo 里提到这些词是【说明】，不是执行
const FORBIDDEN_CMDS = [
  { re: /^\s*certutil\b/i, why: '安装证书（会降低用户整机防护等级）' },
  { re: /^\s*reg\s+(add|delete)\b/i, why: '改注册表' },
  { re: /^\s*rd\b/i, why: '删目录（曾经删的是所有加载项共用的 Wef 缓存）' },
  { re: /^\s*rmdir\b/i, why: '删目录' },
  { re: /^\s*del\b/i, why: '删文件' },
  { re: /^\s*copy\b/i, why: '复制文件' },
  { re: /^\s*xcopy\b/i, why: '复制文件' },
];

for (const bat of DEPRECATED_BATS) {
  const text = read(bat);
  if (!text) {
    fail(`找不到 ${bat}（它应该存在并保持"拒绝执行"的状态）`);
    continue;
  }

  for (const line of text.split(/\r?\n/)) {
    // REM 开头的是注释，echo 开头的是给用户看的文字，都不执行
    if (/^\s*(REM|::)/i.test(line) || /^\s*echo\b/i.test(line)) continue;
    for (const f of FORBIDDEN_CMDS) {
      if (f.re.test(line)) {
        fail(`${bat} 里出现了会${f.why}的命令：${line.trim()}\n` +
             `    这两个脚本已作废，必须保持"只提示、不执行"。` +
             `AI 的安装统一走 Excel 工具箱的一键安装包。`);
      }
    }
  }

  // 必须以非零退出，否则批量部署脚本会以为它成功了
  if (!/exit\s+\/b\s+[1-9]/i.test(text)) {
    fail(`${bat} 没有以非零退出码结束——调用方会误以为安装成功`);
  }
  // 没有 pause 的话，双击运行时窗口一闪而过，作废提示根本看不到
  if (!/^\s*pause\s*$/im.test(text)) {
    fail(`${bat} 缺少 pause：双击运行时窗口会一闪而过，用户看不到提示`);
  }
}
notes.push(`作废脚本检查：${DEPRECATED_BATS.length} 个，无可执行的状态改动命令`);

//----------------------------------------------------------------------------
// 2. 文档不能重新出现已经被否掉的做法
//
// 【只查"建议去做"的语气】。这些词出现在"不要这么做"的警告里是正常的，
// 所以下面按行看，命中禁用模式且【同一行没有否定词】才算问题。
//----------------------------------------------------------------------------
const DOC_FILES = ['README.md', 'docs/DEPLOY.md', 'server/README.md'];

const NEGATIONS = ['不要', '不再', '已作废', '已废弃', '已判为', '错误做法',
                   '不该', '别', '⚠', '早先', '旧的', '曾经', '不需要', '不会'];

const BANNED_DOC = [
  { re: /certutil\s+-user\s+-addstore/i, why: '教人安装根证书' },
  { re: /双击\s*[`「]?\s*(install[\\/])?(安装|卸载)\.bat/, why: '教人用已作废的旧安装脚本' },
  { re: /删除\s*`?%LOCALAPPDATA%\\Microsoft\\Office\\16\.0\\Wef\\?`?/, why: '教人删除所有加载项共用的缓存目录' },
  { re: /rd\s+\/s\s+\/q.*Wef/i, why: '教人删除所有加载项共用的缓存目录' },
  { re: /IP\s*\+\s*自签证书/, why: '教人用自签证书（已改为必须使用客户端已信任的证书）' },
];

for (const f of DOC_FILES) {
  const text = read(f);
  if (!text) continue;
  text.split(/\r?\n/).forEach((line, i) => {
    for (const b of BANNED_DOC) {
      if (!b.re.test(line)) continue;
      if (NEGATIONS.some((n) => line.includes(n))) continue;   // 是警告，不是建议
      fail(`${f}:${i + 1} 重新出现了被否掉的做法（${b.why}）：\n    ${line.trim()}`);
    }
  });
}
notes.push(`文档一致性：扫了 ${DOC_FILES.length} 个文件，无复发的旧做法`);

//----------------------------------------------------------------------------
// 3. dist/ 必须是【这份源码】构建出来的
//
// 【最容易出、也最没有征兆的一类】：源码改了、忘了重新构建，
// 部署上去的是旧版。用户报"新功能没有啊"，而所有测试都是绿的。
//
// 【判据是内容指纹，不是修改时间】。mtime 太容易骗也太容易失真：
// touch 一下就能让"产物比源码新"成立；git checkout / 解压 / 跨机器复制
// 之后 mtime 全变；CI 上 checkout 的源码 mtime 常常比缓存产物还新。
// 拿它当验收依据等于没验。
//
// dist/ 不入库，所以本地没有它是正常的（还没构建过），只提示不算失败。
//----------------------------------------------------------------------------
const FP_FILE = 'dist/.build-fingerprint';

if (!existsSync('dist')) {
  notes.push('dist/ 不存在（还没构建过）——部署前记得 npm run build');
} else if (!existsSync(FP_FILE)) {
  fail(
    `dist/ 存在但没有 ${FP_FILE}。
` +
    `    说明它不是用当前的构建流程产出的（或是手工拼的），无法确认对应哪份源码。
` +
    `    请重新 npm run build。`,
  );
} else {
  const { hash: nowHash, fileCount } = sourceFingerprint('.');
  let recorded = null;
  try {
    recorded = JSON.parse(readFileSync(FP_FILE, 'utf8'));
  } catch {
    fail(`${FP_FILE} 读不出来或格式坏了，无法确认产物与源码是否一致。请重新构建。`);
  }

  if (recorded) {
    if (recorded.hash !== nowHash) {
      fail(
        `dist/ 不是当前源码构建的（指纹对不上）。
` +
        `    产物记录：${String(recorded.hash).slice(0, 16)}…（构建于 ${recorded.builtAt ?? '未知'}）
` +
        `    当前源码：${nowHash.slice(0, 16)}…（${fileCount} 个源文件）
` +
        `    直接部署会把旧版发上线，而且【没有任何征兆】。请先 npm run build。`,
      );
    } else {
      notes.push(`dist/ 与当前源码一致（指纹 ${nowHash.slice(0, 12)}…，${fileCount} 个源文件）`);
    }
  }
}

//----------------------------------------------------------------------------
// 4. server/ 下不能有被 git 跟踪的真实配置
//
// 【.gitignore 挡不住"先提交、后加 ignore"】。一旦某个含 API Key 或内网
// 地址的文件已经被跟踪，之后再往 .gitignore 里加规则是没用的——它照样在
// 库里，而且已经进了历史。所以这里直接问 git 要"当前被跟踪的文件"。
//
// server/ 下只允许三类：程序本身、说明、样例模板。
//----------------------------------------------------------------------------
// 【路径必须锚定到 server/ 直属】。只写 /\.example$/ 的话，
// server/subdir/secret.example 这种嵌套路径能绕过去（评审发现）。
const SERVER_ALLOWED = [
  /^server\/README\.md$/,
  /^server\/gateway\.mjs$/,
  /^server\/nginx\.conf\.sample$/,
  /^server\/[^/]+\.example$/,
];

/**
 * 【"不在 git 仓库"和"git 执行失败"必须分开】。
 *
 * 原先是一个 try/catch 全兜住，任何异常都记成"跳过检查"——
 * 那是 fail-open：git 坏了、权限不对、仓库损坏，安全检查统统静默放行，
 * 而输出看起来一切正常。
 *
 * 现在先单独判断是不是仓库：不是就明确跳过（从 zip 解压出来跑就是这种情况）；
 * 是仓库却执行失败，那是真问题，必须报出来。
 */
function insideGitRepo() {
  try {
    return execSync('git rev-parse --is-inside-work-tree', {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim() === 'true';
  } catch {
    return false;
  }
}

if (!insideGitRepo()) {
  notes.push('跳过入库文件检查（当前不在 git 仓库中，例如从压缩包解压出来）');
} else {
  try {
  const tracked = execSync('git ls-files server', { encoding: 'utf8' })
    .split(/\r?\n/)
    .map((l) => l.trim().replace(/\\/g, '/'))
    .filter(Boolean);

  const unexpected = tracked.filter((f) => !SERVER_ALLOWED.some((re) => re.test(f)));
  if (unexpected.length) {
    fail(
      'server/ 下有不该入库的文件（可能含内网地址或 API Key）：\n' +
        unexpected.map((f) => `      ${f}`).join('\n') +
        '\n    只有 gateway.mjs / README.md / nginx.conf.sample / *.example 该入库。' +
        '\n    真实配置保持在本地；已经提交过的要 git rm --cached 撤下来' +
        '（注意它仍留在历史里）。',
    );
  } else {
    notes.push(`server/ 入库 ${tracked.length} 个文件，均为程序/说明/样例`);
  }

  // 顺带扫一遍【所有被跟踪的文件】有没有像密钥的东西。
  // 锁文件里全是包名，噪声太大，跳过。
  const allTracked = execSync('git ls-files', { encoding: 'utf8' })
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean)
    .filter((f) => !/package-lock\.json$/.test(f));

  const SECRET_PAT = [
    { re: /\bsk-[A-Za-z0-9]{16,}/, why: 'OpenAI 风格的 API Key' },
    { re: /\bBearer\s+[A-Za-z0-9._-]{24,}/, why: '硬编码的 Bearer token' },
    { re: /-----BEGIN (RSA |EC )?PRIVATE KEY-----/, why: '私钥' },
  ];
  let scanned = 0;
  for (const f of allTracked) {
    if (!existsSync(f)) continue;
    let text;
    try {
      text = readFileSync(f, 'utf8');
    } catch {
      continue;
    }
    scanned++;
    for (const p of SECRET_PAT) {
      if (p.re.test(text)) fail(`${f} 里疑似含有${p.why}，不应入库`);
    }
  }
    notes.push(`密钥扫描：${scanned} 个入库文件，未发现疑似密钥`);
  } catch (e) {
    // 【在仓库里却执行失败 = 真问题】，不能当作"跳过"放行
    fail(`在 git 仓库里但入库文件检查执行失败：${e?.message ?? e}\n` +
         `    这道检查是防止真实配置/密钥入库的，不能静默跳过。`);
  }
}

//----------------------------------------------------------------------------
// 报告
//----------------------------------------------------------------------------
for (const n of notes) console.log(`  OK   ${n}`);

if (problems.length) {
  console.error('\n一致性检查未通过：\n');
  for (const p of problems) console.error(`  ✗ ${p}\n`);
  process.exit(1);
}

console.log('\n一致性检查通过。');

/**
 * 一致性静态检查。
 *
 * 守的是【不会让任何单元测试变红的失败】——这类问题在本项目里反复出现，
 * 而且每次都是靠人工评审才发现的：
 *
 *   1. 作废的安装脚本被人"顺手修好"又投入使用
 *   2. 文档里重新出现已经被否掉的做法（装根证书、删公共缓存、双击旧 bat）
 *   3. dist/ 比源码旧，部署上去是旧版且没有任何征兆
 *
 * tsc 和 vitest 都发现不了这三类，所以单独做一个检查。
 *
 * 用法：node scripts/check-consistency.mjs
 * 退出码 0 = 通过。
 */

import { readFileSync, existsSync, statSync, readdirSync } from 'node:fs';
import { join, extname } from 'node:path';

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
// 3. dist/ 不能比源码旧
//
// 【这是最容易出、也最没有征兆的一类】：源码改了、忘了重新构建，
// 部署上去的是旧版。用户报"新功能没有啊"，而所有测试都是绿的。
//
// dist/ 不入库，所以本地没有它是正常的（还没构建过），只提示不算失败。
//----------------------------------------------------------------------------
function newestMtime(dir, skip = new Set(['node_modules', 'dist', '.git'])) {
  let newest = 0;
  let newestFile = '';
  const walk = (d) => {
    for (const name of readdirSync(d, { withFileTypes: true })) {
      if (skip.has(name.name)) continue;
      const full = join(d, name.name);
      if (name.isDirectory()) { walk(full); continue; }
      // 只看会进构建产物的源码
      if (!['.ts', '.tsx', '.css', '.html', '.json'].includes(extname(name.name))) continue;
      // 测试文件不进产物，改它不需要重新构建——把它算进来就是误报，
      // 而一条经常误报的检查等于没有检查
      if (/\.(test|spec)\.(ts|tsx)$/.test(name.name)) continue;
      const m = statSync(full).mtimeMs;
      if (m > newest) { newest = m; newestFile = full; }
    }
  };
  walk(dir);
  return { newest, newestFile };
}

if (!existsSync('dist')) {
  notes.push('dist/ 不存在（还没构建过）——部署前记得 npm run build');
} else {
  const { newest, newestFile } = newestMtime('src');
  const distNewest = Math.max(
    ...readdirSync('dist', { recursive: true, withFileTypes: true })
      .filter((e) => e.isFile())
      .map((e) => statSync(join(e.parentPath ?? e.path, e.name)).mtimeMs),
  );
  if (newest > distNewest) {
    fail(
      `dist/ 比源码旧：${newestFile} 改于 ${new Date(newest).toLocaleString()}，` +
      `dist 最新文件是 ${new Date(distNewest).toLocaleString()}。\n` +
      `    直接部署会把旧版发上线，而且【没有任何征兆】。请先 npm run build。`,
    );
  } else {
    notes.push('dist/ 不比源码旧');
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

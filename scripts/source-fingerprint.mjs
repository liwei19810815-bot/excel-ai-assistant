/**
 * 源码指纹：把所有会进构建产物的源文件内容哈希成一个值。
 *
 * 【为什么不用修改时间】：mtime 太容易骗，也太容易失真——
 *   · `touch dist/x` 就能让"产物比源码新"成立
 *   · git checkout / 解压 / 跨机器复制之后 mtime 全变
 *   · CI 上 checkout 出来的源码 mtime 往往比缓存的产物还新
 * 拿它当验收依据，等于验了个寂寞。内容哈希不受这些影响。
 *
 * 构建时把指纹写进 dist/.build-fingerprint；
 * 检查时重新算一遍源码指纹，对不上就说明产物不是这份源码构建的。
 */

import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, statSync, existsSync } from 'node:fs';
import { join, extname, relative, sep } from 'node:path';

/**
 * 会影响构建产物的源文件后缀。
 *
 * 【.xml 必须在里面】：manifest.xml 虽然列在 roots 里，但如果扩展名
 * 不在这张表上，isSource() 会把它滤掉——结果是"列了但没算"，
 * 改了 manifest 检查也不报。评审发现的。
 */
const SOURCE_EXT = ['.ts', '.tsx', '.css', '.html', '.json', '.xml', '.mjs'];

/** 这些目录不进产物 */
const SKIP_DIRS = new Set(['node_modules', 'dist', '.git', '.certs']);

/**
 * 【测试文件不算输入】：它们不进产物，改测试不需要重新构建。
 * 算进来的话每改一次测试就报一次"产物过期"，
 * 而一条经常误报的检查等于没有检查。
 */
function isSource(name) {
  if (!SOURCE_EXT.includes(extname(name))) return false;
  if (/\.(test|spec)\.(ts|tsx)$/.test(name)) return false;
  return true;
}

/**
 * 【不用 readdirSync(..., {recursive:true})】：那个选项要 Node 20+，
 * 而 package.json 声明支持 >=18。自己走一遍，兼容面大得多。
 */
function walk(dir, out = []) {
  if (!existsSync(dir)) return out;
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    if (SKIP_DIRS.has(e.name)) continue;
    const full = join(dir, e.name);
    if (e.isDirectory()) walk(full, out);
    else if (isSource(e.name)) out.push(full);
  }
  return out;
}

/**
 * 算指纹。
 *
 * 【路径分隔符要归一化】：Windows 上是 \，Linux 上是 /，
 * 不归一化的话同一份源码在两个平台上指纹不同，CI 永远报过期。
 */
export function sourceFingerprint(root = '.') {
  // 【构建脚本本身也是输入】。改了 gen-version.mjs 之类的东西却没重新
  // 构建，产物同样是过期的——而那种情况只看 src/ 是发现不了的。
  const roots = [
    'src',
    'public',
    'scripts',
    'manifest.xml',
    'package.json',
    // 【锁文件决定依赖版本，依赖变了产物就变了】。
    // 只看 package.json 是不够的：^1.2.0 这种范围下，
    // 锁文件一变实际装的版本就变，而 package.json 一个字没动。
    'package-lock.json',
    'tsconfig.json',
    'vite.config.ts',
    'index.html',
  ];

  const files = [];
  for (const r of roots) {
    const full = join(root, r);
    if (!existsSync(full)) continue;
    if (statSync(full).isDirectory()) files.push(...walk(full));
    else if (isSource(full)) files.push(full);
  }

  files.sort();   // 目录遍历顺序不保证稳定，排序后才可复现

  const h = createHash('sha256');
  for (const f of files) {
    h.update(relative(root, f).split(sep).join('/'));
    h.update('\0');

    // 【换行必须归一化】。Windows 上 git 的 autocrlf 会在 checkout 时
    // 把 LF 换成 CRLF——文件内容"没变"，字节却变了。
    // 不归一化的话，任何一次 checkout 之后指纹都对不上，
    // 这条检查就变成了常态误报，而常态误报等于没有检查。
    h.update(readFileSync(f, 'utf8').replace(/\r\n/g, '\n'));
    h.update('\0');
  }
  return { hash: h.digest('hex'), fileCount: files.length };
}

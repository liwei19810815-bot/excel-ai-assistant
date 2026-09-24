/**
 * 工具注册入口 —— 导入即注册（各模块顶层调用 register）。
 * 新增工具只需在此处 import 一行。
 */
import './excel/read';
import './excel/write';
import './excel/structure';
import './excel/chart';
import './sandbox/runScript';
import './sidecar/powerQuery';
import './sidecar/runMacro';
import './powerpoint/read';
import './powerpoint/write';

export * from './registry';

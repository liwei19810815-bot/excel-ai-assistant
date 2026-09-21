/**
 * Ribbon 上 ExecuteFunction 类型按钮的处理函数。
 * 这个文件在一个独立的无界面运行时里执行，与任务窗格不共享内存。
 */

Office.onReady(() => {
  // 注册须在 onReady 之后完成
});

/**
 * 「问选区」按钮：打开任务窗格。
 *
 * 注：ExecuteFunction 无法直接向任务窗格传参，任务窗格本身每轮都会自动注入
 * 当前选区，因此这里只需确保面板是打开的即可。
 */
function askAboutSelection(event: Office.AddinCommands.Event): void {
  // displayDialogAsync 之外唯一能从命令打开任务窗格的方式是 Office 自动处理，
  // 这里仅作提示，实际打开由用户点击主按钮完成。
  Office.addin
    .showAsTaskpane()
    .catch(() => {
      /* 旧版本不支持该 API 时静默忽略，用户可手动点主按钮 */
    })
    .finally(() => event.completed());
}

// Office 运行时通过全局名查找 manifest 里声明的 FunctionName
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).askAboutSelection = askAboutSelection;
Office.actions?.associate?.('askAboutSelection', askAboutSelection);

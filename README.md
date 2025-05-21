# 这个Excel加载项只能帮你做一件事：

## 查假期

- 输入日期（比如 `2024/5/1`），告诉你之后所有的法定休息日
- 输入年份（比如 `2024`），直接看全年的法定休息日
- 输入负数（比如 `-1`），查过去 6 个月到未来 1 年的法定休息日
- 不输入参数，直接返回所有的法定休息日

## 怎么用？

1. **将下载的[`.xlam`](https://objects.githubusercontent.com/github-production-release-asset-2e65be/983929041/1d79fc71-d553-4529-ac77-bdf41f0bbf73?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=releaseassetproduction%2F20250521%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20250521T030219Z&X-Amz-Expires=300&X-Amz-Signature=fa31361d5777e9935729802bf82bd7ed3bd03ed53a7de33f88121b3aa328b176&X-Amz-SignedHeaders=host&response-content-disposition=attachment%3B%20filename%3DLegalRestDay_EX.xlam&response-content-type=application%2Foctet-stream)文件在Excel中进行加载项导入**
   1. 将文件放置在`C:\Users\Users\AppData\Roaming\Microsoft\AddIns`文件夹下
   2. Excel操作步骤：`文件` → `选项` → `加载项` → `Excel 加载项 转到` → `勾选加载项以启用`

2. **在单元格里直接写公式**，比如：
   - `=NETWORKDAYS.INTL("2025-9-29","2025-10-9","0000000",LegalRestDay())` → 计算两个日期之间的所有工作日数
   - `=LegalRestDay(TODAY())` → 显示今天之后所有的法定休息日
   - `=LegalRestDay(2024)` → 显示 2024 全年的法定休息日

## 关于假期数据的说明

- 只包含2011年至当前年份的法定休息日日期
- 每年 12 月自动下载下一年的假期数据
- 打开文件时静默检查更新（需要联网）
- 更新失败不影响已有数据使用
- **数据来源**：[laomor/holiday-data](https://github.com/laomor/holiday-data/All%20Years) 这个 GitHub 仓库

## 注意事项
- 如果公司网络限制 GitHub 访问，可能无法更新
- 自动同步最新数据到后台时有失败的概率

*（用起来就像个带自动更新功能的日历查询器，数据维护交给 GitHub 作者，您只管用）*

# 格致 Excel 对比编辑工具

本地网页工具，用于对比和编辑 Excel 工作簿。项目包含页面入口、Excel 处理依赖文件和导出回归测试。

## 启动与验证

- 双击启动：`启动Excel工具.command`
- 页面入口：`index.html`
- 导出回归测试：`node tests/export-regression.cjs`

首次使用前确认工作簿文件来源和导出目标。测试可用于检查导出回归，不会代替对真实工作簿的人工核验。该目录未发现独立依赖清单；请勿删除 `xlsx.full.min.js`，页面依赖它处理工作簿。

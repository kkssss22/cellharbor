# CellHarbor 的定位与调研依据

调研日期：**2026 年 11 月 4 日**。

格子助手 CellHarbor 面向希望少记公式、少找菜单的 Excel 初学者，把常见的清洗、整理、检查和转换任务做成中文操作卡片。它是一套操作界面与安全措施的组合，不是新的表格计算方法，也不声称是 GitHub 上第一个或唯一的 Excel 工具箱。

## 检索范围与名称

本次检索了公开 GitHub 项目、项目 README 和微软 Office Add-ins 文档。关键词包括 `excel addin utilities trim remove duplicates text tools office.js`，并对 `CellHarbor`、`GridCaddy`、`SheetNavi` 做了精确名称检索。

检索结果未发现这三个名称对应的同名表格软件；最终使用 **CellHarbor / 格子助手**。搜索结果受索引和可见性限制，不能证明名称全球无人使用，也不是商标查询。没有遍历所有 GitHub 仓库；未公开项目和未被索引的代码不在本次范围内。

## 已有项目

| 项目 | 已有能力 | CellHarbor 的定位 |
| --- | --- | --- |
| [MEx Automate](https://github.com/pauluwaifo/mex-automate) | 免费的 Excel 任务窗格插件，包含检查错误、清洗、合并、报表、命令操作、预览和撤销；无需服务器或 AI 模型。它是本次发现的近似方案。 | 以中文初学者为主要使用者，以可搜索的固定操作卡片、示例和前后预览降低使用门槛。任务窗格、无服务器、预览和撤销本身不是独有能力。 |
| [Excel Essential Tools](https://github.com/Max-Schmeling/excel-essential-tools) | 提供 50 多项桌面 Excel 工具，包括文本清洗、重复检查、拆分和合并表格；使用 `.xlam` 加载项，不支持 Excel 网页版。 | 使用 Office.js 任务窗格组织常见任务，把选区、公式保护和结果写入位置放进同一操作流程。不能仅凭使用 Office.js 就宣称所有 Excel 版本都兼容。 |
| [RelaxTools Addin](https://github.com/RelaxTools/RelaxTools-Addin) | 多功能 Excel 加载项，覆盖查找替换、文本转换、重复检查、文字插入删除等；其说明列出 Windows 桌面 Excel 支持。 | 聚焦一组可理解、可预览的日常动作，而非覆盖全部高级功能；采用中文说明和统一的步骤。 |

上表是对项目公开文档的比较，不是完整代码审计或性能测评。已有项目也可能提供未在上述摘要中列出的保护措施。

## 产品差异应怎样表述

设计目标是把约 30 项常见操作放进同一个容易上手的流程：**选择单元格 → 选择工具 → 检查前后变化 → 应用结果**。具体可用工具以界面和 README 的实际清单为准。

- **中文任务说明。** 使用“去掉首尾空格”“统一大小写”“检查重复”等直接描述结果的名称，并给出示例。
- **先看变化。** 应用前显示处理范围、预计变化与结果示例，避免把预览按钮和写入按钮混在一起。
- **保护公式。** 区分公式、计算结果和显示文本；普通文本清洗不应把公式静默替换成当前计算值。
- **明确写入位置。** 需要改变表格形状的任务可输出到新工作表；覆盖选区的操作应检查预览之后数据是否改变。
- **有范围的撤销。** 插件自己的撤销应说明保存什么、能恢复哪一次操作，并防止覆盖用户随后手动修改的内容。它不是整个工作簿的备份系统。

这些差异是体验目标，不构成“其他项目都没有”的结论。功能数量也不能单独证明好用；真正的判断标准是任务能否完成、变化能否理解，以及失败时是否保留原数据。

## Office.js 实现依据

以下均为微软或 GitHub 官方资料，供维护者核对实现。

| 需求 | 官方接口与参考 |
| --- | --- |
| 读取用户当前选择 | `context.workbook.getSelectedRange()` 返回当前连续选区，可加载地址后展示给用户。[获取工作表范围](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-get) |
| 区分原值、显示文字和公式 | `Range.values` 读取原始值或公式计算结果，`Range.text` 读取显示文本，`Range.formulas` 读取或写入公式。不要仅凭 `values` 恢复含公式的选区。[读写值、文本和公式](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-set-get-values) |
| 在新表输出结果 | 使用工作表集合创建工作表，再向明确的目标范围写入数据。[管理工作表](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-worksheets) |
| 标记需要检查的单元格 | `Range.format.fill.color` 设置填充色；格式变化与数据变化应分别说明。[设置范围格式](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-set-format) |
| 撤销能力 | Excel JavaScript API 提供撤销和操作分组，但有不支持撤销的接口；部分调用会清除已有撤销历史。不能承诺所有版本、所有动作都能用 Ctrl+Z 恢复。[撤销支持与限制](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities) |
| 加载 Office 运行环境 | 在 HTML 的 `head` 中从微软 CDN 引用 Office.js，并等待 `Office.onReady()`。Office.js 不应打进应用自己的代码包。[连接 Office.js 与前端框架](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/connect-to-javascript-frameworks) |
| 声明插件入口与权限 | Office 加载项包含清单和 Web 应用；XML 加载项清单支持 Excel，入口与内容 URL 应使用 HTTPS。[加载项清单](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests) |
| 托管静态界面 | GitHub Pages 支持 HTTPS，可用于托管静态页面和资源；托管页面不等于完成 Office 插件安装或通过商店审核。[GitHub Pages HTTPS](https://docs.github.com/en/pages/getting-started-with-github-pages/securing-your-github-pages-site-with-https)、[发布 Web 应用并更新清单](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-visual-studio#publish-the-add-in) |
| 检查宿主兼容性 | 不同 Office 版本和平台支持的 API 集合不同。应声明必要的 requirement sets，并对不可用功能提供清晰提示。[Office 版本与要求集](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/office-versions-and-requirement-sets) |

## 使用边界

CellHarbor 不需要把工作表交给 AI 服务来执行这些固定操作。插件页面的托管和微软 Office.js 仍涉及资源加载，因此“数据在本地处理”不应被写成“完全没有网络访问”。

纯浏览器中的演示只能验证工具逻辑和界面；真实 Excel 的读写、权限、保护单元格、表格、合并单元格和动态数组行为需要在支持的宿主里另行验证。发布记录应区分浏览器测试和真实 Excel 测试，不把前者表述成后者。

对含公式、日期、前导零编号、长整数、错误值或复杂格式的数据，展示文本不一定等于底层值。转换工具应把规则说清楚；无法安全处理的选区应说明原因，不能靠静默转换完成操作。

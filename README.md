# Engineering Accreditation Skill

面向工程认证课程材料的 Codex Skill：成绩与课程目标达成度核查、Excel 图表、Word/PDF 材料整理，以及批量交付验证。

A Codex skill for engineering accreditation coursework: grade auditing, course outcome attainment analysis, Excel charts, and Word/PDF document workflows.

本项目将课程材料处理中的计算规则、保真要求和核验方法整理成可复用的技能。由 Codex 根据当前文件和任务选择流程，辅助脚本完成独立计算及交付清单检查。使用说明和技能正文以中文为主。

## 能处理什么

| 场景 | 工作内容 |
|---|---|
| 成绩与达成度 | 核对满分、贡献权重、目标映射、公式、统计人数和舍入口径 |
| Excel 图表 | 检查汇总柱状图、学生散点图、数据引用、自动标签和阈值线 |
| 个人评分表 | 从已核对的数据填写模板，检查分项、总评、等级和签名结构 |
| 作业与实验报告 | 按名册整理真实材料、处理嵌套压缩包、修改获准评分字段 |
| Word 批量材料 | 复用确认模板、检查身份与题分、验证 DOC 转换和目录完整性 |
| 扫描试卷 PDF | 拆页、转正、重命名、按授权局部修改题头并检查原始批阅 |
| 模拟材料 | 生成明确标注的演示数据或作答样例，独立检查计分约束 |

这些是技能指导的工作流程，实际执行需要 Codex 和相应工具。本仓库附带的四个脚本分别用于达成度复算、文件清单核验、DOCX格式与渲染证据检查，以及发现已运行的 Excel；没有内置完整的 Word/PDF 生成或渲染程序。

2026-10-19同步了真实作业/实验报告的目标识别、最小修改、异常及打包口径，补充原表直接引用、逐段系列选择与图表联动核验，并新增作业格式检查器。公开版保留通用规则，不包含私人项目实例、历史学生数据或原始截图。

2026-10-22补充分层模拟版式、精确身份命名与模拟标识、段落黑点控制清理后的分页复查，以及全量自动检查和代表页面实看的区分。学生身份示例全部虚构，未同步私人课程实例。

## 目录

```text
engineering-accreditation-skill/
├── README.md
├── LICENSE                         # MIT 许可证
├── PUBLICATION.md
├── .gitignore
├── examples/
│   └── attainment/                 # 完全虚构的复算示例
│       ├── scores.csv
│       ├── rules.json
│       └── expected-results.json
├── skills/
│   └── engineering-accreditation/
│       ├── SKILL.md                # 技能入口
│       ├── agents/openai.yaml      # Codex 展示信息
│       ├── references/             # 按任务读取的通用流程
│       └── scripts/                # 四个辅助脚本
└── tests/
    ├── test_helpers.py             # 虚构数据与临时文件验证
    └── test_homework_format.py     # 合成DOCX与过期渲染证据验证
```

## 环境

- 使用技能：能够读取本地技能文件并执行相应任务的 Codex 环境。
- 运行三个 Python 脚本及测试：Python 3.10 或更新版本，仅使用标准库，无须安装 pip 依赖。
- 发现已运行的 Excel：Windows、PowerShell、桌面版 Microsoft Excel，以及已打开的工作簿。非 Windows 环境可使用其他流程，跳过这个辅助脚本。
- Word/PDF 制作、格式转换、渲染和原生 Office 核验：自行提供相应工具，或使用目标环境中已安装的文档、表格、PDF 技能。它们没有随本项目打包。

脚本本身不调用 AI，也不要求 API 密钥；使用 Codex 的账号和运行权限由使用者的环境决定。

## 安装到 Codex

实际技能文件夹是 `skills/engineering-accreditation`，不是整个仓库根目录。

### 让 Codex 从 GitHub 安装

将下面请求发给 Codex：

```text
请使用 skill-installer 从这个 GitHub 仓库安装 skills/engineering-accreditation：
https://github.com/8aus1R/engineering-accreditation-skill
```

### 手动安装

下载并解压仓库，将 `skills/engineering-accreditation` 文件夹完整复制到当前用户的 `~/.agents/skills/` 下。最终应存在 `~/.agents/skills/engineering-accreditation/SKILL.md`。

Windows 用户可以在仓库根目录运行以下 PowerShell 命令。目标已存在时会停止，先自行比较或备份已有版本：

```powershell
$skillInstallRoot = Join-Path $HOME '.agents/skills'
$skillTarget = Join-Path $skillInstallRoot 'engineering-accreditation'
if (Test-Path -LiteralPath $skillTarget) {
    throw '同名技能已存在，请先比较或备份，避免覆盖。'
}
New-Item -ItemType Directory -Path $skillInstallRoot -Force | Out-Null
Copy-Item -LiteralPath './skills/engineering-accreditation' -Destination $skillTarget -Recurse
```

复制完成后重新启动 Codex，并在技能列表检查名称。旧环境或 skill-installer 也可能使用 `~/.codex/skills/`；遵循当前环境的安装位置，不重复安装同名技能。

技能结构及安装机制参考 [OpenAI 官方技能说明](https://learn.chatgpt.com/docs/build-skills) 和 [官方 skill-installer](https://github.com/openai/skills/tree/main/skills/.system/skill-installer)。

## 使用示例

在自己的任务中提供实际文件、允许修改区域和输出位置：

```text
使用 $engineering-accreditation，只读核查这份成绩与达成度工作簿。
以当前名册为准，独立复算全部学生和目标汇总，列出具体单元格及异常。
```

```text
使用 $engineering-accreditation，按已确认模板批量制作个人评分表。
分数以提供的 Excel 为准，保留原签名结构，核验所有输出并另存副本。
```

```text
使用 $engineering-accreditation，制作一份虚构学生的模拟作答样例。
文件名和说明中标注“模拟”，核对题分、总分、页数和版式。
```

## 运行虚构复算示例

以下命令都从仓库根目录运行：

```text
python skills/engineering-accreditation/scripts/check_attainment.py --csv examples/attainment/scores.csv --config examples/attainment/rules.json --output work/demo-calculation.json --observed examples/attainment/expected-results.json
```

示例中 `DEMO-001` 至 `DEMO-003` 均为虚构身份，分数和规则完全为演示编写，不来自任何课程或学生。第三行的 0 是有效分数，用于展示“0 与缺失不同”。

| 示例目标 | 未舍入均值 | 两位小数 |
|---|---:|---:|
| goal1 | 0.536666… | 0.54 |
| goal2 | 0.5 | 0.50 |

脚本输出沿用了 `excel_numeric_results_matched` 状态名，表示提供的观察值与计算结果匹配。此示例使用仓库内的预期结果文件，**没有读取 Excel，不能据此宣称工作簿或图表核验通过**。

输入规范、观察结果结构与适用限制见 [Excel 复算流程](skills/engineering-accreditation/references/excel.md)。真实课程的数据、满分、权重和排除规则须重新确认。

## 交付清单检查

本地准备好 Word 文件、SHA-256 清单和可选 CSV 后，可运行：

```text
python skills/engineering-accreditation/scripts/verify_homework_delivery.py --root local-delivery --inventory inventory.json --csv scores.csv --path-column file
```

本仓库没有附带 `local-delivery`。具体清单格式及参数见 [批量交付流程](skills/engineering-accreditation/references/batch-student-homework.md)。该脚本只核对文件集合、哈希、目录数量与 CSV 引用，不检查答案、评分、Word 可打开性或页面渲染。

## 新生成作业的格式与渲染证据检查

在已完成真实渲染及版式检查后，提供当前哈希绑定的证据：

```text
python skills/engineering-accreditation/scripts/check_homework_format.py --root local-delivery --expected-count 1 --expected-score-count 20 --score-color C00000 --render-report work/render-evidence.json --output work/format-check.json
```

人数、题分字段数和颜色都是当前任务参数。它检查题分前方框/全角空格、部分正文评语、段落/样式控制、文件覆盖及提供的渲染证据；不判答案、不比较Excel、不查看图片像素。卷名字体、全部位置的全角空格、页脚及续页标题仍需独立检查。`--structure-only`只用于调试，不能代替交付核验；完整证据结构见[作业格式与全量检查](skills/engineering-accreditation/references/homework-delivery-gate.md)。这套规则不授权删除真实原始作业的批阅或签名。

## 测试

```text
python -m unittest discover -s tests -v
```

测试使用虚构分数和临时文件，检查正常计算、真实数值 0、缺失/重复/越界输入、观察结果不一致、输出覆盖保护，以及交付文件篡改、额外文件、路径越界和 CSV 覆盖。格式检查器另覆盖跨run的全角空格、扣分评语、段落/样式继承、真实列表、漏文件和过期渲染哈希。测试不会启动 Office，也不验证真实课程文件或打印版式；合成微型PNG只用于检验证据匹配分支。

## 数据与发布边界

- 本公开版不包含真实名册、成绩、答卷、教师签名、原始课程文件、私人路径或原始工作截图。脱敏处理和验证说明见 [PUBLICATION.md](PUBLICATION.md)。
- 真实、订正和模拟材料分别标识。模拟作答与反推分项不能作为原始学生记录或真实评分证据。
- 只修改当前用户授权区域，默认保留原件并另存副本。签名复用须具备相应授权，保留图片不表示签署者重新批准了内容。
- 运行日志可能包含学生键、本地完整路径和工作簿名。真实任务的计算输出、清单和错误日志不要直接提交到公共仓库。
- `.gitignore` 排除常见课程文件、缓存和本地输出；它不会清理已跟踪文件或 Git 历史，也不能替代提交前检查。

本项目是材料处理与核验工具，不提供认证机构背书，不替代当前认证标准、课程规则或专业评审。

## 许可证

本项目的代码和文档采用 [MIT 许可证](LICENSE)，版权署名为 `Copyright (c) 2026 8aus1R`。

允许免费使用、复制、修改、分发和商用；分发副本或实质性部分时须保留版权及许可声明。MIT 不要求修改版开源，软件按现状提供，不附带担保。具体条款以 [LICENSE](LICENSE) 英文全文为准。

真实课程素材、教师签名和第三方模板没有随本项目提供，其权利不因本仓库的许可证而改变。

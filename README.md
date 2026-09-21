# 数据处理中心

一个纯前端、可部署到 GitHub Pages 的数据处理工具台。文件会在浏览器内存中读取和处理，不会上传到服务器。

## 本地启动

```powershell
python -m http.server 8899
```

访问：

```text
http://localhost:8899/
```

## 发布成可分享网站

1. 在 GitHub 创建一个新仓库，例如 `data-processing-center`。
2. 把本目录内容推送到仓库的 `main` 分支。
3. 打开仓库 Settings -> Pages。
4. Source 选择 `GitHub Actions`。
5. 等待 Actions 跑完，访问：

```text
https://你的用户名.github.io/data-processing-center/
```

也可以不用 Actions，直接在 Settings -> Pages 里选择 `Deploy from a branch`，分支选 `main`，目录选 `/root`。

## 已支持

- 智能匹配：按关键列把源表字段填充到目标表
- 格式转换：Excel / CSV / JSON 转换
- 数据清洗：去空格、去空行、去重、去换行
- 报表合并：多个文件或多个 Sheet 合并
- 数据比对：按主键识别新增、删除、修改
- 敏感脱敏：手机号、身份证、姓名、邮箱
- 数据拆分：按列值或固定行数拆分并打包
- 派生列：用 `[列名]` 公式生成新字段
- 分组汇总：求和、计数、平均、最大、最小
- 多列对比：支持指定表头行，选择多列后保留全部明细行并取最大值、最小值
- 图表生成：按字段聚合并绘制图表

## 技术

- HTML / CSS / JavaScript
- SheetJS
- JSZip
- Chart.js

## 付款申请单整理

侧栏进入“付款申请单整理”，上传 `.xlsx` 或 `.xls` 付款明细，选择工作表，核对各付款主体后生成预览。内置 `assets/payment-template.docx`，也可上传同结构 `.docx` 替换。无需服务器处理或额外安装。

- 表头使用“付款申请单信息”中的字段：申请日期、付款主体、申请部门、经办人、付款总金额、供应商全称、开户银行及账号、合同编号、合同总金额、已付金额、本次申请付款金额、票据状态、付款事由、备注。允许表头前有空行。
- 按付款主体分组，不合并或删除重复明细；按“本次申请付款金额”逐笔合计，金额以整数分计算。顶部小写、顶部人民币大写、底部合计使用同一个结果。“付款总金额”仅校验；存在差异时需要勾选核实确认。
- 申请日期、部门、经办人可修改；同主体有不同值时必须填写统一值。计划付款日期可填写，空白时保留模板占位线。付款类型固定“采购货款”、付款方式固定“对公转账”。
- 银行信息完整保留；审批签字、财务办结文字、空栏及“一、三、四”编号保持模板内容。只调整版式，不修改上传源文件。
- 自动调整字体、行距和边距，每个主体一页 A4 横向。极端内容低于 Word 可表示的 1 pt 字号仍无法容纳时会报错，不交付截断文件。预览显示最终字号。
- Word 为可编辑的真实 DOCX；PDF 是高分辨率固定打印版，内容已按 85% 缩放，打印应选择实际大小/100%。不同 Word/WPS 或字体替换可能改变分页，固定打印以 PDF 为准。
- 识别到异常金额、缺少主体、公式错误或长编号数字精度风险时，显示源行，修正后重新上传。其他明细空白及“/”原样保留。输入变化会立即使旧下载失效。

### 开发与验证

本项目是静态网站，无生产打包步骤。无需安装依赖即可运行 `npm test`（金额/分组回归）及 `npm run check`（脚本语法）。使用本地 HTTP 服务预览，不能直接通过 `file://` 读取内置模板。

浏览器验收：先启动 `python -m http.server 8899 --bind 127.0.0.1`，在已有 Playwright 环境运行 `npm run test:browser`。可通过 `PLAYWRIGHT_MODULE` 指定现有 Playwright 包路径、`CHROME_PATH` 指定现有浏览器、`QA_URL` 指定本地网址、`QA_OUTPUT` 指定仓库外输出目录。测试使用虚构数据，不包含真实付款资料。

浏览器测试覆盖样例导出、大小写金额、固定选项、明细保留、模板替换、异常拦截、工具切换、XLS、移动端和大数据压缩。还需使用 Word/WPS 或 LibreOffice 渲染生成 DOCX，确认 `stress.docx` 的 3 个主体实际为 3 页且办结栏完整；不能只依据预览或 PDF 页数判断 Word 分页。

渲染后运行 `python tests/payment-docx-verify.py assets/payment-template.docx /tmp/qa/stress.docx /tmp/render/stress.pdf`（按实际路径替换；验证环境需有 pypdf），检查实际页数、全部明细、银行账号、审批及办结文字与保留的模板包成员。

`payment-core.js` 负责金额与分组，`payment-document.js` 负责模板和共享排版，`payment-tool.js` 负责界面。新增 PDF 依赖为本地 vendored pdf-lib（MIT 许可证见 `vendor/pdf-lib.LICENSE.md`）；公用模板之外的真实明细和验收文件不得提交到仓库。

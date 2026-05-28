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
- 多列对比：选择多列后按行或按列取最大值、最小值
- 图表生成：按字段聚合并绘制图表

## 技术

- HTML / CSS / JavaScript
- SheetJS
- JSZip
- Chart.js

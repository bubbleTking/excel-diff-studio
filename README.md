# Excel Diff Studio

一个本地运行的 Excel 可视化对比工具。把两个 Excel 文件拖进页面后，浏览器会直接显示不同的单元格、所在 Sheet、File 1 内容和 File 2 内容。

Excel Diff Studio is a local, browser-based Excel comparison tool. Drop two Excel files into the page and it will show different cells, sheet names, File 1 values, and File 2 values immediately.

## 特点

- 文件只在本地浏览器解析，不上传服务器
- 支持拖拽导入 `.xlsx`、`.xlsm`、`.xls`、`.csv`
- 按 Sheet 过滤差异
- 搜索单元格地址、Sheet 名或单元格内容
- 可切换忽略首尾空格、比较公式、严格空值
- 对旧值和新值做颜色区分
- 支持中文 / English 双语界面

## Features

- Files are parsed locally in your browser and are never uploaded
- Drag-and-drop support for `.xlsx`, `.xlsm`, `.xls`, and `.csv`
- Filter differences by sheet
- Search by cell address, sheet name, or cell content
- Toggle trim spaces, compare formulas, and strict blank handling
- Old and new values are color coded
- Chinese / English bilingual UI

## 本地运行

## Local Development

```bash
npm install
npm run dev
```

打开终端里显示的本地地址，例如：

```text
http://127.0.0.1:5173/
```

## 测试文件

## Sample Files

项目附带两个示例 Excel：

- `samples/sample_file_1.xlsx`
- `samples/sample_file_2.xlsx`

可以直接拖进页面查看差异效果。

## 构建

## Build

```bash
npm run build
```

## 逻辑测试

## Smoke Test

```bash
node test-smoke.mjs
```

## Portable Python Script

如果只想在另一台电脑上跑 Excel 比较，可以复制 `excel_compare_portable.py`。它只依赖 Python 标准库，不需要安装第三方包。

直接运行后按提示粘贴或拖入两个 `.xlsx` 文件路径：

```bash
python3 excel_compare_portable.py
```

Windows 用户也可以直接双击：

```text
run_excel_compare_windows.bat
```

默认报告会保存到用户的 Downloads 文件夹：

```text
Downloads/excel_compare_report.html
```

也可以继续用命令行参数：

```bash
python3 excel_compare_portable.py file1.xlsx file2.xlsx
```

It generates a local visual HTML report in the user's Downloads folder by default.

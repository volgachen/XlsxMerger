# XlsxMerger：一款好用的excel合并脚本插件
.exe 可执行文件下载地址：[v0.1.0](https://github.com/volgachen/XlsxMerger/releases/tag/v0.1.0)

#### 【最新说明】

- 2022-07-08 修复了处理公式单元格出现的错误，现在合并时仅合并数值不合并公式。

## 适用场景

多个excel中同名子表合并
子表格式要求：
- 子表开头为若干行说明与表头
- 子表中间每行为一个统计项，且在中间无合并单元格情况
- 子表最后可能有结束行标志，通过该行第一个单元格中的文字识别


## 使用说明

我们提供了在windows系统中的可执行文件，可以在`cmd`或`powershell`中通过命令调用。同时，我们也鼓励用户通过修改python脚本实现更加多样化的功能。

以下对如何在命令行调用可执行文件进行说明。

1、请将命令行路径切换至`generateConfig.exe`与`process.exe`所在目录，或将目录添加到环境变量中。

2、准备一个模板文件 [TemplateFile]。模板文件中每个表格应只保留说明及表头，删除所有内含数据

3、生成配置文件[ConfigFile]

`.\generateConfig [TemplateFile] [ConfigFile]`

4、修改配置文件
- 修改Path值为所有待统计文件的通配符，一般为`目录\*.后缀名`
- 修改Output值为想要的输出文件名
- 删除Sheets中不需要统计的表格
- 修改Sheets中每一个表格，`prefix_lines`应为统计项之前的说明及表头所占行数，`posfix_string`为终止行的第一个单元格中的文字

4、运行`.\process --cfg [ConfigFile]`

## 异常警告

程序针对以下遇到的情况进行警告输出

- 统计文件中缺少需要统计的子表
- 统计文件某子表中未出现终止行文字

## 更新日志

- 2022-07-08 修复了处理公式单元格出现的错误，现在合并时仅合并数值不合并公式。

## 代码打包

1、生成`version file`。

`create-version-file [YAML_IN_FILE] --outfile [TXT_OUT_FILE] --version [VERSION]`

2、使用pyinstaller打包

`pyinstaller -F -w --version-file [TXT_OUT_FILE] [PYTHON_FILE]`

## License

Released under MIT License.
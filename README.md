# Excel工作簿拆分器正式版

此项目可生成两种可直接分发给普通用户的 Windows 软件：

- `ExcelWorkbookSplitter_Portable.zip`：绿色免安装版
- `ExcelWorkbookSplitter_Setup.exe`：标准安装程序，支持桌面快捷方式和卸载

## 推荐：使用 GitHub 在线生成 Windows 软件

本方法不要求本机安装 Python，也不会再遇到 BAT 中文编码问题。

1. 在 GitHub 新建一个空仓库。
2. 将本压缩包中的所有文件上传到仓库根目录，包括 `.github` 文件夹。
3. 打开仓库的 `Actions` 页面。
4. 选择 `Build Windows software`。
5. 点击 `Run workflow`。
6. 等待完成后，在页面底部下载 `ExcelWorkbookSplitter-Windows`。
7. 解压后即可得到绿色版 ZIP 和安装版 EXE。

## 发布正式版本

在 GitHub 创建标签，例如 `v1.0.0`，工作流会自动创建 Release，并附带：

- `ExcelWorkbookSplitter_Portable.zip`
- `ExcelWorkbookSplitter_Setup.exe`

## 本机打包

在已经安装 Python 的 Windows 电脑上双击 `build_windows.cmd`。

本地打包只生成绿色 EXE：
`dist\ExcelWorkbookSplitter.exe`

## 文件说明

- `workbook_splitter_gui.py`：程序源码
- `assets/app.ico`：软件图标
- `installer/setup.iss`：安装包配置
- `.github/workflows/build-windows.yml`：云端自动构建配置
- `build_windows.cmd`：本机打包脚本，全文仅使用 ASCII，避免 CMD 乱码

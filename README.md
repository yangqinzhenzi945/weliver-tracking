# Weliver 出库状态查询工具

Windows 上运行的 Python 自动查询工具。读取 Excel 中 `weliver` 工作表 **A 列**的 WOSO 运单号，进入 Weliver 的 **Transport tracking** 页面，逐单查询：若查询结果的 **Consignment state** 栏至少有 **1 个绿色圆点**，则在原行的 **C 列**写入 `已发出`。原 C 列已标记 `已发出` 的订单会跳过；其他结果保持不变。B 列不修改。

> 注意：绿色圆点的网页元素选择器是根据截图编写的，**尚未经过真实网页结构验证**。首次请用 2–3 条已知订单测试并核对结果；登录需手动进行。

## Windows 使用方法

1. 安装 Python 3（需要支持 `py` 命令）。
2. 双击 **`安装环境.bat`**（只需第一次运行）。
3. 把自己的 Excel 文件复制到此文件夹，关闭正在打开的 Excel。
4. 双击 **`开始查询.bat`**，输入 Excel 文件名（包括 `.xlsx`）。
5. 在自动打开的浏览器登录网站，进入 **Transport tracking**；回到命令窗口按回车。
6. 查询完成后，在同一文件夹查看 `原文件名_已更新.xlsx`。

也可手动运行：

```cmd
py -m pip install -r requirements.txt
py -m playwright install chromium
py weliver_tracking.py "cargoboard待查询出库.xlsx"
```

## Excel 结构

| A 列 | B 列 | C 列 |
| --- | --- | --- |
| WOSO 物流号 | ERP ID | 出库情况 |

工作表名称必须是 `weliver`（不区分大小写）。脚本不会覆盖原 Excel，定期保存输出文件。

## GitHub 上传

建议将仓库设为 **Private**（私有）。不要上传真实 Excel、订单数据、登录信息或浏览器用户配置。

GitHub 用于存放代码；**GitHub Pages 不能运行这个需要登录的浏览器自动化脚本**，仍需要在自己的 Windows 电脑运行。

## 数据与安全

- 订单列表只在本地读取和写入。
- 浏览器登录由本人完成；Chromium 的本地配置放在用户主目录的 `.weliver_browser_profile`。
- 查询出错时可能产生页面截图 `weliver_error_*.png`，请不要上传 GitHub。


## 网页版（推荐）

本工具提供 **Streamlit 本地网页界面**，无需每次输入命令，支持上传 Excel、查看查询进度和下载已更新文件。**它在你的 Windows 电脑运行，不是公共云端网站。**

1. 在 GitHub 仓库点击 `Code → Download ZIP`，解压到电脑。
2. **首次运行**：双击 `安装环境.bat`，安装依赖及 Chromium。
3. 双击 `打开网页工具.bat`，网页通常会自动打开；若没有打开，在浏览器访问 `http://localhost:8501`。
4. 上传你的 Excel（包含 `weliver` 工作表；A 列 WOSO，C 列状态）。
5. 点击「开始查询」，在自动弹出的 Chromium 窗口**手动登录** Weliver，并进入 Transport tracking；程序检测到列表后查询所有尚未标记已发出的订单。
6. 在网页点击「下载查询结果 Excel」。

如果上次未关闭自动化 Chromium，请先关闭旧窗口再重新执行。注意不要在自动查询过程中操作其页面。

**重要**：网站绿色状态点的 CSS/DOM 结构仅根据截图推断，未经过实际登录验证。请先拿 2–3 条已知状态的 WOSO 做测试再批量使用。没有检测到的行不会写入「已发出」。遇到查询异常请核查网页提示和结果。



<img width="492" height="595" alt="image" src="https://github.com/user-attachments/assets/0a5c708f-d52d-48ce-b9ba-af906c3193fa" />



# Excel 工作簿管理器安装文档

## 一、工具简介

**Excel 工作簿管理器** 是一个 Excel 桌面插件，用于管理当前打开的多个 Excel 文件。

安装后，Excel 右侧会出现一个「工作簿管理器」面板，可以查看、搜索、切换、保存、关闭当前打开的 Excel 工作簿。

主要功能包括：

* 查看当前打开的 Excel 文件
* 按文件夹 / 文件名前缀 / 保存状态分组
* 搜索文件名和路径
* 双击快速切换工作簿
* 保存单个文件
* 关闭单个文件
* 保存整个分组
* 关闭整个分组
* 生成当前打开文件清单

---

## 二、适用环境

### 1. 操作系统

支持：

* Windows 10
* Windows 11

暂不支持：

* macOS
* Linux

### 2. Office 版本

需要安装桌面版 Microsoft Excel。

支持常见版本：

* Microsoft Office 2016
* Microsoft Office 2019
* Microsoft Office 2021
* Microsoft 365 桌面版

暂不支持：

* Excel 网页版
* WPS 表格
* macOS 版 Excel

### 3. 运行环境

电脑需要具备以下运行环境：

* .NET Framework
* Microsoft Visual Studio Tools for Office Runtime

如果安装时提示缺少运行环境，请先安装：

**Microsoft Visual Studio Tools for Office Runtime**

---

## 三、安装包内容

安装包解压后，一般包含以下内容：

```text
ExcelWorkbookTabManager/
  setup.exe
  ExcelWorkbookTabManager.vsto
  Application Files/
```

注意：

**不要只复制 setup.exe。**

安装时需要保留整个文件夹结构，否则可能无法正常安装。

---

## 四、安装步骤

### 步骤 1：解压安装包

将安装包解压到本地目录，例如：

```text
D:\Tools\ExcelWorkbookTabManager
```

或者：

```text
C:\Users\你的用户名\Downloads\ExcelWorkbookTabManager
```

---

### 步骤 2：关闭 Excel

安装前建议先关闭所有 Excel 窗口。

如果 Excel 正在运行，可能导致插件安装后不能立即加载。

---

### 步骤 3：运行安装程序

双击安装包中的：

```text
setup.exe
```

如果系统弹出安全提示，选择：

```text
运行
```

如果出现 Office 自定义项安装提示，选择：

```text
安装
```

---

### 步骤 4：启动 Excel

安装完成后，打开 Excel。

如果插件加载成功，Excel 右侧会出现：

```text
工作簿管理器
```

面板中会显示当前打开的 Excel 文件。

---

## 五、检查插件是否安装成功

如果打开 Excel 后没有看到右侧面板，可以按下面方式检查。

打开 Excel 后进入：

```text
文件 → 选项 → 加载项
```

在底部找到：

```text
管理：COM 加载项
```

点击：

```text
转到
```

查看列表中是否有：

```text
ExcelWorkbookTabManager
```

如果有，但没有勾选，请勾选它，然后点击「确定」。

---

## 六、插件使用说明

### 1. 查看当前打开的文件

插件会自动扫描当前 Excel 实例中打开的工作簿。

面板中会显示：

* 文件名
* 所在分组
* 完整路径
* 是否已保存
* 是否只读
* 当前 Sheet

---

### 2. 搜索文件

在搜索框中输入关键词，例如：

```text
预算
名单
合同
库存
排班
```

插件会自动过滤匹配的 Excel 文件。

---

### 3. 切换文件

在右侧面板中，双击某个 Excel 文件，即可快速切换到对应工作簿。

---

### 4. 保存单个文件

选中某个具体文件后，点击：

```text
保存文件
```

即可保存当前选中的工作簿。

---

### 5. 关闭单个文件

选中某个具体文件后，点击：

```text
关闭文件
```

如果文件未保存，插件会提示：

```text
是否保存后关闭？
```

可以选择：

```text
保存
不保存
取消
```

---

### 6. 保存整个分组

选中某个分组后，点击：

```text
保存分组
```

插件会保存该分组下所有未保存的 Excel 文件。

---

### 7. 关闭整个分组

选中某个分组后，点击：

```text
关闭分组
```

如果分组内存在未保存文件，插件会提示是否保存后关闭。

---

### 8. 生成文件清单

点击：

```text
生成清单
```

插件会在当前工作簿中新增一个 Sheet，列出当前打开的所有 Excel 文件。

清单内容包括：

* 分组
* 文件名
* 完整路径
* 当前 Sheet
* 是否保存
* 是否只读

这个功能适合用于工作记录、任务交接、文件整理和当前处理范围确认。

---

## 七、重要说明

### 1. 插件只能识别当前 Excel 实例里的文件

如果你已经提前打开了多个 Excel，而插件没有识别出来，可能是因为这些文件在不同的 Excel 进程中。

建议使用方式：

```text
先打开带插件的 Excel
再通过这个 Excel 打开其他文件
```

或者：

```text
关闭所有 Excel
重新打开 Excel
再打开需要管理的文件
```

---

### 2. 建议一个任务一个文件夹

为了获得更好的分组效果，建议按任务管理 Excel 文件。

例如：

```text
年会筹备/
  01_活动预算.xlsx
  02_嘉宾名单.xlsx
  03_物料清单.xlsx
  04_人员分工.xlsx
```

插件可以按文件夹自动分组。

---

### 3. 不建议直接在微信、QQ、浏览器下载目录中处理文件

建议先把文件保存到固定文件夹，再打开处理。

这样更容易分组和管理，也更方便后续查找。

---

## 八、常见问题

### 问题 1：安装后 Excel 右侧没有面板

请检查：

```text
文件 → 选项 → 加载项 → COM 加载项 → 转到
```

确认：

```text
ExcelWorkbookTabManager
```

是否已经勾选。

如果没有勾选，请勾选后点击「确定」。

---

### 问题 2：COM 加载项里没有这个插件

可能原因：

* 安装没有成功
* 安装包文件不完整
* 只运行了 setup.exe，但缺少 Application Files 文件夹
* 缺少 VSTO Runtime

解决方式：

```text
重新解压完整安装包
关闭 Excel
重新运行 setup.exe
```

---

### 问题 3：提示缺少运行环境

需要安装：

```text
Microsoft Visual Studio Tools for Office Runtime
```

安装完成后，重新运行：

```text
setup.exe
```

---

### 问题 4：插件只显示一个 Excel 文件

这通常是因为其他 Excel 文件在不同的 Excel 进程里。

解决方式：

```text
关闭所有 Excel
重新打开 Excel
在同一个 Excel 窗口中打开多个文件
点击插件里的“刷新”
```

---

### 问题 5：关闭文件时报错

可能原因：

* 文件已经被手动关闭
* 文件处于只读状态
* 文件被其他程序占用
* Excel 正在弹出保存确认框

可以点击：

```text
刷新
```

重新加载当前工作簿列表。

---

## 九、卸载方式

打开 Windows 设置：

```text
设置 → 应用 → 已安装的应用
```

搜索：

```text
ExcelWorkbookTabManager
```

点击：

```text
卸载
```

也可以在控制面板中卸载：

```text
控制面板 → 程序 → 程序和功能 → ExcelWorkbookTabManager → 卸载
```

卸载后重新打开 Excel，插件面板将不再显示。

---

## 十、推荐使用流程

建议日常这样使用：

```text
1. 先打开 Excel
2. 确认右侧出现“工作簿管理器”
3. 在 Excel 中打开需要处理的多个文件
4. 使用右侧面板搜索、切换、保存、关闭文件
5. 处理完成后生成文件清单
6. 保存并关闭对应分组
```

一句话总结：

**先打开插件所在的 Excel，再打开需要管理的文件。**

---

## 十一、工具定位

Excel 工作簿管理器不是为了替代 Excel 原有功能。

它解决的是一个更日常的问题：

> 当 Excel 文件打开太多时，怎么快速知道哪个是哪个。

它希望让 Excel 多文件办公变得更清楚、更顺手。

像管理浏览器标签页一样，管理 Excel 工作簿。

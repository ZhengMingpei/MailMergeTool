# MailMergeTool
该工具可以指定1个xlsx作为数据来源、指定1个或多个docx作为模板批量进行**邮件合并**操作生成大量docx。

**目前稳定命令行版为1.0版本，图形界面版已发布，目前为1.0beta版本，源码待完善后上传。**

图形界面版运行图如下：

![图形界面版截图](https://github.com/ZhengMingpei/MailMergeTool/blob/master/截图2019-12-13-211139.png)

工具核心部分的代码使用[Bouke写的docx-mailmerge](https://github.com/Bouke/docx-mailmerge)，做了部分修改。

## 支持平台及下载链接

可执行程序目前仅支持windows的64bit系统，python源码支持所有满足依赖的平台。

不需要安装有MS office，但是工具运行需要导入xlsx和docx文件，这些文件的编辑需要使用MS office或者WPS office的较新版本（2007及以后）。

### 下载链接：

源码：敬请期待

[windows命令行版(释放绿色版exe文件，即开即用)](https://github.com/ZhengMingpei/MailMergeTool/releases/download/v1.0/maincli.exe)

[windows64位图形界面版](https://github.com/ZhengMingpei/MailMergeTool/releases/download/v1.0/mailmergeTool-GUI-cn-x64.exe)

linux(将会以源码形式)

### 使用方法

#### windows平台下的图形界面版使用方法

下载完成后，双击运行下载好的`mailmergeTool-GUI-cn-x64.exe`，准备好xlsx和docx文件（格式可见下文）后，根据界面指示操作即可

#### 命令行版使用方法（含数据及模板介绍）

1. 指定你的xlsx文件作为数据源，比如当前目录下有`demodata.xlsx`，数据需形如：

|  姓名   | 性别  | 民族  |
|  ----  | ----  | ----  |
|宋江	|男|	汉族
|晁盖	|男|	汉族
|李逵	|男|	汉族

    第一行为数据域名字，第二行以后为数据域，每行为一条数据，数据域可为空

2. 指定你的docx文件作为模板，比如当前目录下有`.\test\1.docx`，也可以直接指定目录`.\test\`，程序会自动查找指定目录中的所有docx文件

    docx文件中包含插入的邮件合并域（wps中的插入方法是：插入-文档组件-域-邮件合并），形如：

![commit1](https://github.com/ZhengMingpei/MailMergeTool/blob/master/commit1.png)

3. 运行`maincli.exe .\demodata.xlsx .\test\1.docx`或者`maincli.exe .\demodata.xlsx .\test\`，将会生成以当前时间命名的文件夹，内含生成的所有docx. 形如：

![commit2](https://github.com/ZhengMingpei/MailMergeTool/blob/master/commit2.png)

### linux平台下的使用方法

获取源码后，`pip`安装要求的库后

`python .\maincli.py .\demodata.xlsx .\test\`

## 其他说明

* 目前仅支持xlsx作为数据源，1条数据生成单个docx文件

* 暂不支持自定义数据域范围，默认生成所有合法的数据域，上面例子中的表格数据，会对每个模板根据所有三条数据生成三个docx文件。

* 图形版已支持自定义输出目录，不支持自定义输出文件名等

## 作者备注

本工具旨在本人方便邮件合并操作，觉得能用就在github提供分享，尚未经过大量测试，不对工具使用者造成的任何损失负责。

如有疑问或任何其他问题，可通过邮箱联系我（zhengmingpei@163.com）


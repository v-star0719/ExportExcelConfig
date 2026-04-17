# ExportExcelConfig
在Excel中编辑游戏的配置，导出游戏。也支持导出多语言。

## 介绍
使用NPOI库读取excel，导出对应的cs配置脚本，以及对应的xml数据文件。应用那边可以直接读取xml。如果觉得慢，可以直接把excel数据序列化为二进制数据；或者读取xml然后序列化成二进制数据。

由于使用的是.net framework 4.8，NPOI最高到2.3.0版本，否则会引入一堆库。

如果使用新版本的NPOI，请使用最新版.net
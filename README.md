# ExportExcelConfig
在Excel中编辑游戏的配置，导出到游戏内。也支持导出多语言。 
- 配置导出到Xml，Unity中可以直接序列号Xml加载配置数据，然后将配置写入ScriptableObject。
- 配置支持生成配置对应的C#代码。
- 多语言为一列一个语种，支持无限个语种。

## 介绍
使用NPOI库读取excel，导出对应的cs配置脚本，以及对应的xml数据文件。应用那边可以直接读取xml。如果觉得慢，可以直接把excel数据序列化为二进制数据；或者读取xml然后序列化成二进制数据。

由于使用的是.net framework 4.8，NPOI最高到2.3.0版本，否则会引入一堆库。

如果使用新版本的NPOI，请使用最新版.net

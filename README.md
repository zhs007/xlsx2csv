# xlsx2csv

批量的将xlsx或xls文件转换为csv或json格式的命令行工具。

配合 [ds-lang](https://github.com/zhs007/ds-lang) 使用，可以实现从数据建模一直到最终的数据加载都一键操作。

编码集
---
只支持utf-8。

csv文件格式
---
我们仅支持excel的windows版本默认导出的csv格式。

格式以**","**作为列的分隔，**"\r\n"**作为行的分隔。

如果内容中存在**","**，这段内容会被引号引用起来。

json文件格式
---
json格式紧密没有做过多的格式化，但一般来说，json格式会比csv大很多。


安装
---

```
npm install xlsx2csv -g
```

使用&命令行参数
---
**xlsx2csv** 和 **xlsx2json** 用法几乎一样，就是输出文件格式不同。

* 基本用法 - 传入目录，即可将目录下所有xlsx或xls文件转换为csv格式。

```
xlsx2csv [path]
```

> 注意：默认情况下，csv的文件名和xlsx文件名一致，且只导出excel文件的第一张表。

* 遍历子目录 - 支持遍历子目录。

```
xlsx2csv **/win32/*.*
```

> 注意：即便是遍历 ```*.*``` ，最后也只会处理 ```*.xlsx``` 和 ```*.xls``` 2种扩展名的文件。

* 排除某一行 - 配合 [ds-lang](https://github.com/zhs007/ds-lang) 使用时，由于 [ds-lang](https://github.com/zhs007/ds-lang) 默认定义第一行为注释行，所以我们会需要文件转换时排除第一行，可以这样使用。

```
xlsx2csv [path] -e=0
```

* 导出多表 - excel文件可能会有多张表，我们可以根据表名来导出多张表。

```
xlsx2csv [path] -t
```

> 注意：这种情况下，会建一层xlsx文件名的目录，且csv的文件名和xlsx表名一致。


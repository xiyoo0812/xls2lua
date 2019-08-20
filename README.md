# xls2lua

## excel文件转换成lua的工具

+ 1 支持xls/xlsx格式的文档导出到lua。

+ 2、用法：xls2lua.exe xls_dir [lua_dir]。

```bat
	::xls_dir是excel文件所在位置，lua_dir为lua文件导出位置
	::lua路径不传则导出到xls_dir
	xls2lua.exe xls_dir lua_dir
```

+ 3、Excel配置表第一行为字段名，第二行为字段类型（string为字符串，不填默认数字），第三行为字段解析，从第四行开始为数据。

|  id   |  name  |  desc |
|  ----  | ----  |  ----  |
|    | string | string |
| ID | 名字 | 描述 |
| 10001  | 测试 | 测试数据 |

+ 4、字段解析不导出，字段名为空该列不导出。

+ 5、如果提示缺少驱动程序，请下载安装AccessDatabaseEngine。

+ 6、如果希望导出{[key] = {...}}格式lua表，则第一列字段命名为id。
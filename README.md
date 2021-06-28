# OfficeVbaTool

office 文档vba脚本插入工具

## 关于此项目
这算是一个老早之前的需求了，需要对word文件与excel文件实现一键插入vba宏的效果（用来干啥自己脑补）。
前前后后翻了不少资料，由于office还是微软自家的无奈只能用c#编写（老子c#一窍不通）。最后在厚着脸皮在
学长的帮助下对着一段不知道从哪个论坛里拾来的代码一顿瞎改，算是跑了起了。  
顺便吐槽下visual Studio的git不好用。

## 构建
构建环境为Visual Studio 2019, Office 2019, 其余环境不能保证构建与运行成功。  
如果是相近版本的Office可以尝试修改调用的com与对应的接口调用。

## 运行
当前只支持Word类型与Excel类型的文件  
Wrod类型支持输入格式为 ".doc", ".docm", ".docx"  
支持输出格式为 ".doc", ".docm"  
Excel类型支持输入格式为 ".xls", ".xlsm", ".xlsx"  
支持输出格式为 ".xlsm", ".xls"
```
OfficeVbaTool.exe src_path dst_path vba_script_path
```

## 源码与PE文件下载路径
[OfficeVbaTool](https://github.com/zn-chen/OfficeVbaTool/releases/tag/v1.0.0-rc)

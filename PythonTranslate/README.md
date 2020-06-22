翻译工具:
支持lua, xml文件的翻译,注释类的内容将会被忽略
支持中文直接翻译成繁体

Env:
需要python2.7(注意不支持python3.3),在运行python install.py安装用到的相关库（xlrd、xlwt和xlutils）

翻译lua、xml步骤:
1.打开config.ini修改SourcePath（需要翻译的文件夹路径）
2.执行python setup.py -e  提取中文并在output文件夹生成excel格式的语言包,将其交给翻译人员
3.翻译人员翻译好后复制粘贴单元格内容到translate.xls(必须跟config.ini中的LangPackageName值相同)
4.执行python setup.py -t  根据语言包翻译SourcePath文件夹下的文件
5.end!

翻译excel步骤：
1.打开config.ini修改ExcelPath（需要翻译的文件夹路径）
2.执行python setup.py -a 根据语言包翻译ExcelPath文件夹下的文件，并导出未翻译的
3.翻译人员翻译好后复制粘贴单元格内容到translate.xls(必须跟config.ini中的LangPackageName值相同)，然后重复2直到全部翻译完

修改翻译的步骤看根目录的“修改越南文说明.doc”

Other:
1.翻译excel的时候注意别打开需要翻译的excel表


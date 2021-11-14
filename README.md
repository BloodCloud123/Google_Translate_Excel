# Google_Translate_Excel
自用谷歌翻译Excel表格python脚本

## 食用方法
需要安装以下包:openpyxl、pygoogletranslation和tqdm。
修改pygoogletranslation源码中的utils.py的第8行为
(源码在.\python\Lib\site-packages\pygoogletranslation文件夹中)
```Python
from pygoogletranslation.models import TranslatedPart
```

然后在命令行输入以下命令后按提示操作即可
```cmd
python Excel_Google.py -f <Execl文件位置> -s <表名>
```

也可以输入以下命令查看帮助
```cmd
python Excel_Google.py -h
```
最后Enjoy it!
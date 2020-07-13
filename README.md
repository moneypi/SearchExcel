# SearchExcel
# 安装依赖
代码依赖pyqt5和xlrd  
```bash
pip install PyQt5
pip install xlrd
```
# 基本功能
选择文件夹后，输入搜索内容会自动出现搜索结果，无需点击搜索按钮
# 已知bug
无法搜索日期，如日期内容为"2018/04/01"时，因为没有正确进行日期格式转换，所以搜不到结果
# 未实现功能
没有实现子文件夹迭代搜索
# 参考链接
https://blog.csdn.net/zhqh100/article/details/107315520

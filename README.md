#PPT样机--python3
主要用于生成PPT详情模板图片，好几个月前就写好，之前和朋友一起在弄了一个素材类网站，当时有一个问题就是要把PPT是每一张幻灯片放到一个很长的模板图片里，用于放在网站页面展示。有上千的PPT文档，当时学python也没多久，就试着去用python的**win32com**和**wx**库，做了这个样机批量生成软件。经过多次调试修改，最终成型。由于当时是新手，还有很多不足的地方，勿喷。
===================================
#介绍
由python 3.5编写。
**需要安装的pythony库**
import wx
import win32com.client
from PIL import Image

**运行环境**
win7,win8,win10
需要安装office 2007及以上版本

#项目文本介绍
show.py  py运行文件
st.ico   LOGO
m.jpg    生成的模版样式
c.txt    一些配置与模版配合(可以用ConfigParser的，当时我不知道有这个)

最后我是用的pyinstaller库把文件打包的


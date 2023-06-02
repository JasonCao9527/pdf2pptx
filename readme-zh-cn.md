# pdf2pptx-exe
[English version](readme.md)|中文版本

本项目<mark>主要是用于本人存档和使用</mark>，其次就是实践一下，使用```pyinstaller```将python脚本打包成可执行的exe文件。

:warning:本项目需要在有```python3```的环境中运行。

另外，已经有很多人针对pdf2pptx这个做出了很多有意思的项目。比如kevinmcguinness把pdf2pptx做成了python的可调用的软件包[[pdf2pptx]](https://github.com/kevinmcguinness/pdf2pptx),
如果想要更方便快捷可以查看这个软件包。

根据phasedOut在[[pdf2pptx项目]](https://github.com/phasedOut/pdf2pptx)的源代码进行修改，
使其更加适合把使用```LaTeX Beamer```制作的演示文稿转化成pptx文件进行播放。

本文件是基于phasedOut的源代码进行二次开发，并将其封装为可执行的exe格式，
更加方便文件数量少的情况下的使用。

其原理就是把PDF的每一页转换成图片，然后逐张图片作为背景图插入ppt的每一页中。这样配合```LaTeX Beamer```
的```alert```或者```\pause```就可以伪装出一些播放的效果。

<mark>本项目无Release</mark>，如有需要使用，可自行clone

## 适用于
* 将使用```LaTeX Beamer```制作的演示文稿转化成pptx文件进行播放
* 或者其他每页16:9的PDF文件转化成pptx文件(其他比例未测试)


## 如何使用
进入「dist」文件夹后：
* 首先先点「setup.bat」创建三个文件夹
* 然后把<mark>需要转换的PDF文件放在「source_files」文件夹</mark>里面
* 再点击「pdf2pptx.exe」
* 等待转换完成， 完成后文件会存放在「result」文件夹中。

_*注：<mark>不要删除「default.pptx」</mark>_，这个是模板文件，
如果删除这个文件，转换脚本将无法执行。

## 删改说明
针对原来的项目：
* 删除了使得生成的pptx文件中插入生成时间的代码```dt_name = datetime.datetime.now()```，让功能更加纯粹
* 调整了原来的```slide.shapes.add_picture```插入的位置，现在不会「跑偏」
* 调整了输出图片的```每英寸像素(PPI)```
* 解决了在```Python 3.10```及更高版本中，由于```MutableMapping类```被删除，造成的报错
* 解决了源码封装成exe文件之后其使用的```pptx库```的报错问题

[//]: # ( :warning: )

[//]: # (<br>Please note that the English translation is done by Google Translate, please refer to the Chinese version.)

[//]: # (<br>请注意，英文翻译是由谷歌翻译进行的，请以中文版为准。)

**本项目参考：**
***
* [GitHub-phasedOut/pdf2pptx](https://github.com/phasedOut/pdf2pptx)
* [CSDN-python pptx库打包EXE报错pptx.exc.PackageNotFoundError解决办法](https://blog.csdn.net/weixin_54693379/article/details/128072858)
* [python-pptx 0.6.21 documentation](https://python-pptx.readthedocs.io/en/latest/api/shapes.html)

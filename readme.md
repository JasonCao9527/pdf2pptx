# pdf2pptx-exe
English version|[中文版本](readme-zh-cn.md)

This project <mark>is mainly used for personal archiving and use</mark>, followed by practice, using ```pyinstaller``` to package python scripts into executable exe files.

:warning: This project needs to run in an environment with ```python3```.

In addition, many people have made many interesting projects for pdf2pptx. For example, kevinmcguinness made pdf2pptx a python callable software package [[pdf2pptx]](https://github.com/kevinmcguinness/pdf2pptx),
If you want to be more convenient and fast, you can check out this package.

Modify according to the source code of phasedOut in [[pdf2pptx project]](https://github.com/phasedOut/pdf2pptx),
It is more suitable for converting presentations made with ```LaTeX Beamer``` into pptx files for playback.

This file is based on the source code of phasedOut for secondary development, and encapsulates it into an executable exe format.
It is more convenient to use when the number of files is small.

The principle is to convert each page of the PDF into a picture, and then insert each picture as a background image into each page of the ppt. This works with ```LaTeX Beamer```
The ```alert``` or ```\pause``` can fake some playback effects.

<mark>There is no Release for this project</mark>, if you need to use it, you can clone it yourself

## apply to
* Convert presentations made using ```LaTeX Beamer``` into pptx files for playback
* Or convert other 16:9 PDF files per page into pptx files (other ratios have not been tested)


## how to use
After entering the "dist" folder:
* First click "setup.bat" to create three folders
* Then put <mark>the PDF file to be converted into the "source_files" folder</mark>
* Then click "pdf2pptx.exe"
* Wait for the conversion to complete, and the file will be stored in the "result" folder after completion.

_*Note: <mark>Do not delete "default.pptx"</mark>_, this is a template file,
If this file is deleted, the conversion script will not execute.

## Description
For the original project:
* Deleted the code ```dt_name = datetime.datetime.now()``` that inserts the generated time into the generated pptx file, making the function more pure
* Adjusted the insertion position of the original ```slide.shapes.add_picture```, now it will not "deviate"
* Adjusted ```pixels per inch (PPI)``` of the output image
* Solved the error reported due to the deletion of the ```MutableMapping class``` in ```Python 3.10``` and later versions
* Solved the error reporting problem of the ```pptx library``` used after the source code is packaged into an exe file


:warning: 
<br>Please note that the English translation is done by Google Translate, please refer to the Chinese version.
<br>请注意，英文翻译是由谷歌翻译进行的，请以中文版为准。

**References**
***
* [GitHub-phasedOut/pdf2pptx](https://github.com/phasedOut/pdf2pptx)
* [CSDN-python pptx库打包EXE报错pptx.exc.PackageNotFoundError解决办法](https://blog.csdn.net/weixin_54693379/article/details/128072858)
* [python-pptx 0.6.21 documentation](https://python-pptx.readthedocs.io/en/latest/api/shapes.html)

本程序使用edge_tts来生成音频，主要目的是为英语学者将word文档表格中的单词生成成音频以供学习。word文档格式实列在main支里。Main支里的py文件为唯一的主程序，由1.ui和2.ui来提供图像用户界面。使用前需要特定的python库
pip install PyQt6
pip install edge-tts
pip install python-docx
pip install pymupdf
pip install pytesseract
pip install pillow
pip install audioread

pdf转文字功能需要pytesseract， 可能需要从https://github.com/tesseract-ocr/tesseract下载并安装，之后使用pip install pytesseract， 如果还是不行就使用pip install tesseract

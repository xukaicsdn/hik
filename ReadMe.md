1.安装python环境
2.pip install opencv-python
3.pip install pyautogui
4.pip install openpyxl
5.pip install pytesseract
6.安装了pytesseract后还需要下载安装Tesseract，下载路径https://github.com/UB-Mannheim/tesseract/wiki
  安装后路径选择再E:\Tesseract-OCR\tesseract.exe，如果安装路径不是E盘，需要再代码中修改默认的路径，如下：
    # 配置tesseract路径
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'E:\Tesseract-OCR\tesseract.exe'
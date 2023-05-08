import tempfile
import win32api
import win32print

# 打印临时文件(占用内存空间,可不生成文件)
def filePrint_TempFile():
    filename = tempfile.mktemp (".txt")
    open (filename, "w").write ("This is a test")
    win32api.ShellExecute (
        0,
        "print",
        filename,
        '/d:"%s"' % win32print.GetDefaultPrinter (),
        ".",
        0
    )

# 指定要打印的文件路径
def filePrint(file_path:str):
    # 获取默认打印机名称
    printer_name = win32print.GetDefaultPrinter()
    # 打印文件
    win32api.ShellExecute(
        0,
        "print",
        file_path,
        f'/d:"{printer_name}"',  # 指定打印机名称
        ".",
        0
    )

# 参考网站
# https://blog.csdn.net/qq_20259383/article/details/80592034
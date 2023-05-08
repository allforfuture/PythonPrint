import tempfile
import win32api
import win32print

# ��ӡ��ʱ�ļ�(ռ���ڴ�ռ�,�ɲ������ļ�)
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

# ָ��Ҫ��ӡ���ļ�·��
def filePrint(file_path:str):
    # ��ȡĬ�ϴ�ӡ������
    printer_name = win32print.GetDefaultPrinter()
    # ��ӡ�ļ�
    win32api.ShellExecute(
        0,
        "print",
        file_path,
        f'/d:"{printer_name}"',  # ָ����ӡ������
        ".",
        0
    )

# �ο���վ
# https://blog.csdn.net/qq_20259383/article/details/80592034
import sys
from schedule_app.gui import run
import ctypes

# Принудительно заставляем Windows распознавать приложение как уникальное на панели задач
if sys.platform == 'win32':
    myappid = 'myorganization.scheduleparser.promax.1.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


if __name__ == "__main__":
    sys.exit(run())

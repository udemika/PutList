import os
import sys

if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(os.path.abspath(sys.executable)))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

from gui import WaybillApp


def main():
    app = WaybillApp()
    app.mainloop()


if __name__ == "__main__":
    main()

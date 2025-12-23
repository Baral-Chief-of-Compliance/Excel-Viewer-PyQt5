import sys


from interface import MainWindow, make_app
from config import APP_WITDH, APP_HEIGHT, APP_TITEL, APP_VERSION


def main():
    app = make_app()

    window = MainWindow(
        '{} {}'.format(APP_TITEL, APP_VERSION),
        APP_WITDH,
        APP_HEIGHT
    )
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

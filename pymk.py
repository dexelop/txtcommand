from pymouse import PyMouseEvent
from pykeyboard import PyKeyboardEvent


class Mouseevent(PyMouseEvent):
    def __init__(self):
        super(Mouseevent, self).__init__()

    def move(self, x, y):
        print('event: click, x: {}, y: {}'.format(x, y))

    def click(self, x, y, button, press):
        self.stop()
        # return (x, y)


class Keyevent(PyKeyboardEvent):
    def __init__(self):
        super(Keyevent, self).__init__()

    def tap(self, keycode, character, press):
        print('event: tab, keycode: {}, character: {}'.format(keycode, character))


if __name__ == "__main__":
    k = Keyevent()
    m = Mouseevent()
    m.run()
    print("Capturing mouse")
    k.start()
    print("Capturing keyboard")
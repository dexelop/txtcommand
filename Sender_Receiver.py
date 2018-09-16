class Sender():
    def __init__(self):
        self.num = 3

    def devide2(self, n):
        self.result = n/2
        return self.result


class Receiver():
    def __init__(self):
        self.hi = 'hello'

    def print_num(self, res):
        print(self.hi + str(res))


s = Sender()
res = s.devide2(2)
# print(res)
r = Receiver()
r.print_num(res)
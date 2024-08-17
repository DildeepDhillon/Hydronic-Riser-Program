from schedule import Schedule

class Levels:

    def __init__(self, floor, riser, suite, schedule):
        self.floor = floor
        self.riser = riser
        self.suite = suite
        self.schedule = schedule
        self.bottomFlow = 0
        self.topFlow = 0
        self.topPipe = None
        self.bottomPipe = None




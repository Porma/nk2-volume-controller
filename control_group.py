class ControlGroup:
    program = ""

    def __init__(self, fader):
        self.fader = fader
        self.select = fader + 32
        self.mute = fader + 48



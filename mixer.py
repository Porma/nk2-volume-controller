import mido

from control_group import ControlGroup


class Mixer:
    # Used for main loop
    running = True
    # Dictionary to hold ControlGroup objects
    groups = {}

    # NanoKontrol input and output
    nk2_in = mido.open_input('nanoKONTROL2 1 SLIDER/KNOB 0')
    nk2_out = mido.open_output('nanoKONTROL2 1 CTRL 1')

    # Control number range of faders to use
    fader_range = range(0, 4)
    # Control number range of select buttons to use
    select_range = range(32, 36)
    # Control number range of mute buttons
    mute_range = range(48, 52)
    # Difference between fader and select control numbers
    select_fader_diff = 32
    # Difference between fader and mute control numbers
    mute_fader_diff = select_fader_diff + 16

    # Control number range of faders for Voicemeeter
    vb_fader_range = range(4, 8)
    # Control number range of mute buttons for Voicemeeter
    vb_mute_range = range(52, 56)

    def __init__(self):
        for i in self.fader_range:
            self.groups[i] = ControlGroup(i)

        self.reset_lights()

    # Turn on the light for a given control
    def enable_light(self, control):
        msg = mido.Message('control_change', control=control, value=127)
        self.nk2_out.send(msg)

    # Turn off the light for a given control
    def disable_light(self, control):
        msg = mido.Message('control_change', control=control, value=0)
        self.nk2_out.send(msg)

    # Turn off lights for all controls
    def reset_lights(self):
        for i in self.select_range:
            msg = mido.Message('control_change', control=i, value=0)
            self.nk2_out.send(msg)

        for i in self.mute_range:
            msg = mido.Message('control_change', control=i, value=0)
            self.nk2_out.send(msg)

import mido
import psutil
import win32process
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import win32gui

from control_group import ControlGroup

# Control number range of faders to use
CONST_FADER_RANGE = range(0, 4)
# Control number range of select buttons to use
CONST_SELECT_RANGE = range(32, 36)
# Control number range of mute buttons
CONST_MUTE_RANGE = range(48, 52)
# Difference between fader and select control numbers
CONST_SELECT_FADER_DIFF = 32
# Difference between fader and mute control numbers
CONST_MUTE_FADER_DIFF = CONST_SELECT_FADER_DIFF + 16


def main():
    nk2 = mido.open_input('nanoKONTROL2 1 SLIDER/KNOB 0')
    sessions = AudioUtilities.GetAllSessions

    # Dictionary to hold ControlGroup objects
    control_groups = {}

    # Populate dict using fader control number for key
    for i in CONST_FADER_RANGE:
        control_groups[i] = ControlGroup(i)

    # print(control_groups)

    reset_lights()

    for session in sessions():
        if session.Process:
            # print(session.Process)
            print(session.Process.name)

    while True:
        # Receive midi message (non-blocking)
        msg = nk2.poll()

        # If no message exists,
        if not msg:
            continue

        # select button pressed, on keyrelease
        if msg.control in CONST_SELECT_RANGE and msg.value == 0:
            group = control_groups[msg.control - CONST_SELECT_FADER_DIFF]

            # If program bound to group, unbind and turn off light
            if group.program != "":
                print(f"{group.program} unbound from fader {group.fader}")
                group.program = ""

                # Disable select button light
                disable_light(group.select)
                # Disable mute button light
                disable_light(group.mute)
            else:
                # Get active program to
                active_exe = get_active_program()
                # Assign active window to ControlGroup object in dict
                group.program = active_exe

                # Enable select button light
                enable_light(group.select)

                # Enable mute button light if program is muted
                if group.isMuted:
                    enable_light(group.mute)

                print(f"{active_exe} bound to fader {group.fader}")

        # Check for mute button press
        elif msg.control in CONST_MUTE_RANGE and msg.value == 0 and control_groups[msg.control - CONST_MUTE_FADER_DIFF].program:
            group = control_groups[msg.control - CONST_MUTE_FADER_DIFF]

            session = next(
                (s for s in sessions() if s.Process and s.Process.name() == group.program), None)

            if session.SimpleAudioVolume.GetMute():
                session.SimpleAudioVolume.SetMute(0, None)
                disable_light(group.mute)
                group.isMuted = False
                print(f"{group.program} unmuted (fader {group.fader})")
            else:
                session.SimpleAudioVolume.SetMute(1, None)
                enable_light(group.mute)
                group.isMuted = True
                print(f"{group.program} muted (fader {group.fader})")

        # Check if key matching a fader input exists
        elif msg.control in control_groups:
            # print(control_groups[msg.control].fader)
            # print(control_groups[msg.control].program)

            # Find audio session with matching program name of input fader control
            session = next(
                (s for s in sessions() if s.Process and s.Process.name() == control_groups[msg.control].program), None)

            # If no matching session was found i.e fader not assigned a program, end current loop
            if not session:
                continue

            # Get volume control object from session
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            # Convert midi value to percentage and set volume
            volume.SetMasterVolume(msg.value / 127, None)

            print(f"{control_groups[msg.control].program} set to {volume.GetMasterVolume() * 100}%")

        # After input is processed delete message to prevent unnecessary looping
        msg = None


# Return process name of active window in foreground
def get_active_program():
    active_window = win32gui.GetForegroundWindow()
    active_pid = win32process.GetWindowThreadProcessId(active_window)[1]
    return psutil.Process(active_pid).name()


# Turn on the light for a given control
def enable_light(control):
    with mido.open_output('nanoKONTROL2 1 CTRL 1') as outport:
        out_msg = mido.Message('control_change', control=control, value=127)
        outport.send(out_msg)


# Turn on the light for a given control
def disable_light(control):
    with mido.open_output('nanoKONTROL2 1 CTRL 1') as outport:
        out_msg = mido.Message('control_change', control=control, value=0)
        outport.send(out_msg)


# Turn off lights for all controls
def reset_lights():
    with mido.open_output('nanoKONTROL2 1 CTRL 1') as outport:
        for i in CONST_SELECT_RANGE:
            out_msg = mido.Message('control_change', control=i, value=0)
            outport.send(out_msg)

        for i in CONST_MUTE_RANGE:
            out_msg = mido.Message('control_change', control=i, value=0)
            outport.send(out_msg)


if __name__ == '__main__':
    main()

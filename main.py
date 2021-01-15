import mido
import psutil
import win32process
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import win32gui
import voicemeeter

from control_group import ControlGroup

# NOTE: Voicemeeter library error -
# Edit remote.py line 172
# from  def connect(kind_id, delay=null):
# to    def connect(kind_id, delay=None):

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

# Control number range of faders for Voicemeeter
CONST_VB_FADER_RANGE = range(4, 8)
# Control number range of mute buttons for Voicemeeter
CONST_VB_MUTE_RANGE = range(52, 56)


def main():
    nk2 = mido.open_input('nanoKONTROL2 1 SLIDER/KNOB 0')
    sessions = AudioUtilities.GetAllSessions

    # Map physical faders to Voicemeeter faders - mappings subject to personal preference
    with voicemeeter.remote('banana', 0.0005) as vmr:
        vb_map = {
            4: vmr.inputs[3],
            5: vmr.inputs[4],
            6: vmr.inputs[1],
            7: vmr.outputs[0]
        }

    # Dictionary to hold ControlGroup objects
    control_groups = {}

    # Populate dict using fader control number for key
    for i in CONST_FADER_RANGE:
        control_groups[i] = ControlGroup(i)

    reset_lights()

    for session in sessions():
        if session.Process:
            # print(session.Process)
            print(session.Process.name)

    while True:
        # Receive midi message (non-blocking)
        msg = nk2.poll()

        # If no message exists, end current loop
        if not msg:
            continue

        # Check for select button press
        if msg.control in CONST_SELECT_RANGE and msg.value:
            group = control_groups[msg.control - CONST_SELECT_FADER_DIFF]

            # If program is not bound bound
            if not group.program:
                # Get active program name
                active_program = get_active_program()
                # Find audio session with matching name
                session = next(
                    (s for s in sessions() if s.Process and s.Process.name() == active_program), None)
                # Assign session to control group
                group.program = session

                # Turn on select button light
                enable_light(group.select)

                # If program is muted turn on mute button light
                if group.program.SimpleAudioVolume.GetMute():
                    enable_light(group.mute)

                print(f"{group.program.Process.name()} bound to fader {group.fader}")

            else:
                print(f"{group.program.Process.name()} unbound from fader {group.fader}")

                # Unassign session from fader
                group.program = None

                # Turn off select button light
                disable_light(group.select)
                # Turn off mute button light
                disable_light(group.mute)

        # Check for mute button press
        elif msg.control in CONST_MUTE_RANGE and msg.value and control_groups[msg.control - CONST_MUTE_FADER_DIFF].program:
            group = control_groups[msg.control - CONST_MUTE_FADER_DIFF]

            # Check if program is muted
            if group.program.SimpleAudioVolume.GetMute():
                # Unmute program
                group.program.SimpleAudioVolume.SetMute(0, None)
                # Turn off mute button light
                disable_light(group.mute)

                print(f"{group.program.Process.name()} unmuted (fader {group.fader})")

            # If program is not muted
            else:
                # Mute program
                group.program.SimpleAudioVolume.SetMute(1, None)
                # Turn on mute button light
                enable_light(group.mute)

                print(f"{group.program.Process.name()} muted (fader {group.fader})")

        # Check for fader input
        elif msg.control in CONST_FADER_RANGE:
            group = control_groups[msg.control]

            # If fader does not have assigned program end current loop
            if not group.program:
                continue

            # Get volume control object from session
            volume = group.program._ctl.QueryInterface(ISimpleAudioVolume)
            # Convert midi value to percentage and set volume
            volume.SetMasterVolume(msg.value / 127, None)

            print(f"{group.program.Process.name()} set to {volume.GetMasterVolume() * 100}%")

        # Check for Voicemeeter fader input
        elif msg.control in CONST_VB_FADER_RANGE:
            # Map midi value (0-127) to VB appropriate gain value (-60-0)
            level = ((127 - msg.value) / 127) * -60
            # Set VB fader gain
            vb_map[msg.control].gain = level

            print(f"fader {msg.control} (VoiceMeeter) gain set to {level}")

        elif msg.control in CONST_VB_MUTE_RANGE and msg.value:
            fader = msg.control - CONST_MUTE_FADER_DIFF
            control = vb_map[fader]

            # ISSUE: inconsistent mute/unmute

            if control.mute:
                # Unmute VB control
                control.mute = False
                # Turn off mute button light
                disable_light(msg.control)

                print(f"fader {fader} (VoiceMeeter) unmuted")
            else:
                # Mute FB control
                control.mute = True
                # Turn on mute button light
                enable_light(msg.control)

                print(f"fader {fader} (VoiceMeeter) muted")

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

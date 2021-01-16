import mido
import psutil
import win32process
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import win32gui
import voicemeeter

from control_group import ControlGroup
from mixer import Mixer


def main():
    sessions = AudioUtilities.GetAllSessions
    mixer = Mixer()

    # Map physical faders to Voicemeeter faders - mappings subject to personal preference
    with voicemeeter.remote('banana', 0.0005) as vmr:
        mixer.vb_map = {
            4: vmr.inputs[3],
            5: vmr.inputs[4],
            6: vmr.inputs[1],
            7: vmr.outputs[0]
        }

    # Populate dict using fader control number for key
    for i in mixer.fader_range:
        mixer.groups[i] = ControlGroup(i)

    for session in sessions():
        if session.Process:
            print(session.Process.name)

    while mixer.running:
        # Receive midi message (non-blocking)
        msg = mixer.nk2_in.poll()

        # If no message exists, end current loop
        if not msg:
            continue

        # Check for select button press
        if msg.control in mixer.select_range and msg.value:
            group = mixer.groups[msg.control - mixer.select_fader_diff]

            # If program is not bound bound
            if not group.program:
                # Get active program name
                active_program = get_active_program()
                # Find audio session with matching name
                session = next(
                    (s for s in sessions() if s.Process and s.Process.name() == active_program), None)

                # If audio session does not exist end current loop
                if not session:
                    continue

                # Assign session to control group
                group.program = session

                # Turn on select button light
                mixer.enable_light(group.select)

                # If program is muted turn on mute button light
                if group.program.SimpleAudioVolume.GetMute():
                    mixer.enable_light(group.mute)

                print(f"{group.program.Process.name()} bound to fader {group.fader}")

            else:
                print(f"{group.program.Process.name()} unbound from fader {group.fader}")

                # Unassign session from fader
                group.program = None

                # Turn off select button light
                mixer.disable_light(group.select)
                # Turn off mute button light
                mixer.disable_light(group.mute)

        # Check for mute button press
        elif msg.control in mixer.mute_range and msg.value and mixer.groups[msg.control - mixer.mute_fader_diff].program:
            group = mixer.groups[msg.control - mixer.mute_fader_diff]

            # Check if program is muted
            if group.program.SimpleAudioVolume.GetMute():
                # Unmute program
                group.program.SimpleAudioVolume.SetMute(0, None)
                # Turn off mute button light
                mixer.disable_light(group.mute)

                print(f"{group.program.Process.name()} unmuted (fader {group.fader})")

            # If program is not muted
            else:
                # Mute program
                group.program.SimpleAudioVolume.SetMute(1, None)
                # Turn on mute button light
                mixer.enable_light(group.mute)

                print(f"{group.program.Process.name()} muted (fader {group.fader})")

        # Check for fader input
        elif msg.control in mixer.fader_range:
            group = mixer.groups[msg.control]

            # If fader does not have assigned program end current loop
            if not group.program:
                continue

            # Get volume control object from session
            volume = group.program._ctl.QueryInterface(ISimpleAudioVolume)
            # Convert midi value to percentage and set volume
            volume.SetMasterVolume(msg.value / 127, None)

            print(f"{group.program.Process.name()} set to {volume.GetMasterVolume() * 100}%")

        # Check for Voicemeeter fader input
        elif msg.control in mixer.vb_fader_range:
            # Map midi value (0-127) to VB appropriate gain value (-60-0)
            level = ((127 - msg.value) / 127) * -60
            # Set VB fader gain
            mixer.vb_map[msg.control].gain = level

            print(f"fader {msg.control} (VoiceMeeter) gain set to {level}")

        elif msg.control in mixer.vb_mute_range and msg.value:
            fader = msg.control - mixer.mute_fader_diff
            control = mixer.vb_map[fader]

            # ISSUE: inconsistent mute/unmute

            if control.mute:
                # Unmute VB control
                control.mute = False
                # Turn off mute button light
                mixer.disable_light(msg.control)

                print(f"fader {fader} (VoiceMeeter) unmuted")
            else:
                # Mute FB control
                control.mute = True
                # Turn on mute button light
                mixer.enable_light(msg.control)

                print(f"fader {fader} (VoiceMeeter) muted")

        # After input is processed delete message to prevent unnecessary looping
        msg = None


# Return process name of active window in foreground
def get_active_program():
    active_window = win32gui.GetForegroundWindow()
    active_pid = win32process.GetWindowThreadProcessId(active_window)[1]
    return psutil.Process(active_pid).name()


if __name__ == '__main__':
    main()

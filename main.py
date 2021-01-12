import mido
import psutil
import win32process
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import win32gui


def main():
    nk2 = mido.open_input('nanoKONTROL2 1 SLIDER/KNOB 0')
    sessions = AudioUtilities.GetAllSessions()
    bindings = {}

    reset_lights()

    for session in sessions:
        if session.Process:
            print(session.Process)
        # print(session.Process.name)

    while True:
        # active_window = win32gui.GetForegroundWindow()
        # active_pid = win32process.GetWindowThreadProcessId(active_window)[1]
        # psutil.Process(active_pid).name()

        # get midi input
        msg = nk2.poll()

        # don't go any further until a message is received
        if not msg:
            continue

        # msg exists, select button pressed, on keyrelease
        if msg and 32 <= msg.control <= 35 and msg.value == 0:
            active_window = win32gui.GetForegroundWindow()
            active_pid = win32process.GetWindowThreadProcessId(active_window)[1]
            active_exe = psutil.Process(active_pid).name()

            # print(active_exe)
            # print(msg.value)

            # get fader number from temp fn
            fader = get_fader(msg.control)

            # assign active exe name to fader
            bindings[fader] = active_exe

            with mido.open_output('nanoKONTROL2 1 CTRL 1') as outport:
                out_msg = mido.Message('control_change', control=msg.control, value=127)
                outport.send(out_msg)

            print(f"{active_exe} bound to fader {fader}")
        # check for bound fader input
        elif msg.control in bindings:
            # loop through audio sessions to find bound program name
            for session in sessions:
                if session.Process and session.Process.name() == bindings[msg.control]:
                    # get volume control object
                    volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                    # convert midi input to percentage and set volume
                    volume.SetMasterVolume(msg.value / 127, None)

                    print(f"{bindings[msg.control]} set to {volume.GetMasterVolume()}%")


# temp fn - get fader control number from select
def get_fader(select):
    return select - 32


def reset_lights():
    with mido.open_output('nanoKONTROL2 1 CTRL 1') as outport:
        for i in range(32, 36):
            out_msg = mido.Message('control_change', control=i, value=0)
            outport.send(out_msg)

if __name__ == '__main__':
    main()

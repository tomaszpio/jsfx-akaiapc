import pytest


NOTE_ON = 0x90
NOTE_OFF = 0x80
CC = 0xB0
SUSTAIN_CC = 64


class MidiSustainHoldSimulator:
    """Minimal simulator mirroring midi_sustain_hold.jsfx behaviour."""

    def __init__(self, channel=-1):
        self.channel = channel  # -1 = omni, otherwise 0-15
        self.pedal_down = False
        self.held = [False] * 128
        self.held_chan = [0] * 128

    def _in_channel(self, ch: int) -> bool:
        return self.channel < 0 or ch == self.channel

    def process(self, events):
        """Yield output events after processing the given inputs.

        Each event is a tuple (status, data1, data2).
        """

        outputs = []
        for status, d1, d2 in events:
            ch = status & 0x0F
            msg_type = status & 0xF0

            if msg_type == CC and d1 == SUSTAIN_CC:
                outputs.append((status, d1, d2))
                if d2 == 0 and self.pedal_down:
                    for note, held in enumerate(self.held):
                        if held:
                            outputs.append((NOTE_OFF + self.held_chan[note], note, 0))
                            self.held[note] = False
                self.pedal_down = d2 > 0
                continue

            # Note handling for the selected channel
            if msg_type in (NOTE_ON, NOTE_OFF) and self._in_channel(ch):
                if msg_type == NOTE_ON and d2 > 0:
                    outputs.append((status, d1, d2))
                    self.held[d1] = False
                else:
                    if self.pedal_down:
                        self.held[d1] = True
                        self.held_chan[d1] = ch
                    else:
                        outputs.append((status, d1, d2))
                continue

            # Other messages or channels pass through unchanged
            outputs.append((status, d1, d2))
        return outputs


def format_events(events):
    return [f"{status:02X}:{d1}:{d2}" for status, d1, d2 in events]


def test_note_off_blocked_until_sustain_release():
    sim = MidiSustainHoldSimulator()
    events = [
        (NOTE_ON, 60, 100),       # Note On
        (CC, SUSTAIN_CC, 127),    # Sustain On
        (NOTE_OFF, 60, 0),        # Note Off should be blocked
        (CC, SUSTAIN_CC, 0),      # Sustain Off triggers release
    ]

    output = sim.process(events)

    assert format_events(output) == [
        "90:60:100",  # forwarded note on
        "B0:64:127",  # forwarded sustain on
        "B0:64:0",    # forwarded sustain off
        "80:60:0",    # delayed note off
    ]


def test_multiple_notes_released_on_pedal_up():
    sim = MidiSustainHoldSimulator()
    events = [
        (NOTE_ON, 60, 90),
        (CC, SUSTAIN_CC, 100),
        (NOTE_ON, 62, 100),
        (NOTE_OFF, 60, 0),
        (NOTE_OFF, 62, 0),
        (CC, SUSTAIN_CC, 0),
    ]

    output = sim.process(events)

    assert format_events(output) == [
        "90:60:90",
        "B0:64:100",
        "90:62:100",
        "B0:64:0",
        "80:60:0",
        "80:62:0",
    ]

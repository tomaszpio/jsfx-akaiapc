import pytest


NOTE_ON = 0x90
NOTE_OFF = 0x80
CC = 0xB0
SUSTAIN_CC = 64


class MidiSustainFilterSimulator:
    """Mirror midi_sustain_filter.jsfx sustain behaviour."""

    def __init__(self):
        self.pedal_down = False
        self.held = [False] * 128
        self.held_chan = [0] * 128
        self.active = [False] * 128
        self.active_chan = [0] * 128

    def process(self, events):
        """Yield output events after processing the given inputs.

        Each event is (status, data1, data2).
        """

        outputs = []
        for status, d1, d2 in events:
            ch = status & 0x0F
            msg_type = status & 0xF0

            # Sustain pedal
            if msg_type == CC and d1 == SUSTAIN_CC:
                outputs.append((status, d1, d2))
                sustain_on = d2 > 0
                # Capture currently sounding notes when pedal goes down
                if (not self.pedal_down) and sustain_on:
                    for note, active in enumerate(self.active):
                        if active:
                            self.held[note] = True
                            self.held_chan[note] = self.active_chan[note]

                if self.pedal_down and not sustain_on:
                    for note, held in enumerate(self.held):
                        if held:
                            outputs.append((NOTE_OFF + self.held_chan[note], note, 0))
                            self.held[note] = False
                            self.active[note] = False
                            self.active_chan[note] = 0
                self.pedal_down = sustain_on
                continue

            # Note handling (all channels)
            if msg_type in (NOTE_ON, NOTE_OFF):
                if msg_type == NOTE_ON and d2 > 0:
                    outputs.append((status, d1, d2))
                    self.active[d1] = True
                    self.active_chan[d1] = ch
                    if self.pedal_down:
                        self.held[d1] = True
                        self.held_chan[d1] = ch
                    else:
                        self.held[d1] = False
                else:
                    self.active[d1] = False
                    if self.pedal_down:
                        self.held[d1] = True
                        self.held_chan[d1] = ch
                    else:
                        outputs.append((status, d1, d2))
                continue

            # Other MIDI messages pass through unchanged
            outputs.append((status, d1, d2))

        return outputs


def format_events(events):
    return [f"{status:02X}:{d1}:{d2}" for status, d1, d2 in events]


def test_note_offs_flushed_when_pedal_value_hits_zero():
    sim = MidiSustainFilterSimulator()
    events = [
        (NOTE_ON, 60, 100),
        (CC, SUSTAIN_CC, 100),
        (NOTE_OFF, 60, 0),
        (CC, SUSTAIN_CC, 0),
    ]

    output = sim.process(events)

    assert format_events(output) == [
        "90:60:100",
        "B0:64:100",
        "B0:64:0",
        "80:60:0",
    ]


def test_pedal_release_occurs_only_at_zero():
    sim = MidiSustainFilterSimulator()
    events = [
        (NOTE_ON, 62, 90),
        (CC, SUSTAIN_CC, 90),
        (NOTE_OFF, 62, 0),
        (CC, SUSTAIN_CC, 50),  # still down (no flush)
        (CC, SUSTAIN_CC, 0),   # pedal up triggers release
    ]

    output = sim.process(events)

    assert format_events(output) == [
        "90:62:90",
        "B0:64:90",
        "B0:64:50",
        "B0:64:0",
        "80:62:0",
    ]


def test_note_on_while_pedal_down_releases_on_pedal_up():
    sim = MidiSustainFilterSimulator()
    events = [
        (CC, SUSTAIN_CC, 127),   # pedal down first
        (NOTE_ON, 65, 110),      # note on while pedal down
        (CC, SUSTAIN_CC, 0),     # pedal up should send note off
    ]

    output = sim.process(events)

    assert format_events(output) == [
        "B0:64:127",
        "90:65:110",
        "B0:64:0",
        "80:65:0",
    ]


def test_note_on_before_pedal_down_releases_on_pedal_up():
    sim = MidiSustainFilterSimulator()
    events = [
        (NOTE_ON, 55, 92),       # note on before pedal engages
        (CC, SUSTAIN_CC, 10),    # pedal partially down (engaged)
        (CC, SUSTAIN_CC, 127),   # pedal fully down
        (CC, SUSTAIN_CC, 0),     # pedal up should release captured note
    ]

    output = sim.process(events)

    assert format_events(output) == [
        "90:55:92",
        "B0:64:10",
        "B0:64:127",
        "B0:64:0",
        "80:55:0",
    ]


def test_ramping_cc_values_still_release_on_zero():
    sim = MidiSustainFilterSimulator()
    events = [
        (NOTE_ON, 55, 92),
        (CC, SUSTAIN_CC, 10),
        (CC, SUSTAIN_CC, 20),
        (CC, SUSTAIN_CC, 32),
        (CC, SUSTAIN_CC, 42),
        (CC, SUSTAIN_CC, 52),
        (CC, SUSTAIN_CC, 64),
        (CC, SUSTAIN_CC, 72),
        (CC, SUSTAIN_CC, 82),
        (CC, SUSTAIN_CC, 94),
        (CC, SUSTAIN_CC, 104),
        (CC, SUSTAIN_CC, 114),
        (CC, SUSTAIN_CC, 126),
        (CC, SUSTAIN_CC, 127),
        (CC, SUSTAIN_CC, 118),
        (CC, SUSTAIN_CC, 108),
        (CC, SUSTAIN_CC, 96),
        (CC, SUSTAIN_CC, 84),
        (CC, SUSTAIN_CC, 70),
        (CC, SUSTAIN_CC, 60),
        (CC, SUSTAIN_CC, 46),
        (CC, SUSTAIN_CC, 34),
        (CC, SUSTAIN_CC, 22),
        (CC, SUSTAIN_CC, 8),
        (CC, SUSTAIN_CC, 0),
    ]

    output = sim.process(events)

    assert format_events(output) == [
        "90:55:92",
        "B0:64:10",
        "B0:64:20",
        "B0:64:32",
        "B0:64:42",
        "B0:64:52",
        "B0:64:64",
        "B0:64:72",
        "B0:64:82",
        "B0:64:94",
        "B0:64:104",
        "B0:64:114",
        "B0:64:126",
        "B0:64:127",
        "B0:64:118",
        "B0:64:108",
        "B0:64:96",
        "B0:64:84",
        "B0:64:70",
        "B0:64:60",
        "B0:64:46",
        "B0:64:34",
        "B0:64:22",
        "B0:64:8",
        "B0:64:0",
        "80:55:0",
    ]

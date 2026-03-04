#!/usr/bin/env python3
"""
REAPER MIDI CC Editor - GUI Application Version 1.1
- View & edit MIDI CC assignments in REAPER .RPP files
- Multi-row selection for bulk CC / Channel / Bus reassignment
- Routing visualisation: node graph, matrix, text list (separate window)

AUXRECV format per official CockosWiki / ReaTeam/Doc:
  AUXRECV src_idx mode vol pan mute mono_sum phase src_ach dst_ach panlaw midi_ch auto_mode
  field 1:  int   - source track index (0-based)
  field 2:  int   - mode: 0=Post Fader/Post Pan, 1=Pre FX, 3=Pre Fader/Post FX
  field 3:  float - volume
  field 4:  float - pan
  field 5:  int   - mute (bool)
  field 6:  int   - mono sum (bool)
  field 7:  int   - invert phase (bool)
  field 8:  int   - source audio channels: -1=none, 0=1+2, 1=2+3, 2=3+4 ...
  field 9:  int   - dest audio channels (same encoding, no -1)
  field 10: float - pan law
  field 11: int   - MIDI channel mapping:
                      0 = no MIDI send
                      source channel = val & 0x1F   (1-16 = ch1-16, 17 = all)
                      dest channel   = floor(val/32) (1-16 = ch1-16, 0 = original)
  field 12: int   - automation mode (-1 = use track mode)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
import math
from typing import List, Dict, Tuple, Optional
import os


# ═══════════════════════════════════════════════════════════════════════════
#  AUXRECV MIDI channel decoding
# ═══════════════════════════════════════════════════════════════════════════

def decode_midi_ch_field(val: int):
    """
    Decode AUXRECV field 11 (MIDI channel mapping).
    Returns (has_midi, src_ch_label, dst_ch_label) or (False, None, None).

    Encoding (from CockosWiki):
      0          = no MIDI send at all
      src  = val & 0x1F         1-16 = specific ch, 17 = all channels
      dst  = floor(val / 32)    1-16 = specific ch,  0 = original (pass-through)
    """
    if val == 0:
        return False, None, None

    src_raw = val & 0x1F
    dst_raw = val >> 5          # same as floor(val/32)

    src_label = 'All' if src_raw == 17 else (f'Ch {src_raw - 1}' if src_raw else '?')
    dst_label = 'Original' if dst_raw == 0 else (f'Ch {dst_raw - 1}')

    return True, src_label, dst_label


def decode_audio_ch(val: int) -> str:
    """Convert AUXRECV audio channel field to human label."""
    if val == -1:
        return 'None'
    ch = val + 1
    return f'{ch}/{ch+1}'


FADER_MODES = {0: 'Post Fader', 1: 'Pre FX', 3: 'Pre Fader'}


# ═══════════════════════════════════════════════════════════════════════════
#  Data model
# ═══════════════════════════════════════════════════════════════════════════

class REAPERProject:
    def __init__(self):
        self.filepath = None
        self.lines: List[str] = []
        self.tracks: List[dict] = []
        self.modified = False

    def load_file(self, filepath: str):
        self.filepath = filepath
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            self.lines = f.readlines()
        self._parse_structure()
        self.modified = False

    def _parse_structure(self):
        self.tracks = []
        current_track = None
        current_fx = None
        current_programenv = None
        in_fxchain = False
        fxchain_depth = 0

        for i, line in enumerate(self.lines):
            indent = len(line) - len(line.lstrip())
            stripped = line.strip()

            if stripped.startswith('<TRACK '):
                guid_match = re.search(r'\{([^}]+)\}', stripped)
                current_track = {
                    'guid': guid_match.group(1) if guid_match else 'Unknown',
                    'name': None,
                    'fx_list': [],
                    'receives': [],   # parsed AUXRECV data
                    'sends': [],      # derived in post-pass
                    'folder_depth': 0,
                    'line_num': i,
                }
                self.tracks.append(current_track)
                in_fxchain = False
                current_fx = None
                current_programenv = None

            elif current_track:
                if stripped.startswith('NAME '):
                    name = stripped[5:].strip('"')
                    current_track['name'] = name if name else None

                elif stripped.startswith('ISBUS '):
                    parts = stripped.split()
                    if len(parts) >= 2:
                        try:
                            current_track['folder_depth'] = int(parts[1])
                        except ValueError:
                            pass

                elif stripped.startswith('AUXRECV '):
                    # AUXRECV src mode vol pan mute mono phase src_ach dst_ach panlaw midi_ch auto
                    #  [0]    [1] [2]  [3] [4] [5]  [6]  [7]   [8]    [9]    [10]   [11]   [12]
                    parts = stripped.split()
                    try:
                        src_idx    = int(parts[1])
                        fader_mode = int(parts[2])   if len(parts) > 2  else 0
                        src_ach    = int(parts[8])   if len(parts) > 8  else 0
                        dst_ach    = int(parts[9])   if len(parts) > 9  else 0
                        midi_field = int(parts[11])  if len(parts) > 11 else 0

                        has_audio = src_ach != -1
                        has_midi, midi_src_ch, midi_dst_ch = decode_midi_ch_field(midi_field)

                        current_track['receives'].append({
                            'src_idx':     src_idx,
                            'fader_mode':  fader_mode,
                            'has_audio':   has_audio,
                            'src_ach':     src_ach,
                            'dst_ach':     dst_ach,
                            'has_midi':    has_midi,
                            'midi_src_ch': midi_src_ch,   # e.g. 'Ch 3', 'All', or None
                            'midi_dst_ch': midi_dst_ch,   # e.g. 'Ch 1', 'Original', or None
                            'midi_raw':    midi_field,
                        })
                    except (ValueError, IndexError):
                        pass

                elif stripped.startswith('<FXCHAIN'):
                    in_fxchain = True
                    fxchain_depth = indent
                    current_fx = None
                    current_programenv = None

                elif in_fxchain:
                    fx_match = re.match(r'<(VST|AU|JS|VST3|CLAP)\s+"(.+?)"', stripped)
                    if fx_match:
                        current_fx = {
                            'type': fx_match.group(1),
                            'name': fx_match.group(2),
                            'line_num': i,
                            'modulations': [],
                        }
                        current_track['fx_list'].append(current_fx)
                        current_programenv = None

                    elif stripped.startswith('<PROGRAMENV '):
                        m = re.match(r'<PROGRAMENV\s+(\S+)\s+(\d+)\s+"([^"]+)"', stripped)
                        if m:
                            current_programenv = {
                                'param_id':       m.group(1),
                                'param_name':     m.group(3),
                                'bypass_flag':    int(m.group(2)),
                                'midi_cc':        None,
                                'midi_channel':   None,
                                'midi_bus':       None,
                                'midi_msg_type':  None,
                                'programenv_line': i,
                                'midiplink_line': None,
                            }
                            if current_fx:
                                current_fx['modulations'].append(current_programenv)
                            elif current_track['fx_list']:
                                current_track['fx_list'][-1]['modulations'].append(current_programenv)

                    elif stripped.startswith('MIDIPLINK ') and current_programenv:
                        mm = re.match(r'MIDIPLINK\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)', stripped)
                        if mm:
                            current_programenv['midi_bus']       = int(mm.group(1))
                            current_programenv['midi_channel']   = int(mm.group(2))
                            current_programenv['midi_msg_type']  = int(mm.group(3))
                            current_programenv['midiplink_line'] = i
                            cc_val = int(mm.group(4))
                            if int(mm.group(3)) == 176:
                                current_programenv['midi_cc'] = cc_val
                            else:
                                current_programenv['midi_note'] = cc_val

                    elif stripped == '>' and current_programenv:
                        current_programenv = None

                    if stripped == '>' and indent <= fxchain_depth:
                        in_fxchain = False
                        current_fx = None
                        current_programenv = None

        # Post-pass: derive sends list on source tracks from receive data on dest tracks
        for dst_idx, track in enumerate(self.tracks):
            for recv in track['receives']:
                src_idx = recv['src_idx']
                if 0 <= src_idx < len(self.tracks):
                    self.tracks[src_idx]['sends'].append({
                        'dst_idx':     dst_idx,
                        'fader_mode':  recv['fader_mode'],
                        'has_audio':   recv['has_audio'],
                        'src_ach':     recv['src_ach'],
                        'dst_ach':     recv['dst_ach'],
                        'has_midi':    recv['has_midi'],
                        'midi_src_ch': recv['midi_src_ch'],
                        'midi_dst_ch': recv['midi_dst_ch'],
                    })

    def update_midi_cc(self, track_idx, fx_idx, mod_idx,
                       new_cc, new_channel, new_bus) -> bool:
        try:
            mod = self.tracks[track_idx]['fx_list'][fx_idx]['modulations'][mod_idx]
            if mod['midiplink_line'] is None:
                return False
            line_num = mod['midiplink_line']
            old_line = self.lines[line_num]
            indent = old_line[:len(old_line) - len(old_line.lstrip())]
            self.lines[line_num] = f"{indent}MIDIPLINK {new_bus} {new_channel} 176 {new_cc}\r\n"
            mod['midi_cc']       = new_cc
            mod['midi_channel']  = new_channel
            mod['midi_bus']      = new_bus
            mod['midi_msg_type'] = 176
            self.modified = True
            return True
        except (IndexError, KeyError) as e:
            print(f"Error updating MIDI CC: {e}")
            return False

    def save_file(self, filepath=None):
        if filepath is None:
            filepath = self.filepath
        with open(filepath, 'w', encoding='utf-8', newline='') as f:
            f.writelines(self.lines)
        self.modified = False
        return True


# ═══════════════════════════════════════════════════════════════════════════
#  Colors
# ═══════════════════════════════════════════════════════════════════════════

AUDIO_COLOR  = '#4a9eff'
MIDI_COLOR   = '#ff7043'
BOTH_COLOR   = '#ffcc00'
FOLDER_COLOR = '#66bb6a'
NODE_FILL    = '#2d2d2d'
NODE_OUTLINE = '#888888'
NODE_TEXT    = '#eeeeee'
BG_COLOR     = '#1a1a1a'


# ═══════════════════════════════════════════════════════════════════════════
#  Routing Visualisation Window
# ═══════════════════════════════════════════════════════════════════════════

class RoutingWindow:
    def __init__(self, parent, project: REAPERProject):
        self.project = project
        self.win = tk.Toplevel(parent)
        self.win.title("Track Routing Visualisation")
        self.win.geometry("1080x740")
        self.win.configure(bg='#1a1a1a')

        self._node_pos: Dict[int, Tuple[float, float]] = {}
        self._drag_node: Optional[int] = None
        self._drag_off = (0, 0)
        self._scale = 1.0
        self._offset = [0, 0]
        self._selected_node: Optional[int] = None
        self._drag_moved = False

        # Graph filter state — must exist before _draw_graph is first called
        self._gf_audio   = tk.BooleanVar(value=True)
        self._gf_midi    = tk.BooleanVar(value=True)
        self._gf_midi_ch = tk.StringVar(value='All')
        self._gf_ch_dir  = tk.StringVar(value='either')

        self._build_ui()
        self._layout_nodes()
        self._refresh_all()

    # ── UI ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        nb = ttk.Notebook(self.win)
        nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        self.graph_frame  = tk.Frame(nb, bg=BG_COLOR)
        self.matrix_frame = tk.Frame(nb, bg='#1a1a1a')
        self.list_frame   = tk.Frame(nb, bg='#1a1a1a')

        nb.add(self.graph_frame,  text='  🔀  Node Graph  ')
        nb.add(self.matrix_frame, text='  ⊞  Matrix  ')
        nb.add(self.list_frame,   text='  ≡  Connection List  ')

        self._build_graph_tab()
        self._build_matrix_tab()
        self._build_list_tab()
        nb.bind('<<NotebookTabChanged>>', lambda e: self._refresh_all())

    # ── Node graph ───────────────────────────────────────────────────────

    def _build_graph_tab(self):
        # ── Row 1: legend + reset ────────────────────────────────────────
        tb = tk.Frame(self.graph_frame, bg='#252525')
        tb.pack(fill=tk.X)
        tk.Label(tb, text='Legend:', bg='#252525', fg='#aaa', font=('Arial', 9)).pack(side=tk.LEFT, padx=8)
        tk.Label(tb, text='━ Audio',      bg='#252525', fg=AUDIO_COLOR,  font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=4)
        tk.Label(tb, text='━ MIDI',       bg='#252525', fg=MIDI_COLOR,   font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=4)
        tk.Label(tb, text='━ Audio+MIDI', bg='#252525', fg=BOTH_COLOR,   font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=4)
        tk.Label(tb, text='╌ Folder',     bg='#252525', fg=FOLDER_COLOR, font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=4)
        tk.Button(tb, text='⟳ Reset layout', bg='#333', fg='#ddd', relief=tk.FLAT,
                  command=self._reset_layout, padx=8).pack(side=tk.RIGHT, padx=6, pady=3)
        tk.Label(tb, text='Drag nodes • Scroll to zoom', bg='#252525', fg='#555',
                 font=('Arial', 8)).pack(side=tk.RIGHT, padx=8)

        # ── Row 2: connection filters ────────────────────────────────────
        fb = tk.Frame(self.graph_frame, bg='#1e1e1e')
        fb.pack(fill=tk.X)

        tk.Label(fb, text='Show:', bg='#1e1e1e', fg='#aaa', font=('Arial', 9)).pack(side=tk.LEFT, padx=(8,4), pady=4)

        tk.Checkbutton(fb, text='Audio', variable=self._gf_audio, bg='#1e1e1e',
                       fg=AUDIO_COLOR, selectcolor='#333', activebackground='#1e1e1e',
                       command=self._draw_graph).pack(side=tk.LEFT, padx=4)
        tk.Checkbutton(fb, text='MIDI',  variable=self._gf_midi,  bg='#1e1e1e',
                       fg=MIDI_COLOR,  selectcolor='#333', activebackground='#1e1e1e',
                       command=self._draw_graph).pack(side=tk.LEFT, padx=4)

        tk.Label(fb, text='│', bg='#1e1e1e', fg='#444').pack(side=tk.LEFT, padx=4)
        tk.Label(fb, text='MIDI ch:', bg='#1e1e1e', fg='#aaa', font=('Arial', 9)).pack(side=tk.LEFT, padx=(4,2))

        ch_options = ['All'] + [str(i) for i in range(0, 16)]
        ch_menu = ttk.Combobox(fb, textvariable=self._gf_midi_ch, values=ch_options,
                                width=5, state='readonly')
        ch_menu.pack(side=tk.LEFT, padx=2)
        ch_menu.bind('<<ComboboxSelected>>', lambda e: self.win.after(10, self._draw_graph))

        tk.Label(fb, text='as:', bg='#1e1e1e', fg='#aaa', font=('Arial', 9)).pack(side=tk.LEFT, padx=(6,2))
        for val, txt in [('src','src'), ('dst','dst'), ('either','either')]:
            tk.Radiobutton(fb, text=txt, variable=self._gf_ch_dir, value=val,
                           bg='#1e1e1e', fg='#ccc', selectcolor='#333',
                           activebackground='#1e1e1e',
                           command=self._draw_graph).pack(side=tk.LEFT, padx=2)

        tk.Button(fb, text='✕ clear', bg='#2a2a2a', fg='#aaa', relief=tk.FLAT,
                  font=('Arial', 8),
                  command=self._graph_filter_clear).pack(side=tk.RIGHT, padx=8, pady=2)

        self.gc = tk.Canvas(self.graph_frame, bg=BG_COLOR, highlightthickness=0)
        self.gc.pack(fill=tk.BOTH, expand=True)
        self.gc.bind('<ButtonPress-1>',   self._gp)
        self.gc.bind('<B1-Motion>',       self._gd)
        self.gc.bind('<ButtonRelease-1>', self._gr)
        self.gc.bind('<MouseWheel>',      self._gz)
        self.gc.bind('<Button-4>',        self._gz)
        self.gc.bind('<Button-5>',        self._gz)
        self.gc.bind('<Configure>',       lambda e: self._draw_graph())
        self.gc.bind('<Motion>',          self._gt)

        self._tip = tk.Label(self.graph_frame, text='', bg='#333', fg='#eee',
                              font=('Arial', 8), relief=tk.FLAT, padx=4, pady=2)

    def _graph_filter_clear(self):
        self._gf_audio.set(True)
        self._gf_midi.set(True)
        self._gf_midi_ch.set('All')
        self._gf_ch_dir.set('either')
        self._selected_node = None
        self.win.after(10, self._draw_graph)

    def _send_passes_filter(self, send: dict) -> bool:
        """Return True if this send should be visible given current graph filters."""
        show_audio = self._gf_audio.get()
        show_midi  = self._gf_midi.get()
        ch_filter  = self._gf_midi_ch.get()   # 'All' or '1'..'16'
        ch_dir     = self._gf_ch_dir.get()    # 'src', 'dst', 'either'

        ha = send.get('has_audio', False)
        hm = send.get('has_midi',  False)

        # Type visibility
        visible_audio = ha and show_audio
        visible_midi  = hm and show_midi

        if not visible_audio and not visible_midi:
            return False

        # MIDI channel filter (only applies when a specific channel is chosen)
        if visible_midi and ch_filter != 'All':
            # Both dropdown and Ch labels are now 0-based
            wanted = f'Ch {ch_filter}'
            src_ch = send.get('midi_src_ch') or ''
            dst_ch = send.get('midi_dst_ch') or ''
            if ch_dir == 'src':
                ch_ok = (src_ch == wanted)
            elif ch_dir == 'dst':
                ch_ok = (dst_ch == wanted)
            else:  # either
                ch_ok = (src_ch == wanted or dst_ch == wanted)
            if not ch_ok:
                visible_midi = False

        return visible_audio or visible_midi

    def _layout_nodes(self):
        n = len(self.project.tracks)
        if n == 0:
            return
        cx, cy = 450, 320
        r = min(280, max(100, n * 22))
        for i in range(n):
            a = 2 * math.pi * i / n - math.pi / 2
            self._node_pos[i] = (cx + r * math.cos(a), cy + r * math.sin(a))

    def _reset_layout(self):
        self._node_pos.clear()
        self._scale = 1.0
        self._offset = [0, 0]
        self._layout_nodes()
        self._draw_graph()

    def _w2s(self, wx, wy):
        s = self._scale; ox, oy = self._offset
        return wx * s + ox, wy * s + oy

    def _s2w(self, sx, sy):
        s = self._scale; ox, oy = self._offset
        return (sx - ox) / s, (sy - oy) / s

    def _node_at(self, sx, sy):
        R = 30 * self._scale
        for i, (wx, wy) in self._node_pos.items():
            cx, cy = self._w2s(wx, wy)
            if math.hypot(sx - cx, sy - cy) <= R:
                return i
        return None

    def _gp(self, e):
        nd = self._node_at(e.x, e.y)
        self._drag_moved = False
        if nd is not None:
            self._drag_node = nd
            wx, wy = self._node_pos[nd]
            sx, sy = self._w2s(wx, wy)
            self._drag_off = (e.x - sx, e.y - sy)
        else:
            self._drag_node = None
            self._drag_off = (e.x, e.y)

    def _gd(self, e):
        self._drag_moved = True
        if self._drag_node is not None:
            dx, dy = self._drag_off
            wx, wy = self._s2w(e.x - dx, e.y - dy)
            self._node_pos[self._drag_node] = (wx, wy)
        else:
            px, py = self._drag_off
            self._offset[0] += e.x - px
            self._offset[1] += e.y - py
            self._drag_off = (e.x, e.y)
        self._draw_graph()

    def _gr(self, e):
        """Release: if no drag, treat as selection click."""
        if not self._drag_moved:
            nd = self._node_at(e.x, e.y)
            if nd is not None:
                self._selected_node = None if nd == self._selected_node else nd
            else:
                self._selected_node = None
            self._draw_graph()
        self._drag_node = None

    def _gz(self, e):
        f = 1.1 if (e.num == 4 or e.delta > 0) else 0.9
        self._offset[0] = e.x + (self._offset[0] - e.x) * f
        self._offset[1] = e.y + (self._offset[1] - e.y) * f
        self._scale *= f
        self._draw_graph()

    def _gt(self, e):
        nd = self._node_at(e.x, e.y)
        if nd is not None and nd < len(self.project.tracks):
            t = self.project.tracks[nd]
            midi_sends  = sum(1 for s in t['sends']    if s['has_midi'])
            audio_sends = sum(1 for s in t['sends']    if s['has_audio'])
            midi_recv   = sum(1 for r in t['receives'] if r['has_midi'])
            audio_recv  = sum(1 for r in t['receives'] if r['has_audio'])
            tip = (f"[{nd+1}] {t['name'] or 'unnamed'} | {len(t['fx_list'])} FX | "
                   f"Audio sends:{audio_sends} recv:{audio_recv} | "
                   f"MIDI sends:{midi_sends} recv:{midi_recv}")
            self._tip.config(text=tip)
            self._tip.place(x=e.x + 12, y=e.y - 20)
        else:
            self._tip.place_forget()

    def _draw_graph(self):
        c = self.gc
        c.delete('all')
        if not self.project.tracks:
            c.create_text(400, 300, text='No project loaded', fill='#555', font=('Arial', 16))
            return

        sel = self._selected_node          # None = show all
        tracks = self.project.tracks

        # Build sets of nodes and edges relevant to the selection,
        # taking the current send filter into account.
        if sel is not None:
            connected_to:   set = {sel}
            connected_from: set = {sel}
            active_edges:   set = set()

            for send in tracks[sel]['sends']:
                if self._send_passes_filter(send):
                    dst = send['dst_idx']
                    connected_to.add(dst)
                    active_edges.add((sel, dst))
            for recv in tracks[sel]['receives']:
                # Build a synthetic send dict for the receive so we can filter it
                src = recv['src_idx']
                recv_as_send = {
                    'has_audio':   recv.get('has_audio', False),
                    'has_midi':    recv.get('has_midi',  False),
                    'midi_src_ch': recv.get('midi_src_ch'),
                    'midi_dst_ch': recv.get('midi_dst_ch'),
                    'src_ach':     recv.get('src_ach', 0),
                    'dst_ach':     recv.get('dst_ach', 0),
                    'fader_mode':  recv.get('fader_mode', 0),
                }
                if self._send_passes_filter(recv_as_send):
                    connected_from.add(src)
                    active_edges.add((src, sel))

            visible_nodes = connected_to | connected_from
        else:
            visible_nodes = set(self._node_pos.keys())
            active_edges  = None   # draw all, dimming handled per-edge

        R  = max(22, int(28 * self._scale))
        fs = max(7,  int(9  * self._scale))
        lfs = max(6, int(8  * self._scale))   # label font size
        drawn: set = set()

        def build_send_label(send: dict) -> str:
            """Build a compact info string for an edge label."""
            parts = []
            if send.get('has_audio'):
                ach = f"{decode_audio_ch(send['src_ach'])}→{decode_audio_ch(send['dst_ach'])}"
                mode = FADER_MODES.get(send.get('fader_mode', 0), '?')
                parts.append(f"♪ {ach} ({mode})")
            if send.get('has_midi'):
                sc = send.get('midi_src_ch') or 'All'
                dc = send.get('midi_dst_ch') or 'Orig'
                parts.append(f"M {sc}→{dc}")
            return "  |  ".join(parts) if parts else ''

        def draw_edge(src, dst, color, dash=(), alpha_dim=False, send=None):
            key = (src, dst)
            if key in drawn or src not in self._node_pos or dst not in self._node_pos:
                return
            drawn.add(key)
            sx2, sy2 = self._w2s(*self._node_pos[src])
            dx2, dy2 = self._w2s(*self._node_pos[dst])
            # Curved midpoint offset (perpendicular to the line)
            mx_ = (sx2 + dx2) / 2 + (dy2 - sy2) * 0.18
            my_ = (sy2 + dy2) / 2 + (sx2 - dx2) * 0.18
            w  = max(1, int(2 * self._scale))
            ar = (max(6, int(10*self._scale)), max(8, int(12*self._scale)), max(3, int(4*self._scale)))
            fill = '#3a3a3a' if alpha_dim else color
            lw   = max(1, int(1 * self._scale)) if alpha_dim else w
            c.create_line(sx2, sy2, mx_, my_, dx2, dy2,
                          smooth=True, fill=fill, width=lw, dash=dash,
                          arrow=tk.LAST, arrowshape=ar)

            # Draw label on active edges when a node is selected
            if not alpha_dim and sel is not None and send is not None:
                label = build_send_label(send)
                if label:
                    # Perpendicular unit vector (points "above" the line)
                    dx_ = dx2 - sx2
                    dy_ = dy2 - sy2
                    length = math.hypot(dx_, dy_) or 1
                    # Perpendicular: rotate 90° CCW = (-dy, dx)
                    px_ = -dy_ / length
                    py_ =  dx_ / length
                    # Offset 10px above the curve midpoint
                    offset = 10
                    lx = mx_ + px_ * offset
                    ly = my_ + py_ * offset
                    c.create_text(lx, ly, text=label, fill=color,
                                  font=('Arial', lfs, 'bold'), anchor='center',
                                  tags='edgelabel')

        # ── Edges ──────────────────────────────────────────────────────
        for ti, track in enumerate(tracks):
            if ti not in self._node_pos:
                continue
            for send in track['sends']:
                dst = send['dst_idx']
                if dst not in self._node_pos:
                    continue
                # Apply connection type + MIDI channel filter
                if not self._send_passes_filter(send):
                    continue
                ha, hm = send['has_audio'], send['has_midi']
                # Recompute color after filter (audio may be hidden)
                show_audio = self._gf_audio.get()
                show_midi  = self._gf_midi.get()
                ha_vis = ha and show_audio
                hm_vis = hm and show_midi
                color = BOTH_COLOR if (ha_vis and hm_vis) else (MIDI_COLOR if hm_vis else AUDIO_COLOR)
                is_active = (active_edges is None) or ((ti, dst) in active_edges)
                draw_edge(ti, dst, color, alpha_dim=not is_active, send=send)

            if track['folder_depth'] > 0:
                for ci in range(ti + 1, len(tracks)):
                    if ci in self._node_pos:
                        is_active = (active_edges is None) or ((ti, ci) in active_edges)
                        draw_edge(ti, ci, FOLDER_COLOR, dash=(6, 3), alpha_dim=not is_active)
                        break

        # ── Nodes ──────────────────────────────────────────────────────
        for i, track in enumerate(tracks):
            if i not in self._node_pos:
                continue
            wx, wy = self._node_pos[i]
            sx, sy = self._w2s(wx, wy)

            is_sel       = (i == sel)
            is_visible   = (sel is None) or (i in visible_nodes)
            is_neighbour = (sel is not None) and (i in visible_nodes) and not is_sel

            if is_sel:
                fill    = '#4a4a4a'
                outline = '#ffffff'
                ow      = max(2, int(3 * self._scale))
                text_col = '#ffffff'
            elif is_neighbour:
                fill    = NODE_FILL
                outline = AUDIO_COLOR
                ow      = max(1, int(2 * self._scale))
                text_col = NODE_TEXT
            elif sel is None:
                has_conn = bool(track['sends'] or track['receives'])
                fill    = NODE_FILL
                outline = AUDIO_COLOR if has_conn else NODE_OUTLINE
                ow      = max(1, int(2 * self._scale))
                text_col = NODE_TEXT
            else:
                # dimmed node — not connected to selection
                fill    = '#222222'
                outline = '#444444'
                ow      = 1
                text_col = '#555555'

            c.create_oval(sx-R, sy-R, sx+R, sy+R, fill=fill, outline=outline, width=ow)

            lbl = (track['name'] or f'T{i+1}')
            if len(lbl) > 12:
                lbl = lbl[:11] + '…'
            c.create_text(sx, sy - 4, text=lbl, fill=text_col,
                          font=('Arial', fs, 'bold'), anchor='center')
            nfx = len(track['fx_list'])
            if nfx:
                c.create_text(sx, sy + fs, text=f'{nfx} FX',
                              fill='#aaa' if is_visible else '#444',
                              font=('Arial', max(6, fs-1)), anchor='center')

        # ── Hint ───────────────────────────────────────────────────────
        if sel is not None:
            track = tracks[sel]
            hint = f"[{sel+1}] {track['name'] or 'unnamed'}  — {len(track['sends'])} sends, {len(track['receives'])} receives  •  click again to deselect"
            c.create_text(8, 8, text=hint, fill='#aaa', font=('Arial', 8), anchor='nw')
        else:
            c.create_text(8, 8, text='Click a node to highlight its connections',
                          fill='#555', font=('Arial', 8), anchor='nw')

    # ── Matrix ───────────────────────────────────────────────────────────

    CELL = 36
    LW   = 170
    LH   = 36

    def _build_matrix_tab(self):
        ctrl = tk.Frame(self.matrix_frame, bg='#252525')
        ctrl.pack(fill=tk.X)
        tk.Label(ctrl, text='Show:', bg='#252525', fg='#aaa', font=('Arial', 9)).pack(side=tk.LEFT, padx=8, pady=4)
        self.mx_audio = tk.BooleanVar(value=True)
        self.mx_midi  = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl, text='Audio', variable=self.mx_audio, bg='#252525',
                       fg=AUDIO_COLOR, selectcolor='#333', command=self._draw_matrix).pack(side=tk.LEFT, padx=4)
        tk.Checkbutton(ctrl, text='MIDI',  variable=self.mx_midi,  bg='#252525',
                       fg=MIDI_COLOR,  selectcolor='#333', command=self._draw_matrix).pack(side=tk.LEFT, padx=4)
        tk.Label(ctrl, text='♪=Audio  M=MIDI  ♪M=Both  |  Hover for channel details',
                 bg='#252525', fg='#777', font=('Arial', 8)).pack(side=tk.RIGHT, padx=10)

        fr = tk.Frame(self.matrix_frame, bg='#1a1a1a')
        fr.pack(fill=tk.BOTH, expand=True)
        self.mc = tk.Canvas(fr, bg='#1a1a1a', highlightthickness=0)
        sy = ttk.Scrollbar(fr, orient=tk.VERTICAL,   command=self.mc.yview)
        sx = ttk.Scrollbar(fr, orient=tk.HORIZONTAL, command=self.mc.xview)
        self.mc.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X)
        self.mc.pack(fill=tk.BOTH, expand=True)

        self._mx_tip = tk.Label(self.matrix_frame, text='', bg='#333', fg='#eee',
                                 font=('Arial', 8), relief=tk.FLAT, padx=4, pady=2)
        self.mc.bind('<Motion>', self._on_mx_motion)
        self._mx_conn: Dict[Tuple[int,int], dict] = {}

    def _on_mx_motion(self, event):
        cx = self.mc.canvasx(event.x)
        cy = self.mc.canvasy(event.y)
        CELL, LW, LH = self.CELL, self.LW, self.LH
        n = len(self.project.tracks)
        col = int((cx - LW) // CELL)
        row = int((cy - LH) // CELL)
        if 0 <= row < n and 0 <= col < n:
            send = self._mx_conn.get((row, col))
            if send:
                src_t = self.project.tracks[row]
                dst_t = self.project.tracks[col]
                parts = []
                if send['has_audio']:
                    parts.append(
                        f"Audio ({FADER_MODES.get(send['fader_mode'],'?')}) "
                        f"ch {decode_audio_ch(send['src_ach'])}→{decode_audio_ch(send['dst_ach'])}")
                if send['has_midi']:
                    parts.append(
                        f"MIDI in:{send['midi_src_ch']} → out:{send['midi_dst_ch']}")
                tip = f"{src_t['name'] or f'T{row+1}'} → {dst_t['name'] or f'T{col+1}'}:  {'  |  '.join(parts)}"
                self._mx_tip.config(text=tip)
                self._mx_tip.place(x=event.x + 10, y=event.y - 28, in_=self.mc)
                return
        self._mx_tip.place_forget()

    def _draw_matrix(self):
        c = self.mc
        c.delete('all')
        tracks = self.project.tracks
        n = len(tracks)
        if n == 0:
            c.create_text(200, 100, text='No project loaded', fill='#555', font=('Arial', 14))
            return

        CELL, LW, LH = self.CELL, self.LW, self.LH
        c.configure(scrollregion=(0, 0, LW + n*CELL + 4, LH + n*CELL + 4))

        show_audio = self.mx_audio.get()
        show_midi  = self.mx_midi.get()

        self._mx_conn = {}
        for si, track in enumerate(tracks):
            for send in track['sends']:
                di = send['dst_idx']
                ha = send['has_audio'] and show_audio
                hm = send['has_midi']  and show_midi
                if ha or hm:
                    self._mx_conn[(si, di)] = send

        # Column headers
        for j, t in enumerate(tracks):
            x0 = LW + j * CELL
            name = (t['name'] or f'T{j+1}')[:8]
            c.create_rectangle(x0, 0, x0+CELL, LH, fill='#2a2a2a', outline='#444')
            c.create_text(x0+CELL//2, LH//2, text=name, fill='#ccc',
                          font=('Arial', 8), angle=45, anchor='center')

        # Rows
        for i, t in enumerate(tracks):
            y0 = LH + i * CELL
            name = (t['name'] or f'T{i+1}')[:24]
            c.create_rectangle(0, y0, LW, y0+CELL, fill='#222', outline='#444')
            c.create_text(6, y0+CELL//2, text=f'{i+1}. {name}',
                          fill='#ccc', font=('Arial', 8), anchor='w')
            for j in range(n):
                x0 = LW + j * CELL
                bg = '#2d2d2d' if (i+j)%2==0 else '#272727'
                c.create_rectangle(x0, y0, x0+CELL, y0+CELL, fill=bg, outline='#333')
                if i == j:
                    c.create_line(x0, y0, x0+CELL, y0+CELL, fill='#444')
                send = self._mx_conn.get((i, j))
                if send:
                    ha = send['has_audio'] and show_audio
                    hm = send['has_midi']  and show_midi
                    if ha and hm:
                        color, sym = BOTH_COLOR, '♪M'
                    elif hm:
                        color, sym = MIDI_COLOR, 'M'
                    else:
                        color, sym = AUDIO_COLOR, '♪'
                    p = 5
                    c.create_rectangle(x0+p, y0+p, x0+CELL-p, y0+CELL-p, fill=color, outline='')
                    c.create_text(x0+CELL//2, y0+CELL//2, text=sym,
                                  fill='white', font=('Arial', 8, 'bold'))

        c.create_text(LW//2, LH//2, text='SRC \\ DST', fill='#888', font=('Arial', 8, 'bold'))

    # ── Text list ────────────────────────────────────────────────────────

    def _build_list_tab(self):
        ctrl = tk.Frame(self.list_frame, bg='#252525')
        ctrl.pack(fill=tk.X)
        tk.Label(ctrl, text='Group by:', bg='#252525', fg='#aaa', font=('Arial', 9)).pack(side=tk.LEFT, padx=8, pady=4)
        self.list_grp = tk.StringVar(value='source')
        for val, txt in [('source','Source'),('dest','Destination'),('type','Type')]:
            tk.Radiobutton(ctrl, text=txt, variable=self.list_grp, value=val,
                           bg='#252525', fg='#ccc', selectcolor='#333',
                           command=self._draw_list).pack(side=tk.LEFT, padx=6)

        self.lt = scrolledtext.ScrolledText(self.list_frame, bg='#1e1e1e', fg='#ddd',
                                             font=('Consolas', 10), wrap=tk.NONE, state=tk.DISABLED)
        self.lt.pack(fill=tk.BOTH, expand=True)
        self.lt.tag_config('header', foreground='#fff',      font=('Consolas', 10, 'bold'))
        self.lt.tag_config('audio',  foreground=AUDIO_COLOR)
        self.lt.tag_config('midi',   foreground=MIDI_COLOR)
        self.lt.tag_config('both',   foreground=BOTH_COLOR)
        self.lt.tag_config('dim',    foreground='#555')

    def _draw_list(self):
        t = self.lt
        t.config(state=tk.NORMAL)
        t.delete('1.0', tk.END)
        tracks = self.project.tracks

        if not tracks:
            t.insert(tk.END, 'No project loaded.\n', 'dim')
            t.config(state=tk.DISABLED)
            return

        def tlabel(idx):
            if idx < len(tracks):
                return f"[{idx+1}] {tracks[idx]['name'] or f'Track {idx+1}'}"
            return f'[{idx+1}] ?'

        # Collect all connections
        conns = []
        for si, track in enumerate(tracks):
            for send in track['sends']:
                conns.append({'src': si, **send})

        if not conns:
            t.insert(tk.END, 'No sends/receives found.\n\n', 'dim')
            t.insert(tk.END,
                     'REAPER stores sends as AUXRECV entries on the destination\n'
                     'track. Projects with only a default master bus will appear\n'
                     'empty here.\n', 'dim')
            t.config(state=tk.DISABLED)
            return

        def send_detail(s) -> Tuple[str, List[Tuple[str, str]]]:
            """Returns (tag, [(text, tag), ...])"""
            lines = []
            if s['has_audio'] and s['has_midi']:
                tag = 'both'
            elif s['has_midi']:
                tag = 'midi'
            else:
                tag = 'audio'

            if s['has_audio']:
                ach = f"{decode_audio_ch(s['src_ach'])}→{decode_audio_ch(s['dst_ach'])}"
                lines.append((f"      Audio ({FADER_MODES.get(s['fader_mode'],'?')})  ch {ach}\n", 'audio'))
            if s['has_midi']:
                lines.append((f"      MIDI  in:{s['midi_src_ch']} → out:{s['midi_dst_ch']}\n", 'midi'))
            return tag, lines

        grp = self.list_grp.get()

        if grp == 'source':
            by: Dict[int, list] = {}
            for c in conns:
                by.setdefault(c['src'], []).append(c)
            for src in sorted(by):
                t.insert(tk.END, f'\n▶  {tlabel(src)}\n', 'header')
                for s in by[src]:
                    tag, detail = send_detail(s)
                    sym = '♪M' if (s['has_audio'] and s['has_midi']) else ('M' if s['has_midi'] else '♪')
                    t.insert(tk.END, f'   {sym}  →  {tlabel(s["dst_idx"])}\n', tag)
                    for line, ltag in detail:
                        t.insert(tk.END, line, ltag)

        elif grp == 'dest':
            by: Dict[int, list] = {}
            for c in conns:
                by.setdefault(c['dst_idx'], []).append(c)
            for dst in sorted(by):
                t.insert(tk.END, f'\n◀  {tlabel(dst)}\n', 'header')
                for s in by[dst]:
                    tag, detail = send_detail(s)
                    sym = '♪M' if (s['has_audio'] and s['has_midi']) else ('M' if s['has_midi'] else '♪')
                    t.insert(tk.END, f'   {sym}  ←  {tlabel(s["src"])}\n', tag)
                    for line, ltag in detail:
                        t.insert(tk.END, line, ltag)

        elif grp == 'type':
            audio_c = [c for c in conns if c['has_audio'] and not c['has_midi']]
            midi_c  = [c for c in conns if c['has_midi']  and not c['has_audio']]
            both_c  = [c for c in conns if c['has_audio'] and c['has_midi']]

            for label, lst, ttag in [
                ('♪  AUDIO SENDS',        audio_c, 'audio'),
                ('M  MIDI SENDS',          midi_c,  'midi'),
                ('♪M AUDIO + MIDI SENDS',  both_c,  'both'),
            ]:
                t.insert(tk.END, f'\n{label}\n', ttag)
                if lst:
                    for s in lst:
                        _, detail = send_detail(s)
                        t.insert(tk.END,
                                 f'   {tlabel(s["src"])}  →  {tlabel(s["dst_idx"])}\n', ttag)
                        for line, ltag in detail:
                            t.insert(tk.END, line, ltag)
                else:
                    t.insert(tk.END, '   (none)\n', 'dim')

        t.insert(tk.END, f'\n\nTotal connections: {len(conns)}\n', 'dim')
        t.config(state=tk.DISABLED)

    # ── Refresh ──────────────────────────────────────────────────────────

    def _refresh_all(self, *_):
        if not self._node_pos and self.project.tracks:
            self._layout_nodes()
        self._draw_graph()
        self._draw_matrix()
        self._draw_list()


# ═══════════════════════════════════════════════════════════════════════════
#  Main editor window
# ═══════════════════════════════════════════════════════════════════════════

class MIDICCEditorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("REAPER MIDI CC Editor")
        self.root.geometry("1200x720")
        self.project = REAPERProject()
        self._track_item_ids: List[str] = []
        self._item_to_indices: Dict[str, Tuple[int,int,int]] = {}
        self._routing_win: Optional[RoutingWindow] = None
        self._filter_var = tk.StringVar()   # initialised before _create_widgets
        self._create_widgets()
        self._create_menu()
        self._filter_var.trace_add('write', lambda *_: self._apply_filter())

    def _create_menu(self):
        mb = tk.Menu(self.root)
        self.root.config(menu=mb)

        fm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="File", menu=fm)
        fm.add_command(label="Open RPP File...", command=self.open_file)
        fm.add_command(label="Save",       command=self.save_file,    state=tk.DISABLED)
        fm.add_command(label="Save As...", command=self.save_file_as, state=tk.DISABLED)
        fm.add_separator()
        fm.add_command(label="Exit", command=self.root.quit)

        vm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="View", menu=vm)
        vm.add_command(label="Fold All Tracks",   command=self.fold_all)
        vm.add_command(label="Unfold All Tracks", command=self.unfold_all)
        vm.add_separator()
        vm.add_command(label="Fold All FX",   command=self.fold_all_fx)
        vm.add_command(label="Unfold All FX", command=self.unfold_all_fx)
        vm.add_separator()
        vm.add_command(label="🔀  Routing Visualisation", command=self.open_routing)

        self.file_menu = fm

    def _create_widgets(self):
        top = ttk.Frame(self.root, padding="10")
        top.pack(fill=tk.X)
        ttk.Label(top, text="File:").pack(side=tk.LEFT)
        self.file_label = tk.Label(top, text="No file loaded", fg="#888888",
                                    bg=self.root.cget('bg'), font=('Arial', 9))
        self.file_label.pack(side=tk.LEFT, padx=10)

        # Save buttons in top bar (always visible)
        self.save_as_btn = ttk.Button(top, text="Save As…", command=self.save_file_as, state=tk.DISABLED)
        self.save_as_btn.pack(side=tk.RIGHT, padx=2)
        self.save_btn = ttk.Button(top, text="💾 Save", command=self.save_file, state=tk.DISABLED)
        self.save_btn.pack(side=tk.RIGHT, padx=2)
        ttk.Button(top, text="Open File",  command=self.open_file).pack(side=tk.RIGHT, padx=2)
        ttk.Button(top, text="🔀 Routing", command=self.open_routing).pack(side=tk.RIGHT, padx=6)

        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))

        left = ttk.Frame(main)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        hf = ttk.Frame(left)
        hf.pack(fill=tk.X, pady=(0,2))
        ttk.Label(hf, text="Tracks & FX", font=('Arial',10,'bold')).pack(side=tk.LEFT)
        bf = ttk.Frame(hf)
        bf.pack(side=tk.RIGHT)
        ttk.Button(bf, text="⊟ Fold All",   command=self.fold_all,   width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="⊞ Unfold All", command=self.unfold_all, width=11).pack(side=tk.LEFT, padx=2)

        # ── Filter bar ──────────────────────────────────────────────────
        ff = ttk.Frame(left)
        ff.pack(fill=tk.X, pady=(0,4))
        ttk.Label(ff, text="🔍", font=('Arial', 10)).pack(side=tk.LEFT)
        # _filter_var already created in __init__ — just bind Entry to it
        self._filter_entry = ttk.Entry(ff, textvariable=self._filter_var)
        self._filter_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self._filter_clear = ttk.Button(ff, text="✕", width=2,
                                         command=lambda: self._filter_var.set(''))
        self._filter_clear.pack(side=tk.LEFT)
        self._filter_active = False   # tracks whether filter is currently applied

        ts = ttk.Scrollbar(left)
        ts.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree = ttk.Treeview(left, yscrollcommand=ts.set, selectmode='extended',
                                  columns=('Type','CC','Channel'), show='tree headings')
        self.tree.heading('#0',      text='Name')
        self.tree.heading('Type',    text='Type')
        self.tree.heading('CC',      text='MIDI CC')
        self.tree.heading('Channel', text='Ch')
        self.tree.column('#0',       width=350)
        self.tree.column('Type',     width=80)
        self.tree.column('CC',       width=80)
        self.tree.column('Channel',  width=50)
        self.tree.pack(fill=tk.BOTH, expand=True)
        ts.config(command=self.tree.yview)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Double-1>',         self.on_dbl)

        right = ttk.Frame(main, padding="10")
        right.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10,0))
        ttk.Label(right, text="Edit MIDI CC Assignment",
                  font=('Arial',10,'bold')).pack(pady=(0,4))
        self.sel_lbl = ttk.Label(right, text="No parameters selected",
                                  foreground="gray", font=('Arial',9,'italic'))
        self.sel_lbl.pack()

        ef = ttk.LabelFrame(right, text="Bulk Assign", padding="10")
        ef.pack(fill=tk.X, pady=(8,0))

        ttk.Label(ef, text="MIDI CC:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.cc_sb = ttk.Spinbox(ef, from_=0, to=127, width=8)
        self.cc_sb.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        self.cc_sb.set(0)
        self.cc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(ef, text="apply", variable=self.cc_var).grid(row=0, column=2, sticky=tk.W)

        ttk.Label(ef, text="MIDI Channel:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.ch_sb = ttk.Spinbox(ef, from_=0, to=16, width=8)
        self.ch_sb.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        self.ch_sb.set(0)
        self.ch_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(ef, text="apply", variable=self.ch_var).grid(row=1, column=2, sticky=tk.W)

        ttk.Label(ef, text="MIDI Bus:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.bus_sb = ttk.Spinbox(ef, from_=0, to=15, width=8)
        self.bus_sb.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        self.bus_sb.set(0)
        self.bus_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(ef, text="apply", variable=self.bus_var).grid(row=2, column=2, sticky=tk.W)

        ttk.Label(ef, text="Tick 'apply' to overwrite that field on all selected rows.",
                  foreground="#666", font=('Arial',8), wraplength=220
                  ).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(4,0))
        self.apply_btn = ttk.Button(ef, text="▶  Apply to Selected",
                                     command=self.apply_changes, state=tk.DISABLED)
        self.apply_btn.grid(row=4, column=0, columnspan=3, pady=(12,4))

        inf = ttk.LabelFrame(right, text="Selection Info", padding="10")
        inf.pack(fill=tk.BOTH, expand=True, pady=(10,0))
        self.info_text = scrolledtext.ScrolledText(inf, height=14, width=40,
                                                    wrap=tk.WORD, state=tk.DISABLED)
        self.info_text.pack(fill=tk.BOTH, expand=True)

        self.status = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    # ── Routing window ───────────────────────────────────────────────────

    def open_routing(self):
        if not self.project.tracks:
            messagebox.showinfo("No project", "Please open a REAPER project first.")
            return
        try:
            alive = self._routing_win and self._routing_win.win.winfo_exists()
        except Exception:
            alive = False
        if alive:
            self._routing_win.win.lift()
            self._routing_win._refresh_all()
        else:
            self._routing_win = RoutingWindow(self.root, self.project)

    # ── Fold/Unfold ──────────────────────────────────────────────────────

    def fold_all(self):
        for iid in self._track_item_ids:
            self.tree.item(iid, open=False)
        self.status.config(text="All tracks folded")

    def unfold_all(self):
        for iid in self._track_item_ids:
            self.tree.item(iid, open=True)
            for fx in self.tree.get_children(iid):
                self.tree.item(fx, open=True)
        self.status.config(text="All tracks unfolded")

    def fold_all_fx(self):
        for iid in self._track_item_ids:
            self.tree.item(iid, open=True)
            for fx in self.tree.get_children(iid):
                self.tree.item(fx, open=False)

    def unfold_all_fx(self):
        for iid in self._track_item_ids:
            self.tree.item(iid, open=True)
            for fx in self.tree.get_children(iid):
                self.tree.item(fx, open=True)

    def on_dbl(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            tags = self.tree.item(item, 'tags')
            if tags and tags[0] in ('track', 'fx'):
                self.tree.item(item, open=not self.tree.item(item, 'open'))

    # ── File I/O ─────────────────────────────────────────────────────────

    def open_file(self):
        fp = filedialog.askopenfilename(
            title="Open REAPER Project",
            filetypes=[("REAPER Project", "*.RPP *.rpp"), ("All Files", "*.*")])
        if not fp:
            return
        try:
            self.project.load_file(fp)
            self.file_label.config(text=os.path.basename(fp), fg='#ffffff')
            self.populate_tree()
            self.status.config(text=f"Loaded: {fp}")
            self.file_menu.entryconfig("Save",       state=tk.NORMAL)
            self.file_menu.entryconfig("Save As...", state=tk.NORMAL)
            self.save_btn.config(state=tk.NORMAL)
            self.save_as_btn.config(state=tk.NORMAL)
            self.show_stats()
            try:
                if self._routing_win and self._routing_win.win.winfo_exists():
                    self._routing_win.project = self.project
                    self._routing_win._node_pos.clear()
                    self._routing_win._selected_node = None
                    self._routing_win._refresh_all()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")

    def save_file(self):
        if not self.project.modified:
            messagebox.showinfo("Info", "No changes to save"); return
        try:
            self.project.save_file()
            self.status.config(text=f"Saved: {self.project.filepath}")
            messagebox.showinfo("Success", "File saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

    def save_file_as(self):
        fp = filedialog.asksaveasfilename(title="Save As", defaultextension=".RPP",
                                          filetypes=[("REAPER Project","*.RPP"),("All Files","*.*")])
        if fp:
            try:
                self.project.save_file(fp)
                self.file_label.config(text=os.path.basename(fp), fg='#ffffff')
                self.status.config(text=f"Saved: {fp}")
                messagebox.showinfo("Success", "File saved successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

    # ── Tree ─────────────────────────────────────────────────────────────

    def populate_tree(self):
        """Rebuild tree, honouring the current filter string."""
        query = self._filter_var.get().strip().lower() if hasattr(self, '_filter_var') else ''
        filtering = bool(query)

        for item in self.tree.get_children():
            self.tree.delete(item)
        self._track_item_ids = []
        self._item_to_indices = {}

        for ti, track in enumerate(self.project.tracks):
            track_name = track['name'] or f"Track {ti+1}"
            track_id = None   # inserted lazily

            for fi, fx in enumerate(track['fx_list']):
                fx_id = None   # inserted lazily

                for mi, mod in enumerate(fx['modulations']):
                    if mod['midi_cc'] is None:
                        continue
                    # Filter: skip param rows that don't match (when filter is active)
                    if filtering and query not in mod['param_name'].lower():
                        continue

                    # Ensure track header exists (inserted once per track)
                    if track_id is None:
                        track_id = self.tree.insert(
                            '', 'end', text=track_name,
                            values=('Track', '', ''),
                            tags=('track',), open=True)
                        self._track_item_ids.append(track_id)

                    # Ensure FX header exists under this track
                    if fx_id is None:
                        fx_id = self.tree.insert(
                            track_id, 'end', text=fx['name'],
                            values=(fx['type'], '', ''),
                            tags=('fx',), open=True)

                    iid = self.tree.insert(
                        fx_id, 'end',
                        text=mod['param_name'],
                        values=('Param', f"CC {mod['midi_cc']}", f"{mod['midi_channel']}"),
                        tags=('modulation', str(ti), str(fi), str(mi)))
                    self._item_to_indices[iid] = (ti, fi, mi)

        # Update filter entry appearance to signal active state
        if hasattr(self, '_filter_entry'):
            self._filter_entry.configure(
                style='Filter.TEntry' if filtering else 'TEntry')

    def _apply_filter(self):
        """Called whenever the filter text changes."""
        self.populate_tree()
        q = self._filter_var.get().strip()
        if q:
            total = len(self._item_to_indices)
            self.status.config(text=f"Filter '{q}' — {total} param(s) shown")
        else:
            self.status.config(text="Filter cleared")

    # ── Selection ────────────────────────────────────────────────────────

    def _get_sel(self):
        result = []
        for iid in self.tree.selection():
            idx = self._item_to_indices.get(iid)
            if idx:
                result.append((iid, *idx))
        return result

    def on_tree_select(self, event):
        sel = self._get_sel()
        n = len(sel)
        if n == 0:
            self.apply_btn.config(state=tk.DISABLED)
            self.sel_lbl.config(text="No parameters selected", foreground="gray")
            self.show_stats()
            return
        self.apply_btn.config(state=tk.NORMAL)
        ccs, chs, buses, names = set(), set(), set(), []
        for _, t, f, m in sel:
            mod = self.project.tracks[t]['fx_list'][f]['modulations'][m]
            ccs.add(mod['midi_cc']); chs.add(mod['midi_channel'])
            buses.add(mod['midi_bus']); names.append(mod['param_name'])
        self.cc_sb.set(next(iter(ccs))  if len(ccs)==1  else 0)
        self.ch_sb.set(next(iter(chs))  if len(chs)==1  else 0)
        self.bus_sb.set(next(iter(buses)) if len(buses)==1 else 0)
        if n == 1:
            self.sel_lbl.config(text=f"1 parameter: {names[0]}", foreground="black")
        else:
            self.sel_lbl.config(text=f"{n} parameters selected", foreground="#0055aa")
        lines = [f"{n} parameter(s) selected:\n"]
        for _, t, f, m in sel[:30]:
            tr = self.project.tracks[t]; fx = tr['fx_list'][f]; mod = fx['modulations'][m]
            lines.append(f"• {mod['param_name']}\n  Track: {tr['name'] or f'T{t+1}'}\n"
                         f"  FX: {fx['name']}\n  CC {mod['midi_cc']}  Ch {mod['midi_channel']}  Bus {mod['midi_bus']}\n")
        if n > 30:
            lines.append(f"… and {n-30} more")
        self._set_info("\n".join(lines))

    # ── Apply ────────────────────────────────────────────────────────────

    def apply_changes(self):
        sel = self._get_sel()
        if not sel: return
        do_cc, do_ch, do_bus = self.cc_var.get(), self.ch_var.get(), self.bus_var.get()
        if not any([do_cc, do_ch, do_bus]):
            messagebox.showwarning("Nothing", "Tick at least one 'apply' checkbox."); return
        try:
            ncc = int(self.cc_sb.get()); nch = int(self.ch_sb.get()); nbus = int(self.bus_sb.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid value."); return
        if do_cc  and not (0 <= ncc  <= 127): messagebox.showerror("Error","CC 0–127");      return
        if do_ch  and not (0 <= nch  <= 16):  messagebox.showerror("Error","Channel 0–16"); return
        if do_bus and not (0 <= nbus <= 15):  messagebox.showerror("Error","Bus 0–15");      return
        ok = fail = 0
        for _, t, f, m in sel:
            mod = self.project.tracks[t]['fx_list'][f]['modulations'][m]
            cc  = ncc  if do_cc  else mod['midi_cc']
            ch  = nch  if do_ch  else mod['midi_channel']
            bus = nbus if do_bus else mod['midi_bus']
            if self.project.update_midi_cc(t, f, m, cc, ch, bus): ok += 1
            else: fail += 1
        self.populate_tree()
        parts = ([f"CC→{ncc}"] if do_cc else []) + ([f"Ch→{nch}"] if do_ch else []) + ([f"Bus→{nbus}"] if do_bus else [])
        summary = ",  ".join(parts)
        if fail == 0:
            self.status.config(text=f"Updated {ok} row(s): {summary}  (not saved yet)")
            messagebox.showinfo("Done", f"Updated {ok} parameter(s):\n{summary}")
        else:
            self.status.config(text=f"Updated {ok}, skipped {fail}")
            messagebox.showwarning("Partial", f"Updated: {ok}\nSkipped (no MIDIPLINK): {fail}")

    # ── Stats ────────────────────────────────────────────────────────────

    def show_stats(self):
        tr = self.project.tracks
        midi_sends  = sum(1 for t in tr for s in t['sends'] if s['has_midi'])
        audio_sends = sum(1 for t in tr for s in t['sends'] if s['has_audio'])
        self.sel_lbl.config(text="No parameters selected", foreground="gray")
        self._set_info(
            "Project Statistics:\n\n"
            f"  Tracks:           {len(tr)}\n"
            f"  FX Plugins:       {sum(len(t['fx_list']) for t in tr)}\n"
            f"  MIDI CC Assigned: {sum(1 for t in tr for fx in t['fx_list'] for m in fx['modulations'] if m['midi_cc'] is not None)}\n"
            f"  Audio sends:      {audio_sends}\n"
            f"  MIDI sends:       {midi_sends}\n\n"
            "─────────────────────────────\n"
            "AUXRECV field 11 encodes:\n"
            "  src = val & 0x1F\n"
            "  dst = val >> 5\n"
            "  (per CockosWiki docs)\n\n"
            "Ctrl/Shift+click to multi-select\n"
            "Click 🔀 Routing to visualise."
        )

    def _set_info(self, text):
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(1.0, text)
        self.info_text.config(state=tk.DISABLED)


def main():
    root = tk.Tk()
    MIDICCEditorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

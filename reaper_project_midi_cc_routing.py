#!/usr/bin/env python3
"""
REAPER MIDI CC Editor - GUI Application
Allows viewing and editing MIDI CC assignments in REAPER .RPP files
Supports multi-row selection for bulk CC / Channel / Bus reassignment.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re
from typing import List, Dict, Tuple
import os


class REAPERProject:
    """Class to handle REAPER project file parsing and modification"""

    def __init__(self):
        self.filepath = None
        self.lines = []
        self.tracks = []
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
                    'line_num': i,
                }
                self.tracks.append(current_track)
                in_fxchain = False
                current_fx = None
                current_programenv = None

            elif current_track and stripped.startswith('NAME '):
                name = stripped[5:].strip('"')
                current_track['name'] = name if name else None

            elif current_track and stripped.startswith('<FXCHAIN'):
                in_fxchain = True
                fxchain_depth = indent
                current_fx = None
                current_programenv = None

            elif in_fxchain and current_track:
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
                            'param_id': m.group(1),
                            'param_name': m.group(3),
                            'bypass_flag': int(m.group(2)),
                            'midi_cc': None,
                            'midi_channel': None,
                            'midi_bus': None,
                            'midi_msg_type': None,
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
                        midi_bus = int(mm.group(1))
                        midi_channel = int(mm.group(2))
                        msg_type = int(mm.group(3))
                        cc_or_note = int(mm.group(4))
                        current_programenv['midi_bus'] = midi_bus
                        current_programenv['midi_channel'] = midi_channel
                        current_programenv['midi_msg_type'] = msg_type
                        current_programenv['midiplink_line'] = i
                        if msg_type == 176:
                            current_programenv['midi_cc'] = cc_or_note
                        elif msg_type == 144:
                            current_programenv['midi_note'] = cc_or_note

                elif stripped == '>' and current_programenv:
                    current_programenv = None

            if stripped == '>' and in_fxchain and indent <= fxchain_depth:
                in_fxchain = False
                current_fx = None
                current_programenv = None

    def update_midi_cc(self, track_idx: int, fx_idx: int, mod_idx: int,
                       new_cc: int, new_channel: int, new_bus: int) -> bool:
        try:
            mod = self.tracks[track_idx]['fx_list'][fx_idx]['modulations'][mod_idx]
            if mod['midiplink_line'] is None:
                return False
            line_num = mod['midiplink_line']
            old_line = self.lines[line_num]
            indent = old_line[:len(old_line) - len(old_line.lstrip())]
            self.lines[line_num] = f"{indent}MIDIPLINK {new_bus} {new_channel} 176 {new_cc}\r\n"
            mod['midi_cc'] = new_cc
            mod['midi_channel'] = new_channel
            mod['midi_bus'] = new_bus
            mod['midi_msg_type'] = 176
            self.modified = True
            return True
        except (IndexError, KeyError) as e:
            print(f"Error updating MIDI CC: {e}")
            return False

    def save_file(self, filepath: str = None):
        if filepath is None:
            filepath = self.filepath
        with open(filepath, 'w', encoding='utf-8', newline='') as f:
            f.writelines(self.lines)
        self.modified = False
        return True


class MIDICCEditorGUI:
    """Main GUI application"""

    def __init__(self, root):
        self.root = root
        self.root.title("REAPER MIDI CC Editor")
        self.root.geometry("1200x720")

        self.project = REAPERProject()
        self._track_item_ids: List[str] = []
        # Maps tree item id -> (track_idx, fx_idx, mod_idx)
        self._item_to_indices: Dict[str, Tuple[int, int, int]] = {}

        self._create_widgets()
        self._create_menu()

    # ------------------------------------------------------------------ #
    #  Menu                                                                #
    # ------------------------------------------------------------------ #

    def _create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open RPP File...", command=self.open_file)
        file_menu.add_command(label="Save",      command=self.save_file,    state=tk.DISABLED)
        file_menu.add_command(label="Save As...", command=self.save_file_as, state=tk.DISABLED)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Fold All Tracks",   command=self.fold_all)
        view_menu.add_command(label="Unfold All Tracks", command=self.unfold_all)
        view_menu.add_separator()
        view_menu.add_command(label="Fold All FX",   command=self.fold_all_fx)
        view_menu.add_command(label="Unfold All FX", command=self.unfold_all_fx)

        self.file_menu = file_menu

    # ------------------------------------------------------------------ #
    #  Widgets                                                             #
    # ------------------------------------------------------------------ #

    def _create_widgets(self):
        # ---- top bar ----
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        ttk.Label(top_frame, text="File:").pack(side=tk.LEFT)
        self.file_label = ttk.Label(top_frame, text="No file loaded", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(top_frame, text="Open File", command=self.open_file).pack(side=tk.RIGHT)

        # ---- main split ----
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # ---- LEFT: tree ----
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(left_frame)
        header_frame.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(header_frame, text="Tracks & FX",
                  font=('Arial', 10, 'bold')).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(header_frame)
        btn_frame.pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="⊟ Fold All",   command=self.fold_all,   width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="⊞ Unfold All", command=self.unfold_all, width=11).pack(side=tk.LEFT, padx=2)

        tree_scroll = ttk.Scrollbar(left_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 'extended' selectmode enables Ctrl+click and Shift+click multi-select
        self.tree = ttk.Treeview(left_frame, yscrollcommand=tree_scroll.set,
                                 selectmode='extended',
                                 columns=('Type', 'CC', 'Channel'), show='tree headings')
        self.tree.heading('#0',      text='Name')
        self.tree.heading('Type',    text='Type')
        self.tree.heading('CC',      text='MIDI CC')
        self.tree.heading('Channel', text='Ch')
        self.tree.column('#0',       width=350)
        self.tree.column('Type',     width=80)
        self.tree.column('CC',       width=80)
        self.tree.column('Channel',  width=50)
        self.tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.tree.yview)

        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<Double-1>',         self.on_tree_double_click)

        # ---- RIGHT: editor ----
        right_frame = ttk.Frame(main_frame, padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))

        ttk.Label(right_frame, text="Edit MIDI CC Assignment",
                  font=('Arial', 10, 'bold')).pack(pady=(0, 4))

        # Selection summary line
        self.selection_label = ttk.Label(right_frame,
                                         text="No parameters selected",
                                         foreground="gray", font=('Arial', 9, 'italic'))
        self.selection_label.pack()

        # ---- editor fields ----
        self.editor_frame = ttk.LabelFrame(right_frame, text="Bulk Assign", padding="10")
        self.editor_frame.pack(fill=tk.X, pady=(8, 0))

        # Each field has a spinbox + "apply" checkbox so the user can opt
        # to change only some fields while leaving others untouched.

        # MIDI CC
        ttk.Label(self.editor_frame, text="MIDI CC:").grid(
            row=0, column=0, sticky=tk.W, pady=5)
        self.cc_spinbox = ttk.Spinbox(self.editor_frame, from_=0, to=127, width=8)
        self.cc_spinbox.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        self.cc_spinbox.set(0)
        self.cc_apply_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.editor_frame, text="apply", variable=self.cc_apply_var).grid(
            row=0, column=2, sticky=tk.W)

        # MIDI Channel
        ttk.Label(self.editor_frame, text="MIDI Channel:").grid(
            row=1, column=0, sticky=tk.W, pady=5)
        self.channel_spinbox = ttk.Spinbox(self.editor_frame, from_=0, to=16, width=8)
        self.channel_spinbox.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        self.channel_spinbox.set(0)
        self.channel_apply_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.editor_frame, text="apply", variable=self.channel_apply_var).grid(
            row=1, column=2, sticky=tk.W)

        # MIDI Bus
        ttk.Label(self.editor_frame, text="MIDI Bus:").grid(
            row=2, column=0, sticky=tk.W, pady=5)
        self.bus_spinbox = ttk.Spinbox(self.editor_frame, from_=0, to=15, width=8)
        self.bus_spinbox.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        self.bus_spinbox.set(0)
        self.bus_apply_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.editor_frame, text="apply", variable=self.bus_apply_var).grid(
            row=2, column=2, sticky=tk.W)

        ttk.Label(self.editor_frame,
                  text="Tick 'apply' to overwrite that field on all selected rows.",
                  foreground="#666", font=('Arial', 8), wraplength=220).grid(
            row=3, column=0, columnspan=3, sticky=tk.W, pady=(4, 0))

        self.apply_button = ttk.Button(self.editor_frame, text="▶  Apply to Selected",
                                       command=self.apply_changes, state=tk.DISABLED)
        self.apply_button.grid(row=4, column=0, columnspan=3, pady=(12, 4))

        # ---- info text ----
        info_frame = ttk.LabelFrame(right_frame, text="Selection Info", padding="10")
        info_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.info_text = scrolledtext.ScrolledText(info_frame, height=14, width=40,
                                                   wrap=tk.WORD, state=tk.DISABLED)
        self.info_text.pack(fill=tk.BOTH, expand=True)

        # ---- status bar ----
        self.status_bar = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ------------------------------------------------------------------ #
    #  Fold / Unfold                                                       #
    # ------------------------------------------------------------------ #

    def fold_all(self):
        for item_id in self._track_item_ids:
            self.tree.item(item_id, open=False)
        self.status_bar.config(text="All tracks folded")

    def unfold_all(self):
        for item_id in self._track_item_ids:
            self.tree.item(item_id, open=True)
            for fx_id in self.tree.get_children(item_id):
                self.tree.item(fx_id, open=True)
        self.status_bar.config(text="All tracks unfolded")

    def fold_all_fx(self):
        for item_id in self._track_item_ids:
            self.tree.item(item_id, open=True)
            for fx_id in self.tree.get_children(item_id):
                self.tree.item(fx_id, open=False)
        self.status_bar.config(text="All FX folded")

    def unfold_all_fx(self):
        for item_id in self._track_item_ids:
            self.tree.item(item_id, open=True)
            for fx_id in self.tree.get_children(item_id):
                self.tree.item(fx_id, open=True)
        self.status_bar.config(text="All FX unfolded")

    def on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        tags = self.tree.item(item, 'tags')
        if tags and tags[0] in ('track', 'fx'):
            self.tree.item(item, open=not self.tree.item(item, 'open'))

    # ------------------------------------------------------------------ #
    #  File I/O                                                            #
    # ------------------------------------------------------------------ #

    def open_file(self):
        filepath = filedialog.askopenfilename(
            title="Open REAPER Project",
            filetypes=[("REAPER Project", "*.RPP"), ("All Files", "*.*")]
        )
        if filepath:
            try:
                self.project.load_file(filepath)
                self.file_label.config(text=os.path.basename(filepath), foreground="black")
                self.populate_tree()
                self.status_bar.config(text=f"Loaded: {filepath}")
                self.file_menu.entryconfig("Save",      state=tk.NORMAL)
                self.file_menu.entryconfig("Save As...", state=tk.NORMAL)
                self.show_project_stats()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")

    def save_file(self):
        if not self.project.modified:
            messagebox.showinfo("Info", "No changes to save")
            return
        try:
            self.project.save_file()
            self.status_bar.config(text=f"Saved: {self.project.filepath}")
            messagebox.showinfo("Success", "File saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

    def save_file_as(self):
        filepath = filedialog.asksaveasfilename(
            title="Save REAPER Project As",
            defaultextension=".RPP",
            filetypes=[("REAPER Project", "*.RPP"), ("All Files", "*.*")]
        )
        if filepath:
            try:
                self.project.save_file(filepath)
                self.file_label.config(text=os.path.basename(filepath))
                self.status_bar.config(text=f"Saved: {filepath}")
                messagebox.showinfo("Success", "File saved successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

    # ------------------------------------------------------------------ #
    #  Tree                                                                #
    # ------------------------------------------------------------------ #

    def populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._track_item_ids = []
        self._item_to_indices = {}

        for track_idx, track in enumerate(self.project.tracks):
            track_name = track['name'] if track['name'] else f"Track {track_idx + 1}"
            track_id = self.tree.insert('', 'end', text=track_name,
                                        values=('Track', '', ''),
                                        tags=('track',), open=True)
            self._track_item_ids.append(track_id)

            for fx_idx, fx in enumerate(track['fx_list']):
                fx_id = self.tree.insert(track_id, 'end', text=fx['name'],
                                         values=(fx['type'], '', ''),
                                         tags=('fx',), open=True)

                for mod_idx, mod in enumerate(fx['modulations']):
                    if mod['midi_cc'] is not None:
                        iid = self.tree.insert(
                            fx_id, 'end',
                            text=mod['param_name'],
                            values=('Param',
                                    f"CC {mod['midi_cc']}",
                                    f"{mod['midi_channel']}"),
                            tags=('modulation',
                                  str(track_idx), str(fx_idx), str(mod_idx))
                        )
                        self._item_to_indices[iid] = (track_idx, fx_idx, mod_idx)

    # ------------------------------------------------------------------ #
    #  Selection                                                           #
    # ------------------------------------------------------------------ #

    def _get_selected_modulations(self):
        """Return [(item_id, track_idx, fx_idx, mod_idx), ...] for every
        selected row that is a modulation."""
        result = []
        for iid in self.tree.selection():
            indices = self._item_to_indices.get(iid)
            if indices is not None:
                result.append((iid, *indices))
        return result

    def on_tree_select(self, event):
        selected = self._get_selected_modulations()
        count = len(selected)

        if count == 0:
            self.apply_button.config(state=tk.DISABLED)
            self.selection_label.config(text="No parameters selected", foreground="gray")
            self.show_project_stats()
            return

        self.apply_button.config(state=tk.NORMAL)

        ccs, channels, buses, names = set(), set(), set(), []
        for _, t, f, m in selected:
            mod = self.project.tracks[t]['fx_list'][f]['modulations'][m]
            ccs.add(mod['midi_cc'])
            channels.add(mod['midi_channel'])
            buses.add(mod['midi_bus'])
            names.append(mod['param_name'])

        # Pre-fill spinboxes: use the common value if all rows agree, else 0
        self.cc_spinbox.set(next(iter(ccs))       if len(ccs)      == 1 else 0)
        self.channel_spinbox.set(next(iter(channels)) if len(channels) == 1 else 0)
        self.bus_spinbox.set(next(iter(buses))     if len(buses)    == 1 else 0)

        if count == 1:
            self.selection_label.config(
                text=f"1 parameter selected: {names[0]}", foreground="black")
        else:
            self.selection_label.config(
                text=f"{count} parameters selected  (Ctrl/Shift to add more)",
                foreground="#0055aa")

        # Build info summary (capped at 30 rows to avoid huge text)
        lines = [f"{count} parameter(s) selected:\n"]
        for _, t, f, m in selected[:30]:
            track = self.project.tracks[t]
            fx    = track['fx_list'][f]
            mod   = fx['modulations'][m]
            lines.append(
                f"• {mod['param_name']}\n"
                f"  Track : {track['name'] or f'Track {t+1}'}\n"
                f"  FX    : {fx['name']}\n"
                f"  CC {mod['midi_cc']}  Ch {mod['midi_channel']}  Bus {mod['midi_bus']}\n"
            )
        if count > 30:
            lines.append(f"… and {count - 30} more rows")
        self.update_info_text("\n".join(lines))

    # ------------------------------------------------------------------ #
    #  Apply                                                               #
    # ------------------------------------------------------------------ #

    def apply_changes(self):
        selected = self._get_selected_modulations()
        if not selected:
            return

        do_cc      = self.cc_apply_var.get()
        do_channel = self.channel_apply_var.get()
        do_bus     = self.bus_apply_var.get()

        if not any([do_cc, do_channel, do_bus]):
            messagebox.showwarning("Nothing to apply",
                                   "Tick at least one 'apply' checkbox to make changes.")
            return

        try:
            new_cc      = int(self.cc_spinbox.get())
            new_channel = int(self.channel_spinbox.get())
            new_bus     = int(self.bus_spinbox.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid value in one of the fields.")
            return

        if do_cc and not (0 <= new_cc <= 127):
            messagebox.showerror("Error", "MIDI CC must be 0–127.")
            return
        if do_channel and not (0 <= new_channel <= 16):
            messagebox.showerror("Error", "MIDI Channel must be 0–16.")
            return
        if do_bus and not (0 <= new_bus <= 15):
            messagebox.showerror("Error", "MIDI Bus must be 0–15.")
            return

        ok = fail = 0
        for _, t, f, m in selected:
            mod = self.project.tracks[t]['fx_list'][f]['modulations'][m]
            cc      = new_cc      if do_cc      else mod['midi_cc']
            channel = new_channel if do_channel else mod['midi_channel']
            bus     = new_bus     if do_bus     else mod['midi_bus']
            if self.project.update_midi_cc(t, f, m, cc, channel, bus):
                ok += 1
            else:
                fail += 1

        self.populate_tree()

        parts = []
        if do_cc:      parts.append(f"CC → {new_cc}")
        if do_channel: parts.append(f"Channel → {new_channel}")
        if do_bus:     parts.append(f"Bus → {new_bus}")
        summary = ",  ".join(parts)

        if fail == 0:
            self.status_bar.config(
                text=f"Updated {ok} row(s): {summary}  (not saved yet)")
            messagebox.showinfo("Done", f"Updated {ok} parameter(s):\n{summary}")
        else:
            self.status_bar.config(
                text=f"Updated {ok}, skipped {fail} (no MIDIPLINK)  (not saved yet)")
            messagebox.showwarning("Partial success",
                                   f"Updated : {ok}\n"
                                   f"Skipped : {fail}  (no MIDIPLINK entry)")

    # ------------------------------------------------------------------ #
    #  Info helpers                                                        #
    # ------------------------------------------------------------------ #

    def show_project_stats(self):
        total_tracks  = len(self.project.tracks)
        total_fx      = sum(len(t['fx_list']) for t in self.project.tracks)
        total_mods    = sum(len(fx['modulations'])
                            for t in self.project.tracks for fx in t['fx_list'])
        total_midi_cc = sum(1 for t in self.project.tracks
                            for fx in t['fx_list']
                            for mod in fx['modulations']
                            if mod['midi_cc'] is not None)

        self.selection_label.config(text="No parameters selected", foreground="gray")
        self.update_info_text(
            "Project Statistics:\n\n"
            f"  Tracks:           {total_tracks}\n"
            f"  FX Plugins:       {total_fx}\n"
            f"  Modulations:      {total_mods}\n"
            f"  MIDI CC Assigned: {total_midi_cc}\n\n"
            "─────────────────────────────\n"
            "Multi-select tips:\n"
            "  • Ctrl+click  — add / remove row\n"
            "  • Shift+click — select a range\n"
            "  • Ctrl+A      — select all visible\n\n"
            "Tick the 'apply' checkboxes next\n"
            "to each field to choose which\n"
            "values get overwritten."
        )

    def update_info_text(self, text: str):
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(1.0, text)
        self.info_text.config(state=tk.DISABLED)


def main():
    root = tk.Tk()
    app = MIDICCEditorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

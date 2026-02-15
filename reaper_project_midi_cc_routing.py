#!/usr/bin/env python3
"""
Correct MIDI CC Parser - based on REAPER State Chunk documentation
Parses REAPER .RPP files and extracts MIDI CC assignments for FX parameters
"""

import re
from collections import defaultdict

# Read the REAPER project file
with open('/mnt/user-data/uploads/all.RPP', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Track structure
tracks = []
current_track = None
current_fx = None
current_programenv = None
in_fxchain = False
fxchain_depth = 0

# Parse the file line by line
for i, line in enumerate(lines):
    indent = len(line) - len(line.lstrip())
    stripped = line.strip()
    
    # TRACK start
    if stripped.startswith('<TRACK '):
        guid_match = re.search(r'\{([^}]+)\}', stripped)
        current_track = {
            'guid': guid_match.group(1) if guid_match else 'Unknown',
            'name': None,
            'fx_list': [],
            'line_num': i + 1,
        }
        tracks.append(current_track)
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
        # FX plugin
        fx_match = re.match(r'<(VST|AU|JS|VST3|CLAP)\s+"(.+?)"', stripped)
        if fx_match:
            fx_type = fx_match.group(1)
            fx_name = fx_match.group(2)
            
            current_fx = {
                'type': fx_type,
                'name': fx_name,
                'line_num': i + 1,
                'modulations': [],
            }
            current_track['fx_list'].append(current_fx)
            current_programenv = None
        
        # PROGRAMENV - parameter modulation block start
        elif stripped.startswith('<PROGRAMENV '):
            programenv_match = re.match(r'<PROGRAMENV\s+(\S+)\s+(\d+)\s+"([^"]+)"', stripped)
            if programenv_match:
                param_id = programenv_match.group(1)
                bypass_flag = int(programenv_match.group(2))
                param_name = programenv_match.group(3)
                
                current_programenv = {
                    'param_id': param_id,
                    'param_name': param_name,
                    'bypass_flag': bypass_flag,
                    'midi_cc': None,
                    'midi_channel': None,
                    'midi_bus': None,
                    'line_num': i + 1,
                }
                
                if current_fx:
                    current_fx['modulations'].append(current_programenv)
                elif current_track['fx_list']:
                    current_track['fx_list'][-1]['modulations'].append(current_programenv)
        
        # MIDIPLINK - contains actual MIDI CC assignment
        elif stripped.startswith('MIDIPLINK ') and current_programenv:
            # Format: MIDIPLINK bus(0-15) channel(1-16) msg_type(176=CC) cc_number(0-127)
            midiplink_match = re.match(r'MIDIPLINK\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)', stripped)
            if midiplink_match:
                midi_bus = int(midiplink_match.group(1))
                midi_channel = int(midiplink_match.group(2))
                msg_type = int(midiplink_match.group(3))
                cc_or_note = int(midiplink_match.group(4))
                
                current_programenv['midi_bus'] = midi_bus
                current_programenv['midi_channel'] = midi_channel
                current_programenv['midi_msg_type'] = msg_type
                
                if msg_type == 176:  # CC message
                    current_programenv['midi_cc'] = cc_or_note
                elif msg_type == 144:  # Note
                    current_programenv['midi_note'] = cc_or_note
        
        # Closing PROGRAMENV block
        elif stripped == '>' and current_programenv:
            current_programenv = None
    
    if stripped == '>' and in_fxchain and indent <= fxchain_depth:
        in_fxchain = False
        current_fx = None
        current_programenv = None

# Print analysis
print("=" * 120)
print("REAPER PROJECT - POPRAWNA ANALIZA MIDI CC (według dokumentacji)")
print("=" * 120)
print()

# Collect statistics
total_modulations = 0
total_midi_assigned = 0
cc_usage = defaultdict(int)

for track in tracks:
    for fx in track['fx_list']:
        for mod in fx['modulations']:
            total_modulations += 1
            if mod['midi_cc'] is not None:
                total_midi_assigned += 1
                cc_usage[mod['midi_cc']] += 1

print(f"STATYSTYKI:")
print(f"  Całkowita liczba modulacji: {total_modulations}")
print(f"  Modulacje z przypisanym MIDI CC: {total_midi_assigned}")
print(f"  Modulacje bez MIDI CC: {total_modulations - total_midi_assigned}")
print()

if cc_usage:
    print(f"UŻYTE NUMERY MIDI CC:")
    for cc in sorted(cc_usage.keys()):
        print(f"  CC {cc:3d}: {cc_usage[cc]} przypisań")
    print()

# Print detailed list
print("\n" + "=" * 120)
print("SZCZEGÓŁOWA LISTA MODULACJI Z MIDI CC")
print("=" * 120)

for idx, track in enumerate(tracks, 1):
    track_name = track['name'] if track['name'] else '(Unnamed)'
    
    has_midi_cc = any(
        mod['midi_cc'] is not None 
        for fx in track['fx_list'] 
        for mod in fx['modulations']
    )
    
    if not has_midi_cc:
        continue
    
    print(f"\n{'═' * 120}")
    print(f"TRACK #{idx}: {track_name}")
    print(f"{'═' * 120}")
    
    for fx_idx, fx in enumerate(track['fx_list'], 1):
        fx_has_midi = any(mod['midi_cc'] is not None for mod in fx['modulations'])
        
        if not fx_has_midi:
            continue
        
        print(f"\n  FX #{fx_idx}: {fx['name']}")
        print(f"  {'─' * 116}")
        
        for mod in fx['modulations']:
            if mod['midi_cc'] is not None:
                status = "✓ ENABLED" if mod['bypass_flag'] == 0 else "✗ BYPASSED"
                print(f"    🎛️  CC {mod['midi_cc']:3d} (Ch {mod['midi_channel']:2d}, Bus {mod['midi_bus']}) → {mod['param_name']:<50s} [{status}]")

print("\n" + "=" * 120)
print()

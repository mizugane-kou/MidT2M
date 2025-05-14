

import sys
import os
import tempfile
from typing import List, Dict, Tuple, Optional, Any
import time # For debouncing
import traceback # For detailed error logging in threads
import json # For project save/load

from PySide6.QtCore import (
    Qt, QThread, Signal, Slot, QRectF, QPointF, QSize, QTimer, QSettings, QMimeData, QStandardPaths
)
from PySide6.QtGui import (
    QColor, QFont, QPainter, QPen, QBrush, QPalette, QFontMetrics, QAction, QKeySequence,
    QDragEnterEvent, QDropEvent # Added for D&D
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit,
    QSpinBox, QDoubleSpinBox, QGraphicsView, QGraphicsScene,
    QGraphicsRectItem, QGraphicsLineItem, QGraphicsTextItem,
    QColorDialog, QProgressBar, QMessageBox, QSplitter, QGroupBox,
    QFormLayout, QScrollArea, QComboBox, QCheckBox, QSlider
)

import mido
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import cv2



# --- START OF FONT UTILITIES (Windows specific) ---
IS_WINDOWS = (os.name == 'nt')
if IS_WINDOWS:
    try:
        import winreg
        from pathlib import Path
    except ImportError:
        IS_WINDOWS = False 
else:
    from pathlib import Path


def get_system_fonts_windows() -> Dict[str, str]:
    if not IS_WINDOWS:
        return {}
    fonts = {}
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
            for i in range(winreg.QueryInfoKey(key)[1]):
                try:
                    name, fontfile, _ = winreg.EnumValue(key, i)
                    display_name = name.split(' (')[0]
                    path = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", fontfile)
                    if os.path.exists(path):
                        if display_name not in fonts:
                            fonts[display_name] = path
                except OSError:
                    continue
    except Exception:
        pass
    return fonts

def get_user_fonts_windows() -> Dict[str, str]:
    if not IS_WINDOWS:
        return {}
    fonts = {}
    try:
        user_font_dir_str = os.getenv("LOCALAPPDATA")
        if user_font_dir_str:
            user_font_dir = Path(user_font_dir_str) / "Microsoft" / "Windows" / "Fonts"
            if user_font_dir.exists():
                for font_path in user_font_dir.glob("*.[oOtT][tT][fFcC]"):
                    display_name = font_path.stem
                    if display_name not in fonts:
                         fonts[display_name] = str(font_path)
    except Exception:
        pass
    return fonts

# --- END OF FONT UTILITIES ---


# --- START OF VIDEO GENERATION LOGIC ---
class ILogger:
    def info(self, msg: str): print(f"[INFO] {msg}")
    def warning(self, msg: str): print(f"[WARN] {msg}")
    def error(self, msg: str): print(f"[ERROR] {msg}")

class PrintLogger(ILogger):
    pass

def _calculate_adjusted_font_size(base_size: int, total_duration_seconds: float) -> int:
    if total_duration_seconds <= 0: return base_size
    REFERENCE_DURATION_SEC = 180.0; DURATION_SCALING_POWER = 0.3
    scale_factor = (REFERENCE_DURATION_SEC / max(30.0, total_duration_seconds)) ** DURATION_SCALING_POWER
    adjusted_size = int(base_size * scale_factor)
    return max(int(base_size * 0.5), min(int(base_size * 2.0), adjusted_size)) 

def _get_text_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont) -> int:
    if not text: return 0
    try: 
        return int(draw.textlength(text, font=font)) # Ensure integer
    except AttributeError: 
        try:
            bbox_for_width = draw.textbbox((0,0), text, font=font) 
            return bbox_for_width[2] - bbox_for_width[0]
        except AttributeError: 
            legacy_w, _ = draw.textsize(text, font=font) 
            return legacy_w
    except Exception: 
        return len(text) * font.size // 2 

def _calculate_fixed_layout_for_line_v2(
    segment_event_data_list: List[Dict[str, Any]], 
    font_path: str, font_size_base_for_line: int, char_spacing: int,
    canvas_width: int, line_h_align: str, line_anchor_x_abs: int,
    pitch_size_scale: float, velocity_size_scale: float,
    reference_pitch: int, reference_velocity: int,
    duration_padding_threshold_ticks: int, duration_padding_scale_per_tick: float,
    min_char_render_size: int, max_char_render_size: int,
    font_object_cache: Dict[Tuple[str, int], ImageFont.FreeTypeFont], 
    logger: ILogger
) -> Dict[str, Any]:
    layout_info = {'line_start_x_on_canvas': 0.0, 'total_width_for_alignment': 0.0}

    if not segment_event_data_list:
        start_x = float(line_anchor_x_abs)
        if line_h_align == "center": start_x = float(line_anchor_x_abs) # anchor_x - 0 / 2
        elif line_h_align == "right": start_x = float(line_anchor_x_abs) # anchor_x - 0
        layout_info['line_start_x_on_canvas'] = start_x
        return layout_info

    temp_draw = ImageDraw.Draw(Image.new('RGB',(1,1)))
    total_calculated_width = 0.0
    num_segments = len(segment_event_data_list)

    for idx, event_data in enumerate(segment_event_data_list):
        text_to_render_for_layout = event_data['text'] # This is 'text_for_layout'
        char_pitch = event_data['pitch']; char_velocity = event_data['velocity']
        char_duration_ticks = event_data.get('duration_ticks', 0)

        current_effective_font_size = float(font_size_base_for_line)
        if pitch_size_scale != 0.0:
            current_effective_font_size *= (1.0 + (char_pitch - reference_pitch) * pitch_size_scale)
        if velocity_size_scale != 0.0:
            current_effective_font_size *= (1.0 + (char_velocity - reference_velocity) * velocity_size_scale)
        effective_font_size_int = int(max(min_char_render_size, min(max_char_render_size, current_effective_font_size)))

        font_cache_key = (font_path, effective_font_size_int)
        font_obj = font_object_cache.get(font_cache_key)
        if not font_obj:
            try: font_obj = ImageFont.truetype(font_path, effective_font_size_int)
            except: font_obj = ImageFont.load_default(size=effective_font_size_int) if effective_font_size_int > 0 else ImageFont.load_default()
            font_object_cache[font_cache_key] = font_obj
        
        seg_actual_text_width = _get_text_width(temp_draw, text_to_render_for_layout, font_obj)
        total_calculated_width += seg_actual_text_width

        is_last_segment_in_line = (idx == num_segments - 1)

        if not is_last_segment_in_line: 
            # This segment's own duration-based padding
            current_segment_trailing_padding = 0.0
            if char_duration_ticks > duration_padding_threshold_ticks and duration_padding_scale_per_tick != 0.0:
                excess_duration_ticks = char_duration_ticks - duration_padding_threshold_ticks
                current_segment_trailing_padding = excess_duration_ticks * duration_padding_scale_per_tick
            current_segment_trailing_padding = max(0, int(current_segment_trailing_padding))
            total_calculated_width += current_segment_trailing_padding
        
        if not is_last_segment_in_line and num_segments > 1: 
            total_calculated_width += char_spacing # Inter-segment spacing
            
    layout_info['total_width_for_alignment'] = total_calculated_width

    if line_h_align == "left": layout_info['line_start_x_on_canvas'] = float(line_anchor_x_abs)
    elif line_h_align == "center": layout_info['line_start_x_on_canvas'] = float(line_anchor_x_abs) - total_calculated_width / 2.0
    elif line_h_align == "right": layout_info['line_start_x_on_canvas'] = float(line_anchor_x_abs) - total_calculated_width
    else: layout_info['line_start_x_on_canvas'] = float(line_anchor_x_abs) - total_calculated_width / 2.0
        
    return layout_info

def parse_dynamic_segment(raw_text: str) -> Dict[str, Any]:
    output = {'original_segment_text': raw_text, 'is_dynamic': False, 'sub_segments_timed': [], 'text_for_layout': raw_text.strip()}
    stripped_text = raw_text.strip() # Keep original raw_text for 'original_segment_text'

    if stripped_text.startswith("```") and stripped_text.endswith("```") and len(stripped_text) >= 6:
        literal_text = stripped_text[3:-3]
        output['text_for_layout'] = literal_text
        output['sub_segments_timed'] = [(literal_text, 0.0, 1.0)]
        return output

    if stripped_text == "---":
        output['is_dynamic'] = True; output['text_for_layout'] = ""
        output['sub_segments_timed'] = [("", 0.0, 1.0)]; return output

    parts = stripped_text.split('|'); final_timed_segments = []
    
    if len(parts) > 1: 
        output['is_dynamic'] = True
        num_sequential_slots = len(parts)
        slot_duration_ratio = 1.0 / num_sequential_slots if num_sequential_slots > 0 else 1.0

        for i, part_text_iter in enumerate(parts):
            part_text = part_text_iter # No strip here initially
            part_start_ratio = i * slot_duration_ratio
            part_end_ratio = (i + 1) * slot_duration_ratio
            
            if part_text.startswith("---") and len(part_text) > 3: # e.g. "---abc" or "---" for an empty progressive
                progressive_content = part_text[3:]
                num_progressive_steps = len(progressive_content)
                if num_progressive_steps == 0: # Case "---" within a sequence like "A|---|B"
                     final_timed_segments.append(("", part_start_ratio, part_end_ratio))
                else:
                    step_duration_ratio = (part_end_ratio - part_start_ratio) / num_progressive_steps
                    for j in range(num_progressive_steps):
                        sub_text = progressive_content[:j+1]
                        sub_start_ratio = part_start_ratio + j * step_duration_ratio
                        sub_end_ratio = part_start_ratio + (j + 1) * step_duration_ratio
                        final_timed_segments.append((sub_text, sub_start_ratio, sub_end_ratio))
            elif part_text == "---": # part is just "---", same as ---progressive_content when content is empty
                 final_timed_segments.append(("", part_start_ratio, part_end_ratio))
            else: # Static part in a sequence
                final_timed_segments.append((part_text, part_start_ratio, part_end_ratio))
        
        output['sub_segments_timed'] = final_timed_segments
        # For layout purposes, use the text of the final sub-segment of a dynamic sequence.
        # If it's something like "A|B|---CDE", text_for_layout becomes "CDE".
        # If it's "A|B", text_for_layout becomes "B".
        # If it's "A|---", text_for_layout becomes "".
        output['text_for_layout'] = final_timed_segments[-1][0] if final_timed_segments else ""


    elif stripped_text.startswith("---") and len(stripped_text) > 3: # Single segment, progressive "---abc"
        output['is_dynamic'] = True
        progressive_content = stripped_text[3:]
        num_progressive_steps = len(progressive_content)
        if num_progressive_steps == 0: # Should not happen if len > 3, but defensive
            output['text_for_layout'] = ""
            output['sub_segments_timed'] = [("",0.0,1.0)]
        else:
            step_duration_ratio = 1.0 / num_progressive_steps
            for j in range(num_progressive_steps):
                sub_text = progressive_content[:j+1]
                sub_start_ratio = j * step_duration_ratio
                sub_end_ratio = (j + 1) * step_duration_ratio
                final_timed_segments.append((sub_text, sub_start_ratio, sub_end_ratio))
            output['sub_segments_timed'] = final_timed_segments
            output['text_for_layout'] = final_timed_segments[-1][0] if final_timed_segments else ""
    else: # Static single segment (includes raw_text = "" which becomes stripped_text = "")
        # text_for_layout was already set to stripped_text at the beginning.
        output['sub_segments_timed'] = [(stripped_text, 0.0, 1.0)]
        # output['text_for_layout'] is already stripped_text

    # Normalize timings to ensure continuity and full 0-1 span for the segment
    if output['sub_segments_timed']:
        current_end_time = 0.0
        for k_seg in range(len(output['sub_segments_timed'])):
            text, s_ratio_orig, e_ratio_orig = output['sub_segments_timed'][k_seg]
            # Use original s_ratio as the start of the slot, ensure it's not before previous end
            actual_start_ratio = max(current_end_time, s_ratio_orig) 
            actual_end_ratio = e_ratio_orig # Use original end ratio of the slot
            
            # If it's the very last sub-segment overall for this dynamic part, force it to 1.0
            if k_seg == len(output['sub_segments_timed']) - 1:
                actual_end_ratio = 1.0
            
            actual_end_ratio = max(actual_start_ratio, actual_end_ratio) # Ensure end is not before start

            output['sub_segments_timed'][k_seg] = (text, actual_start_ratio, actual_end_ratio)
            current_end_time = actual_end_ratio
        
        # Final pass to ensure first starts at 0.0 and last ends at 1.0 if list is not empty
        if output['sub_segments_timed']:
             f_text, _, f_end = output['sub_segments_timed'][0]
             output['sub_segments_timed'][0] = (f_text, 0.0, f_end)
             
             # Adjust all subsequent start times to be continuous if they are not already
             for k_fix_cont in range(len(output['sub_segments_timed']) - 1):
                 curr_seg_text, curr_seg_start, curr_seg_end = output['sub_segments_timed'][k_fix_cont]
                 next_seg_text, next_seg_start_orig, next_seg_end_orig = output['sub_segments_timed'][k_fix_cont+1]
                 if next_seg_start_orig < curr_seg_end: # If next starts before current ends, make it continuous
                     output['sub_segments_timed'][k_fix_cont+1] = (next_seg_text, curr_seg_end, max(curr_seg_end, next_seg_end_orig))
             
             l_text, l_start, _ = output['sub_segments_timed'][-1]
             output['sub_segments_timed'][-1] = (l_text, l_start, 1.0) # Force last to end at 1.0

    return output

def generate_lyric_video_v2(
    midi_path: str, lyrics_path: str, output_video_path: str, font_path: str, 
    width: int = 1920, height: int = 1080, fps: int = 30,
    font_size_base_param: int = 60, 
    char_spacing: int = 10, bg_color: tuple = (0,0,0), text_color: tuple = (255,255,255),
    text_vertical_align: str = "center", 
    line_placement_mode: str = "dynamic", 
    line_h_align: str = "center",      
    line_anchor_x: Optional[int] = None, 
    line_anchor_y: Optional[int] = None, 
    pitch_offset_scale: float = 0.0,
    pitch_size_scale: float = 0.0,
    velocity_size_scale: float = 0.0,
    reference_pitch: int = 60, 
    reference_velocity: int = 64,
    duration_padding_threshold_ticks: int = 240, 
    duration_padding_scale_per_tick: float = 0.1,
    min_char_render_size: int = 8,
    max_char_render_size: int = 300,
    logger: ILogger = PrintLogger(), progress_callback: Optional[callable] = None
):
    logger.info(f"動画生成プロセス開始: MIDI='{os.path.basename(midi_path)}', Output='{os.path.basename(output_video_path)}'")
    # (MIDI loading and initial setup remains the same as previous correct version)
    note_events_with_full_duration_info = [] 
    total_midi_duration_sec = 0.0; ticks_per_beat_from_midi = 480
    try:
        mid = mido.MidiFile(midi_path)
        if mid.ticks_per_beat: ticks_per_beat_from_midi = mid.ticks_per_beat
        logger.info(f"MIDI Ticks Per Beat: {ticks_per_beat_from_midi}")
        timed_notes = []
        current_tempo = 500000
        for track_idx, track in enumerate(mid.tracks):
            abs_time_sec = 0; abs_time_tick = 0
            active_notes_on_track = {} # (pitch) -> {start_sec, start_tick, velocity}
            for msg in track:
                delta_sec = mido.tick2second(msg.time, ticks_per_beat_from_midi, current_tempo)
                abs_time_sec += delta_sec; abs_time_tick += msg.time
                if msg.is_meta and msg.type == 'set_tempo': current_tempo = msg.tempo
                elif msg.type == 'note_on' and msg.velocity > 0:
                    active_notes_on_track[msg.note] = {
                        'start_sec': abs_time_sec, 'start_tick': abs_time_tick, 
                        'velocity': msg.velocity, 'pitch': msg.note }
                elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
                    if msg.note in active_notes_on_track:
                        note_on_data = active_notes_on_track.pop(msg.note)
                        timed_notes.append({
                            'time_sec': note_on_data['start_sec'],
                            'duration_sec': max(0.01, abs_time_sec - note_on_data['start_sec']),
                            'duration_ticks': max(1, abs_time_tick - note_on_data['start_tick']),
                            'pitch': note_on_data['pitch'], 'velocity': note_on_data['velocity'] })
        timed_notes.sort(key=lambda x: x['time_sec'])
        note_events_with_full_duration_info = timed_notes
        total_midi_duration_sec = mid.length
        if not note_events_with_full_duration_info: # Fallback
             logger.warning("MIDI解析でノートイベントが見つかりませんでした。mid.play()ベースのフォールバックを試みます。")
             temp_fb_notes = []
             current_time_fb = 0.0
             last_event_time_fb = 0.0
             # Corrected fallback:
             # Use a temporary list to store note_on events from mid.play()
             play_note_ons = []
             for msg_p in mid.play(): 
                current_time_fb += msg_p.time
                if msg_p.type == 'note_on' and msg_p.velocity > 0:
                    play_note_ons.append({'time_sec': current_time_fb,'velocity': msg_p.velocity,'pitch': msg_p.note})
                last_event_time_fb = current_time_fb
             
             if play_note_ons:
                total_midi_duration_sec = max(total_midi_duration_sec, last_event_time_fb)
                for i, ev_p in enumerate(play_note_ons):
                    duration_s_fb = 0.2 # Default duration
                    if i < len(play_note_ons) - 1:
                        duration_s_fb = play_note_ons[i+1]['time_sec'] - ev_p['time_sec']
                    elif total_midi_duration_sec > ev_p['time_sec']:
                         duration_s_fb = total_midi_duration_sec - ev_p['time_sec']
                    
                    avg_tempo_fb = 500000 
                    try: 
                        tempos = [m.tempo for t_ in mid.tracks for m in t_ if m.is_meta and m.type == 'set_tempo']
                        if tempos: avg_tempo_fb = sum(tempos) / len(tempos)
                    except: pass
                    duration_t_fb = int(mido.second2tick(max(0.01, duration_s_fb), ticks_per_beat_from_midi, avg_tempo_fb))

                    note_events_with_full_duration_info.append({
                        'time_sec': ev_p['time_sec'], 'velocity': ev_p['velocity'],
                        'pitch': ev_p['pitch'], 
                        'duration_sec': max(0.01, duration_s_fb),
                        'duration_ticks': max(1, duration_t_fb if duration_t_fb > 0 else int(ticks_per_beat_from_midi * 0.1))
                    })
             else:
                 logger.error("MIDIノートイベントを抽出できませんでした。")
    except Exception as e: logger.error(f"MIDIファイル処理エラー: {e}\n{traceback.format_exc()}"); return

    font_size_base_for_line = _calculate_adjusted_font_size(font_size_base_param, total_midi_duration_sec)
    logger.info(f"基本フォントサイズ(楽曲長調整後): {font_size_base_param} -> {font_size_base_for_line} (曲長: {total_midi_duration_sec:.2f}s)")

    parsed_lyrics = []
    try:
        with open(lyrics_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip(): parsed_lyrics.append(line.strip().split('/'))
                else: parsed_lyrics.append([]) # Keep empty lines as empty lists if they were non-whitespace
    except Exception as e: logger.error(f"歌詞ファイル '{lyrics_path}' 読込エラー: {e}"); return
    if not parsed_lyrics: logger.warning("歌詞ファイルが空か解析可能な行がありません。")

    final_events = []; note_ptr = 0
    if not note_events_with_full_duration_info and parsed_lyrics: logger.warning("MIDIノート不足(イベント生成不可)。")

    for line_idx, line_raw_segments in enumerate(parsed_lyrics):
        if not note_events_with_full_duration_info or note_ptr >= len(note_events_with_full_duration_info):
            if note_events_with_full_duration_info and parsed_lyrics: logger.warning(f"歌詞行 {line_idx+1} 以降MIDIノート不足。"); break
        
        # Add clear_line event at the start time of the first note intended for this line
        # If the line is empty (line_raw_segments is empty), this clear_line might be redundant or closely followed by next line's clear.
        # This is generally okay.
        if note_ptr < len(note_events_with_full_duration_info):
            first_note_time_for_line = note_events_with_full_duration_info[note_ptr]['time_sec']
            # To prevent adding clear_line for completely empty lyric lines that have no segments
            # Add clear_line only if this line has segments or is not the first line (to clear previous)
            if line_raw_segments or line_idx > 0 or (line_idx == 0 and parsed_lyrics): # ensure it's a meaningful line
                 final_events.append({'time': first_note_time_for_line, 'type': 'clear_line', 'data': {'line_idx': line_idx}})

        segment_idx_in_line = 0
        for raw_segment_text in line_raw_segments:
            if note_ptr >= len(note_events_with_full_duration_info): 
                logger.warning(f"行 {line_idx+1} の途中でMIDIノート不足。セグメント '{raw_segment_text}' 以降スキップ。"); break
            
            current_note = note_events_with_full_duration_info[note_ptr]
            parsed_dynamic_info = parse_dynamic_segment(raw_segment_text)
            
            event_data_for_char = {
                'original_segment_text': raw_segment_text, 
                'text_for_layout': parsed_dynamic_info['text_for_layout'], 
                'is_dynamic': parsed_dynamic_info['is_dynamic'], 
                'sub_segments_timed': parsed_dynamic_info['sub_segments_timed'],
                'velocity': current_note['velocity'], 
                'line_idx': line_idx, 
                'segment_idx_in_line': segment_idx_in_line, 
                'pitch': current_note['pitch'], 
                'duration_ticks': current_note['duration_ticks'],
                'note_start_time_sec': current_note['time_sec'], 
                'note_duration_sec': current_note['duration_sec']
            }
            
            # MODIFICATION: Always create a 'char' event for every segment defined in the lyrics file
            # (i.e., parts separated by '/'), as long as a MIDI note is available.
            # An empty raw_segment_text (from "//") will result in a 'char' event
            # with text_for_layout = "" via parse_dynamic_segment.
            final_events.append({
                'time': current_note['time_sec'],
                'type': 'char',
                'data': event_data_for_char
            })
            segment_idx_in_line += 1
            note_ptr += 1
            
    # (Final event sorting, video duration, and VideoWriter setup remains same as previous correct version)
    if final_events:
        max_line_idx = -1; last_lyric_time = 0
        non_clear_evs = [ev for ev in final_events if ev['type'] != 'clear_line']
        if non_clear_evs: last_lyric_time = max(ev['time'] for ev in non_clear_evs if 'time' in ev) if non_clear_evs else 0

        # Determine max_line_idx from all events that have line_idx
        all_line_indices_in_events = [ev['data']['line_idx'] for ev in final_events if 'data' in ev and 'line_idx' in ev['data']]
        if all_line_indices_in_events:
            max_line_idx = max(all_line_indices_in_events)
        
        final_clear_time = max(last_lyric_time + 0.5, total_midi_duration_sec + 0.1) 
        if max_line_idx >=0 : final_events.append({'time': final_clear_time, 'type': 'clear_line', 'data': {'line_idx': max_line_idx + 1}}) # Clear one line beyond the max used
    
    final_events.sort(key=lambda x: (x['time'], 0 if x['type'] == 'clear_line' else 1)) # Ensure clear_line events are processed first at same timestamp
    
    video_total_duration_sec_final = total_midi_duration_sec 
    video_total_frames = int(video_total_duration_sec_final * fps)
    if video_total_frames <= 0:
        if total_midi_duration_sec <=0 and not note_events_with_full_duration_info: logger.error(f"動画フレーム0以下(MIDI長: {total_midi_duration_sec:.2f}s)"); return
        elif total_midi_duration_sec <=0 and note_events_with_full_duration_info: logger.warning(f"MIDI演奏時間0秒だがノートは存在。")
        if final_events: 
            max_event_time = max(ev['time'] for ev in final_events if 'time' in ev) if final_events else 0
            video_total_duration_sec_final = max(video_total_duration_sec_final, max_event_time + 0.5) 
            video_total_frames = int(video_total_duration_sec_final * fps)
            if video_total_frames < 1 and (note_events_with_full_duration_info or parsed_lyrics): video_total_frames = 1 
            logger.info(f"最終イベント時間に基づき動画長を調整: {video_total_duration_sec_final:.2f}s, フレーム: {video_total_frames}")
        if video_total_frames <=0 : logger.error(f"最終的な動画総フレーム数が0以下。中止。"); return

    output_dir = os.path.dirname(output_video_path)
    if output_dir and not os.path.exists(output_dir):
        try: os.makedirs(output_dir); logger.info(f"出力ディレクトリ '{output_dir}' 作成。")
        except Exception as e_mkdir: logger.error(f"出力ディレクトリ '{output_dir}' 作成失敗: {e_mkdir}")

    fourcc_str_options = ['avc1', 'X264', 'H264', 'mp4v']
    try: fourcc_val_options = [cv2.VideoWriter_fourcc(*s) for s in fourcc_str_options]
    except AttributeError: logger.error("cv2.VideoWriter_fourcc 利用不可。"); return

    video_writer = None;
    for fcc_val, fcc_str in zip(fourcc_val_options, fourcc_str_options):
        try:
            video_writer_test = cv2.VideoWriter(output_video_path, fcc_val, float(fps), (width, height))
            if video_writer_test.isOpened(): video_writer = video_writer_test; logger.info(f"動画ライター初期化成功 (FourCC: {fcc_str})"); break
            else: logger.warning(f"FourCC '{fcc_str}' 初期化失敗。"); video_writer_test.release()
        except Exception as e_init: logger.warning(f"FourCC '{fcc_str}' 初期化中エラー: {e_init}。"); video_writer_test.release() # Release on exception too
    if not video_writer or not video_writer.isOpened(): logger.error(f"全FourCC ({', '.join(fourcc_str_options)}) で動画ライター初期化失敗。"); return
    
    try: _ = ImageFont.truetype(font_path, size=10) 
    except Exception as e: logger.error(f"フォント '{font_path}' 問題: {e}"); video_writer.release(); return
    
    actual_line_anchor_x = line_anchor_x if line_anchor_x is not None else width // 2
    actual_line_anchor_y = line_anchor_y if line_anchor_y is not None else height // 2
    
    current_event_ptr = 0; current_line_idx_on_screen = -1
    active_segment_event_data = [] # Stores full event_data for segments on the current line
    num_max_segments_in_current_line = 0

    fixed_layout_cache: Dict[int, Dict[str, Any]] = {}
    font_object_cache: Dict[Tuple[str, int], ImageFont.FreeTypeFont] = {}
    logger.info(f"動画生成ループ開始: {output_video_path} ({width}x{height} @ {fps}fps, 総フレーム: {video_total_frames})")

    for frame_num in range(video_total_frames):
        current_video_time = frame_num / float(fps)
        while current_event_ptr < len(final_events) and final_events[current_event_ptr]['time'] <= current_video_time:
            event = final_events[current_event_ptr]; event_data_from_final_event = event['data']
            if event['type'] == 'clear_line':
                clear_line_target_idx = event_data_from_final_event['line_idx']
                # Only update if clearing a "future" line or the current one being shown (or first line)
                if clear_line_target_idx > current_line_idx_on_screen or current_line_idx_on_screen == -1 :
                    current_line_idx_on_screen = clear_line_target_idx
                    # Get num_max_segments from parsed_lyrics, ensuring line_idx is valid
                    if 0 <= current_line_idx_on_screen < len(parsed_lyrics):
                        num_max_segments_in_current_line = len(parsed_lyrics[current_line_idx_on_screen])
                    else: # This happens for the final clear_line (max_line_idx + 1)
                        num_max_segments_in_current_line = 0
                    
                    active_segment_event_data = [None] * num_max_segments_in_current_line # Reset for new line
                    
                    if line_placement_mode == "fixed" and current_line_idx_on_screen not in fixed_layout_cache and num_max_segments_in_current_line > 0:
                        segment_event_data_list_for_layout = []
                        # Collect all 'char' events for this specific line_idx for fixed layout calculation
                        char_events_for_this_line_raw = [
                            ev_s['data'] for ev_s in final_events 
                            if ev_s['type'] == 'char' and 'data' in ev_s and ev_s['data'].get('line_idx') == current_line_idx_on_screen
                        ]
                        # Sort them by their segment_idx_in_line to ensure correct order
                        char_events_for_this_line_raw.sort(key=lambda x: x.get('segment_idx_in_line', float('inf')))
                        
                        for ed_raw in char_events_for_this_line_raw:
                            segment_event_data_list_for_layout.append({
                                'text': ed_raw['text_for_layout'], 'pitch': ed_raw['pitch'], 
                                'velocity': ed_raw['velocity'], 'duration_ticks': ed_raw['duration_ticks']})
                        
                        if segment_event_data_list_for_layout: # Only calculate if there are segments
                            logger.info(f"固定レイアウト計算中 Line {current_line_idx_on_screen}: Segments for layout: {len(segment_event_data_list_for_layout)}")
                            layout_info = _calculate_fixed_layout_for_line_v2(
                                segment_event_data_list_for_layout, font_path, font_size_base_for_line, char_spacing,
                                width, line_h_align, actual_line_anchor_x, pitch_size_scale, velocity_size_scale,
                                reference_pitch, reference_velocity, duration_padding_threshold_ticks, 
                                duration_padding_scale_per_tick, min_char_render_size, max_char_render_size,
                                font_object_cache, logger )
                            fixed_layout_cache[current_line_idx_on_screen] = layout_info
                        elif num_max_segments_in_current_line > 0 : # Line has segments in lyrics, but no char events (e.g. MIDI ran out)
                            logger.warning(f"固定レイアウト計算 Line {current_line_idx_on_screen}: 歌詞セグメントあり ({num_max_segments_in_current_line}) だがノートイベントなし。空レイアウト作成。")
                            fixed_layout_cache[current_line_idx_on_screen] = _calculate_fixed_layout_for_line_v2(
                                [], font_path, font_size_base_for_line, char_spacing, width, line_h_align, actual_line_anchor_x, 
                                pitch_size_scale, velocity_size_scale, reference_pitch, reference_velocity, 
                                duration_padding_threshold_ticks, duration_padding_scale_per_tick, 
                                min_char_render_size, max_char_render_size, font_object_cache, logger)


            elif event['type'] == 'char':
                # Process 'char' event only if its line_idx matches the one currently supposed to be on screen
                if event_data_from_final_event['line_idx'] == current_line_idx_on_screen:
                    seg_idx = event_data_from_final_event['segment_idx_in_line'] 
                    if 0 <= seg_idx < num_max_segments_in_current_line: 
                        active_segment_event_data[seg_idx] = event_data_from_final_event
                    # else: logger.warning(f"Frame {frame_num}: Char event seg_idx {seg_idx} out of bounds for line {current_line_idx_on_screen} (max_segs: {num_max_segments_in_current_line})")
            current_event_ptr += 1

        image = Image.new('RGB', (width, height), color=bg_color); draw = ImageDraw.Draw(image)
        current_line_prepared_segments_for_draw: List[Optional[Dict[str, Any]]] = [None] * num_max_segments_in_current_line
        idx_of_last_segment_to_draw_this_frame = -1 # Index of the last segment that has *any* render info for this frame

        if num_max_segments_in_current_line > 0 :
            for i in range(num_max_segments_in_current_line):
                full_event_data_for_segment_i = active_segment_event_data[i] # This is the original dict from final_events
                if full_event_data_for_segment_i:
                    idx_of_last_segment_to_draw_this_frame = i # Track the last segment whose event has fired
                    text_to_render_this_frame = ""
                    if full_event_data_for_segment_i['is_dynamic']:
                        note_start_s = full_event_data_for_segment_i['note_start_time_sec']; note_dur_s = full_event_data_for_segment_i['note_duration_sec']
                        if note_dur_s <= 1e-6: text_to_render_this_frame = full_event_data_for_segment_i['sub_segments_timed'][-1][0] if full_event_data_for_segment_i['sub_segments_timed'] else ""
                        else:
                            time_into_note = current_video_time - note_start_s
                            progress_ratio = max(0.0, min(1.0, time_into_note / note_dur_s))
                            matched_text_found = False
                            for txt_part, s_ratio, e_ratio in full_event_data_for_segment_i['sub_segments_timed']:
                                # Exact match for 0 progress at 0 start, 0 end
                                if s_ratio == 0.0 and e_ratio == 0.0 and progress_ratio == 0.0 : text_to_render_this_frame = txt_part; matched_text_found = True; break
                                # Interval check: s_ratio <= progress < e_ratio (typical case)
                                # Make e_ratio inclusive for the last segment if progress_ratio is 1.0
                                if s_ratio <= progress_ratio < e_ratio: text_to_render_this_frame = txt_part; matched_text_found = True; break
                                if e_ratio == 1.0 and progress_ratio == 1.0 : text_to_render_this_frame = txt_part; matched_text_found = True; break # Explicitly handle end of segment
                            if not matched_text_found and full_event_data_for_segment_i['sub_segments_timed']:
                                if progress_ratio >= 1.0: text_to_render_this_frame = full_event_data_for_segment_i['sub_segments_timed'][-1][0]
                                elif progress_ratio <= 0.0: text_to_render_this_frame = full_event_data_for_segment_i['sub_segments_timed'][0][0]
                                else: # Fallback if no match, take last subsegment's text (e.g. if timing slightly off)
                                     text_to_render_this_frame = full_event_data_for_segment_i['sub_segments_timed'][-1][0]

                    else: text_to_render_this_frame = full_event_data_for_segment_i['text_for_layout']
                    
                    char_pitch = full_event_data_for_segment_i['pitch']; char_velocity = full_event_data_for_segment_i['velocity']
                    eff_font_size = float(font_size_base_for_line) 
                    if pitch_size_scale != 0.0: eff_font_size *= (1.0 + (char_pitch - reference_pitch) * pitch_size_scale)
                    if velocity_size_scale != 0.0: eff_font_size *= (1.0 + (char_velocity - reference_velocity) * velocity_size_scale)
                    eff_font_size_int = int(max(min_char_render_size, min(max_char_render_size, eff_font_size)))

                    font_key = (font_path, eff_font_size_int)
                    font_obj = font_object_cache.get(font_key)
                    if not font_obj:
                        try: font_obj = ImageFont.truetype(font_path, eff_font_size_int)
                        except: font_obj = ImageFont.load_default(size=eff_font_size_int) if eff_font_size_int > 0 else ImageFont.load_default()
                        font_object_cache[font_key] = font_obj
                    
                    temp_draw_metrics = ImageDraw.Draw(Image.new('RGB',(1,1))) # Small image for metrics
                    seg_actual_text_width = _get_text_width(temp_draw_metrics, text_to_render_this_frame, font_obj)
                    seg_actual_height = 0; bbox_top_offset = 0
                    if text_to_render_this_frame: # Only get bbox if text exists
                        try: 
                            bbox = temp_draw_metrics.textbbox((0,0), text_to_render_this_frame, font=font_obj) 
                            seg_actual_height = bbox[3] - bbox[1]; bbox_top_offset = bbox[1] 
                        except: # Fallback
                            try: _, legacy_h = temp_draw_metrics.textsize(text_to_render_this_frame,font=font_obj); asc, desc = font_obj.getmetrics(); seg_actual_height = asc+desc; bbox_top_offset = -asc
                            except: seg_actual_height = eff_font_size_int; bbox_top_offset = -int(eff_font_size_int * 0.8)
                    
                    # Calculate this segment's own duration-based trailing padding
                    current_segment_trailing_padding = 0.0
                    char_dur_ticks = full_event_data_for_segment_i['duration_ticks']
                    if char_dur_ticks > duration_padding_threshold_ticks and duration_padding_scale_per_tick != 0.0:
                        current_segment_trailing_padding = (char_dur_ticks - duration_padding_threshold_ticks) * duration_padding_scale_per_tick
                    current_segment_trailing_padding = max(0, int(current_segment_trailing_padding))

                    current_line_prepared_segments_for_draw[i] = {
                        'text': text_to_render_this_frame, 'font_obj': font_obj,
                        'actual_text_width': seg_actual_text_width, # Width of current text being rendered
                        'actual_height': seg_actual_height, 
                        'bbox_top_offset': bbox_top_offset, 'pitch': char_pitch, 
                        'segment_trailing_padding': current_segment_trailing_padding # Padding derived from this segment's note duration
                    }
        
        current_x_to_draw = 0.0
        # num_drawable_segments_this_frame is the count of segments that have *any* render info prepared
        # (i.e., their event has fired and they are part of the active_segment_event_data for this line)
        num_drawable_segments_this_frame = sum(1 for s_prep in current_line_prepared_segments_for_draw if s_prep is not None)
        
        if num_drawable_segments_this_frame > 0: # Only proceed if there's something to potentially draw or space out
            if line_placement_mode == "dynamic":
                current_dynamic_line_total_width_for_alignment = 0.0
                # Iterate up to the last segment that has *actually* appeared so far (idx_of_last_segment_to_draw_this_frame)
                # This ensures dynamic layout correctly adjusts as segments appear one by one.
                for k_idx_dyn in range(idx_of_last_segment_to_draw_this_frame + 1):
                    seg_info_dyn = current_line_prepared_segments_for_draw[k_idx_dyn]
                    if seg_info_dyn: # If it's a segment to be rendered (even if text is "")
                        current_dynamic_line_total_width_for_alignment += seg_info_dyn['actual_text_width']
                        # Apply segment's own padding to its width for alignment if not last segment of entire line
                        # AND if it's not the very last segment being drawn in this dynamic appearance sequence.
                        if k_idx_dyn < num_max_segments_in_current_line - 1: # Not last segment of the *full* line definition
                             current_dynamic_line_total_width_for_alignment += seg_info_dyn['segment_trailing_padding']
                        
                        # Add char spacing if not the last segment being drawn *in this dynamic appearance sequence*
                        if k_idx_dyn < idx_of_last_segment_to_draw_this_frame :
                             current_dynamic_line_total_width_for_alignment += char_spacing
                
                if line_h_align == "left": current_x_to_draw = float(actual_line_anchor_x)
                elif line_h_align == "center": current_x_to_draw = float(actual_line_anchor_x) - current_dynamic_line_total_width_for_alignment / 2.0
                elif line_h_align == "right": current_x_to_draw = float(actual_line_anchor_x) - current_dynamic_line_total_width_for_alignment
                else: current_x_to_draw = float(actual_line_anchor_x) - current_dynamic_line_total_width_for_alignment / 2.0 
            
            elif line_placement_mode == "fixed":
                layout_info_for_fixed = fixed_layout_cache.get(current_line_idx_on_screen)
                if layout_info_for_fixed: current_x_to_draw = layout_info_for_fixed['line_start_x_on_canvas']
                else: # Fallback if fixed layout somehow not cached (e.g., line has no segments in lyrics but clear_line occurred)
                    if line_h_align == "left": current_x_to_draw = float(actual_line_anchor_x)
                    elif line_h_align == "center": current_x_to_draw = float(actual_line_anchor_x) # Empty line centered at anchor
                    elif line_h_align == "right": current_x_to_draw = float(actual_line_anchor_x) # Empty line right-aligned to anchor
                    else: current_x_to_draw = float(actual_line_anchor_x)


            # Actual drawing pass: Iterate only up to the last segment that has appeared
            for i_draw in range(idx_of_last_segment_to_draw_this_frame + 1):
                render_info = current_line_prepared_segments_for_draw[i_draw]
                if render_info: # If segment is active and prepared (its event has fired)
                    baseline_y_ref = float(actual_line_anchor_y); y_pixel_offset = 0.0
                    if pitch_offset_scale != 0.0: y_pixel_offset = (render_info['pitch'] - reference_pitch) * pitch_offset_scale * -1.0 
                    
                    if text_vertical_align == "center": y_draw = baseline_y_ref - (render_info['bbox_top_offset'] + render_info['actual_height'] / 2.0)
                    elif text_vertical_align == "top": y_draw = baseline_y_ref - render_info['bbox_top_offset']
                    elif text_vertical_align == "bottom": y_draw = baseline_y_ref - (render_info['bbox_top_offset'] + render_info['actual_height'])
                    else: y_draw = baseline_y_ref # Baseline alignment
                    
                    final_draw_y_for_pil = y_draw + y_pixel_offset
                    
                    segment_text_this_frame = render_info['text'] # Text currently visible this frame
                    font_obj = render_info['font_obj']
                    padding_for_this_segment_note = render_info['segment_trailing_padding'] 
                    
                    # Padding should apply if this segment is NOT the last one in the *entire line definition*
                    apply_this_segment_padding_visually = (i_draw < num_max_segments_in_current_line - 1)
                    effective_padding_to_use_for_segment = padding_for_this_segment_note if apply_this_segment_padding_visually else 0.0

                    original_event_data_for_segment = active_segment_event_data[i_draw] # Should be valid if render_info is valid

                    if segment_text_this_frame: 
                        text_basis_for_padding_distribution = segment_text_this_frame 
                        if original_event_data_for_segment: # Should always be true here
                            is_dynamic_segment = original_event_data_for_segment.get('is_dynamic', False)
                            # Use text_for_layout for padding distribution if it's a progressive segment
                            # to ensure padding is distributed over the full final text form.
                            original_raw_text_stripped = original_event_data_for_segment.get('original_segment_text', "").strip()
                            is_progressive_type = (
                                is_dynamic_segment and 
                                original_raw_text_stripped.startswith("---") and 
                                not (original_raw_text_stripped.startswith("```") and original_raw_text_stripped.endswith("```"))
                            )
                            if is_progressive_type:
                                text_basis_for_padding_distribution = original_event_data_for_segment.get('text_for_layout', segment_text_this_frame)
                        
                        num_chars_for_padding_divisor = len(text_basis_for_padding_distribution)
                        padding_per_char_distributed = 0.0
                        if num_chars_for_padding_divisor > 0 and effective_padding_to_use_for_segment > 0:
                            padding_per_char_distributed = effective_padding_to_use_for_segment / num_chars_for_padding_divisor
                        
                        for char_visual_idx, char_visual in enumerate(segment_text_this_frame):
                            draw.text( (current_x_to_draw, final_draw_y_for_pil), char_visual, font=font_obj, fill=text_color )
                            char_width = _get_text_width(draw, char_visual, font_obj) 
                            current_x_to_draw += char_width
                            # Distribute padding after each character of the current segment's text
                            # This applies only if the segment itself is eligible for padding (not last in line)
                            # and it has text.
                            current_x_to_draw += padding_per_char_distributed
                            
                    elif effective_padding_to_use_for_segment > 0: # Empty text, but padding applies as a block
                        current_x_to_draw += effective_padding_to_use_for_segment
                    
                    # Add inter-segment char_spacing if this is not the last segment *currently being drawn this frame*
                    if i_draw < idx_of_last_segment_to_draw_this_frame: 
                        current_x_to_draw += char_spacing
        
        frame_bgr = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        if frame_bgr is None or frame_bgr.shape[0] != height or frame_bgr.shape[1] != width:
            logger.error(f"フレームサイズ不正。期待({width}x{height}), 実際:{frame_bgr.shape if frame_bgr is not None else 'None'}"); video_writer.release(); return 
        try: video_writer.write(frame_bgr)
        except Exception as e_write: logger.error(f"フレーム {frame_num} 書き込みエラー: {e_write}"); video_writer.release(); return 
        
        if progress_callback: progress_callback(frame_num + 1, video_total_frames)
        if (frame_num + 1) % (fps * 10) == 0: logger.info(f"処理中: { (frame_num + 1) / float(fps):.1f}s / {video_total_duration_sec_final:.1f}s")

    if video_writer: 
        video_writer.release(); logger.info(f"動画 '{output_video_path}' 書き込み終了。")
        try:
            file_size = os.path.getsize(output_video_path)
            logger.info(f"生成ファイルサイズ: {file_size} バイト")
            if file_size < 1024 and video_total_frames > 0 : logger.warning(f"ファイルサイズ極小 ({file_size} バイト)。")
        except OSError as e_size: logger.warning(f"ファイルサイズ確認失敗: {e_size}")
    else: logger.error("VideoWriter未初期化。")
    logger.info(f"動画生成プロセス完了: {output_video_path}")

# --- END OF VIDEO GENERATION LOGIC ---

# --- Custom Widgets for Drag & Drop ---
class LyricsTextEdit(QTextEdit):
    file_dropped = Signal(str); request_focus = Signal() 
    def __init__(self, parent=None): super().__init__(parent); self.setAcceptDrops(True)
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile() and urls[0].toLocalFile().lower().endswith(".txt"): event.acceptProposedAction(); return
        event.ignore()
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                self.file_dropped.emit(urls[0].toLocalFile()); self.request_focus.emit(); event.acceptProposedAction(); return
        event.ignore()

class DropTargetGroupBox(QGroupBox):
    file_dropped = Signal(str)
    def __init__(self, title: str, accepted_extensions: List[str], parent=None):
        super().__init__(title, parent); self.setAcceptDrops(True); self.accepted_extensions = [ext.lower() for ext in accepted_extensions]
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                if any(urls[0].toLocalFile().lower().endswith(ext) for ext in self.accepted_extensions): event.acceptProposedAction(); return
        event.ignore()
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile(): self.file_dropped.emit(urls[0].toLocalFile()); event.acceptProposedAction(); return
        event.ignore()



# --- START OF MODIFIED SECTION IN MidiLoadThread.run ---
class MidiLoadThread(QThread):
    finished = Signal(list, float, str, list, int) 
    def __init__(self, midi_path: str): super().__init__(); self.midi_path = midi_path
    def run(self):
        detailed_notes_for_roll: List[Dict[str, Any]] = []
        raw_note_events_for_mapping_with_full_info: List[Dict[str, Any]] = []
        total_duration_sec_for_video: float = 0.0
        error_msg: str = ""
        ticks_per_beat_from_midi: int = 480 
        
        try:
            mid = mido.MidiFile(self.midi_path)
            if mid.ticks_per_beat:
                ticks_per_beat_from_midi = mid.ticks_per_beat
            
            # --- Unified Note Processing with Overlap Adjustment ---
            all_processed_notes: List[Dict[str, Any]] = []
            current_tempo_global = 500000  # Global tempo tracking for conversions

            # Map to store the actual end time (tick and sec) of the last note processed 
            # for a given (track_idx, pitch). This is used for overlap adjustment.
            # Key: (track_idx, pitch), Value: {'tick': end_tick, 'sec': end_sec}
            last_note_actual_end_map: Dict[Tuple[int, int], Dict[str, float]] = {}

            for track_idx, track in enumerate(mid.tracks):
                abs_tick_track = 0  # Absolute ticks from the start of this track
                abs_sec_track = 0.0 # Absolute seconds from the start of this track
                
                # Stores active note_on events for the current track: pitch -> {data}
                active_notes_in_track: Dict[int, Dict[str, Any]] = {}
                current_tempo_track = 500000 # Tempo for current track, might differ from global if meta is per track

                # First pass on track to ensure correct tempo for initial delta ticks
                # This is a bit simplified; a full tempo map build before this loop is more robust for global time.
                # However, mido.tick2second uses the tempo *at the time of the delta*.
                initial_tempo_for_track_deltas = 500000 
                temp_abs_tick_for_tempo = 0
                for msg_check in track: # Check for initial tempo if any before first note
                    temp_abs_tick_for_tempo += msg_check.time
                    if msg_check.is_meta and msg_check.type == 'set_tempo':
                        initial_tempo_for_track_deltas = msg_check.tempo
                        break # Found first tempo for this track
                    if not msg_check.is_meta: # Stop if non-meta event is found before tempo
                        break
                current_tempo_track = initial_tempo_for_track_deltas


                for msg in track:
                    delta_ticks_msg = msg.time
                    # Use current_tempo_track for converting this specific delta
                    delta_sec_msg = mido.tick2second(delta_ticks_msg, ticks_per_beat_from_midi, current_tempo_track)
                    
                    abs_tick_track += delta_ticks_msg
                    abs_sec_track += delta_sec_msg

                    if msg.is_meta and msg.type == 'set_tempo':
                        current_tempo_track = msg.tempo
                        current_tempo_global = msg.tempo # Update global tempo as well
                    
                    elif msg.type == 'note_on' and msg.velocity > 0:
                        pitch = msg.note
                        
                        adjusted_start_tick = abs_tick_track
                        adjusted_start_sec = abs_sec_track
                        
                        track_pitch_key = (track_idx, pitch)
                        if track_pitch_key in last_note_actual_end_map:
                            prev_note_end_data = last_note_actual_end_map[track_pitch_key]
                            prev_end_tick = prev_note_end_data['tick']
                            prev_end_sec = prev_note_end_data['sec']

                            if abs_tick_track < prev_end_tick: # Overlap
                                adjusted_start_tick = prev_end_tick
                                adjusted_start_sec = prev_end_sec
                        
                        active_notes_in_track[pitch] = {
                            'pitch': pitch,
                            'velocity': msg.velocity,
                            'start_tick_abs_track': adjusted_start_tick, # Adjusted, relative to track start
                            'start_sec_abs_track': adjusted_start_sec,   # Adjusted, relative to track start
                            'tempo_at_note_on': current_tempo_track # Tempo at the moment of this note_on
                        }

                    elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
                        pitch = msg.note
                        if pitch in active_notes_in_track:
                            note_on_data = active_notes_in_track.pop(pitch)
                            
                            actual_start_tick = note_on_data['start_tick_abs_track']
                            actual_start_sec = note_on_data['start_sec_abs_track']
                            tempo_at_note_on = note_on_data['tempo_at_note_on']
                            
                            # End time is current message's time (abs_tick_track, abs_sec_track)
                            current_event_end_tick = abs_tick_track
                            current_event_end_sec = abs_sec_track

                            duration_ticks = current_event_end_tick - actual_start_tick
                            duration_sec = current_event_end_sec - actual_start_sec
                            
                            final_end_tick = current_event_end_tick
                            final_end_sec = current_event_end_sec

                            min_duration_ticks = 1 
                            # Use tempo at note_on for min_duration_sec calc for stability if start was pushed.
                            # Or current_tempo_track (at note_off) if preferred. Let's use tempo_at_note_on.
                            min_duration_sec = mido.tick2second(min_duration_ticks, ticks_per_beat_from_midi, tempo_at_note_on)
                            min_duration_sec = max(0.01, min_duration_sec) 

                            if duration_ticks < min_duration_ticks:
                                duration_ticks = min_duration_ticks
                                final_end_tick = actual_start_tick + duration_ticks
                            
                            if duration_sec < min_duration_sec:
                                duration_sec = min_duration_sec
                                final_end_sec = actual_start_sec + duration_sec
                            
                            # Ensure final_end_tick and final_end_sec are consistent if one was adjusted
                            if duration_ticks == min_duration_ticks and duration_sec != min_duration_sec: # ticks adjusted, secs not
                                final_end_sec = actual_start_sec + mido.tick2second(duration_ticks, ticks_per_beat_from_midi, tempo_at_note_on)
                                final_end_sec = max(final_end_sec, actual_start_sec + 0.01) # Ensure min sec again
                                duration_sec = final_end_sec - actual_start_sec

                            elif duration_sec == min_duration_sec and duration_ticks != min_duration_ticks: # secs adjusted, ticks not
                                # This case is less likely if min_duration_sec is derived from min_duration_ticks
                                pass


                            all_processed_notes.append({
                                'pitch': note_on_data['pitch'],
                                'velocity': note_on_data['velocity'],
                                'start_time_sec': actual_start_sec, # This is absolute for the track
                                'duration_sec': duration_sec,
                                'start_time_tick': actual_start_tick, # This is absolute for the track
                                'duration_ticks': duration_ticks,
                            })
                            
                            last_note_actual_end_map[(track_idx, pitch)] = {
                                'tick': final_end_tick,
                                'sec': final_end_sec
                            }
            
            # Sort all notes globally by their start time in seconds
            # Note: start_time_sec here is relative to each track's start.
            # This approach implicitly handles tracks starting at different real times if MIDI implies that
            # (e.g. via a delay on track 0 before other tracks start).
            # For a truly global timeline, Mido's merge_tracks and then processing is an alternative.
            # However, the current per-track `abs_sec_track` is often what's intended for `time_sec`.
            # If all tracks are meant to start at "time 0" of the piece, then `abs_sec_track` is fine.
            all_processed_notes.sort(key=lambda x: x['start_time_sec'])

            detailed_notes_for_roll = []
            raw_note_events_for_mapping_with_full_info = []

            max_overall_end_time_sec = 0.0
            for note in all_processed_notes:
                detailed_notes_for_roll.append({
                    'pitch': note['pitch'],
                    'start_time_sec': note['start_time_sec'],
                    'duration_sec': note['duration_sec'],
                    'velocity': note['velocity']
                })
                raw_note_events_for_mapping_with_full_info.append({
                    'time_sec': note['start_time_sec'],
                    'duration_sec': note['duration_sec'],
                    'duration_ticks': note['duration_ticks'],
                    'pitch': note['pitch'],
                    'velocity': note['velocity']
                })
                max_overall_end_time_sec = max(max_overall_end_time_sec, note['start_time_sec'] + note['duration_sec'])

            total_duration_sec_for_video = max_overall_end_time_sec
            if not all_processed_notes: # Fallback if no notes were processed
                 total_duration_sec_for_video = mid.length # Use original length


        except Exception as e:
            error_msg += f"MIDI解析エラー: {e}\n{traceback.format_exc()}"
        
        if not raw_note_events_for_mapping_with_full_info and not error_msg:
            error_msg = "MIDIノートイベント(マッピング用)処理失敗。"
        if not detailed_notes_for_roll and not error_msg and not raw_note_events_for_mapping_with_full_info: # Avoid double error if mapping also failed
             error_msg += "ピアノロール表示用MIDIノート処理失敗。"
        
        self.finished.emit(detailed_notes_for_roll, total_duration_sec_for_video, error_msg, raw_note_events_for_mapping_with_full_info, ticks_per_beat_from_midi)
# --- END OF MODIFIED SECTION IN MidiLoadThread.run ---


class PianoRollScene(QGraphicsScene):
    lyrics_display_items: List[QGraphicsTextItem]; lyric_segment_lines: List[QGraphicsLineItem]
    def __init__(self, parent=None):
        super().__init__(parent)
        self.note_items = []; self.lyric_note_map = {}; self.highlighted_items = []
        self.lyrics_display_items = []; self.lyric_segment_lines = []
        self.total_duration_sec = 10.0; self.min_pitch = 21; self.max_pitch = 108
        self.pixels_per_second = 50; self.pixels_per_pitch = 10
        self.lyrics_display_y_offset = -30; self.lyrics_font = QFont("Yu Mincho", 5, QFont.Bold)
        self.note_color = QColor(100,150,255); self.highlight_color = QColor(255,100,100,200)
        self.grid_pen = QPen(QColor(50,50,50)); self.text_color = QColor(200,200,200)
        self.lyric_segment_line_pen = QPen(QColor(200,200,0,150), 1, Qt.DashLine)
        self.setBackgroundBrush(QColor(30,30,30))
        self.initial_background_color = QColor(30,30,30) # Store initial color

    def time_to_x(self, time_sec: float) -> float: return time_sec * self.pixels_per_second
    def pitch_to_y(self, pitch: int) -> float: return (self.max_pitch - pitch) * self.pixels_per_pitch
    def draw_grid(self):
        scene_h = (self.max_pitch-self.min_pitch+1)*self.pixels_per_pitch; scene_w = self.total_duration_sec*self.pixels_per_second
        items_to_remove = [item for item in self.items() if (isinstance(item, QGraphicsLineItem) and item.pen().color() == self.grid_pen.color()) or \
                           (isinstance(item, QGraphicsTextItem) and item.defaultTextColor() == self.text_color and ("s" in item.toPlainText() or item.toPlainText().startswith("C")))]
        for item in items_to_remove:
            if item.scene() == self: self.removeItem(item)
        for p in range(self.min_pitch, self.max_pitch+1):
            y = self.pitch_to_y(p); line = self.addLine(0,y,scene_w,y,self.grid_pen)
            if p%12==0: line.setPen(QPen(QColor(80,80,80),1.5)); txt=QGraphicsTextItem(f"C{p//12-1}");txt.setDefaultTextColor(self.text_color);txt.setPos(-40,y-self.pixels_per_pitch/2);self.addItem(txt)
        for t_s in range(int(self.total_duration_sec)+2):
            x = self.time_to_x(float(t_s)); line = self.addLine(x,self.lyrics_display_y_offset -10 ,x,scene_h,self.grid_pen) 
            if t_s%5==0: line.setPen(QPen(QColor(80,80,80),1.5))
            txt=QGraphicsTextItem(f"{t_s}s");txt.setDefaultTextColor(self.text_color);txt.setPos(x-10,scene_h+5);self.addItem(txt)
        self.setSceneRect(-50, self.lyrics_display_y_offset-20, scene_w+70, scene_h+50-(self.lyrics_display_y_offset-20) )
    def load_midi_notes(self, notes: List[Dict[str, Any]], total_duration_sec: float):
        self.clear_scene_notes_and_highlights(); self.total_duration_sec = max(1.0, total_duration_sec)
        if notes: 
            pitches = [n['pitch'] for n in notes if 'pitch' in n]
            self.min_pitch=max(0,min(pitches)-5) if pitches else 21; self.max_pitch=min(127,max(pitches)+5) if pitches else 108
        else: self.min_pitch=21; self.max_pitch=108
        self.draw_grid(); new_note_items = []
        for note_info in notes:
            if not all(k in note_info for k in ['start_time_sec', 'pitch', 'duration_sec', 'velocity']): continue 
            x=self.time_to_x(note_info['start_time_sec']); y=self.pitch_to_y(note_info['pitch'])
            w=self.time_to_x(note_info['duration_sec']); h=self.pixels_per_pitch
            rect_item = QGraphicsRectItem(x,y,w,h); vel_factor = note_info['velocity'] / 127.0
            current_note_color = QColor(self.note_color); current_note_color.setHsv(self.note_color.hue(), int(self.note_color.saturationF()*255*(0.7+0.3*vel_factor)), int(self.note_color.valueF()*255*(0.7+0.3*vel_factor)))
            current_note_color.setAlpha(int(150 + 105 * vel_factor)); rect_item.setBrush(QBrush(current_note_color)); rect_item.setPen(QPen(Qt.black,0.5))
            self.addItem(rect_item); new_note_items.append(rect_item)
        self.note_items = new_note_items
    def display_lyrics_on_roll(self, current_line_segments: List[str], segment_start_times: List[float]):
        for item_list in [self.lyrics_display_items, self.lyric_segment_lines]:
            for item in item_list: 
                if item.scene() == self: self.removeItem(item)
            item_list.clear()
        if not current_line_segments or not segment_start_times: return
        scene_height = (self.max_pitch - self.min_pitch + 1) * self.pixels_per_pitch; new_lyrics_items = []; new_line_items = []
        for i, seg_text in enumerate(current_line_segments):
            # Do not display if seg_text is empty, even if time exists, to avoid clutter for "---" type segments on roll.
            if not seg_text.strip() or i >= len(segment_start_times): continue 
            start_x = self.time_to_x(segment_start_times[i])
            text_item = QGraphicsTextItem(seg_text); text_item.setFont(self.lyrics_font); text_item.setDefaultTextColor(self.text_color); text_item.setPos(start_x, self.lyrics_display_y_offset)
            self.addItem(text_item); new_lyrics_items.append(text_item)
            line = QGraphicsLineItem(start_x, self.lyrics_display_y_offset, start_x, scene_height); line.setPen(self.lyric_segment_line_pen)
            self.addItem(line); new_line_items.append(line)
        self.lyrics_display_items = new_lyrics_items; self.lyric_segment_lines = new_line_items
    def map_lyrics_to_notes(self, final_events_for_mapping: List[Dict[str, Any]], detailed_midi_notes_for_roll: List[Dict[str, Any]]):
        self.lyric_note_map.clear() 
        if not final_events_for_mapping or not detailed_midi_notes_for_roll or not self.note_items: return
        new_lyric_note_map = {}; time_tolerance = 0.05 
        for event in final_events_for_mapping: 
            if event['type'] == 'char':
                data = event['data']; event_time_sec = event['time']; event_pitch = data.get('pitch')
                if event_pitch is None: continue
                matched_rects = []
                for i, rect_item in enumerate(self.note_items):
                    if i < len(detailed_midi_notes_for_roll): 
                        note = detailed_midi_notes_for_roll[i]
                        # Match based on the note_start_time_sec from the event data for precision
                        if note['pitch'] == event_pitch and abs(note['start_time_sec'] - data.get('note_start_time_sec',event_time_sec)) < time_tolerance:
                            matched_rects.append(rect_item)
                if matched_rects:
                    key = (data['line_idx'], data['segment_idx_in_line'])
                    if key not in new_lyric_note_map: new_lyric_note_map[key] = []
                    if matched_rects: new_lyric_note_map[key].append(matched_rects[0]) # Typically one note per lyric segment
        self.lyric_note_map = new_lyric_note_map
    def highlight_lyric_segment(self, line_idx: int, segment_idx_in_line: int):
        items_to_reset = list(self.highlighted_items); self.highlighted_items.clear()
        for item in items_to_reset: 
            if item.scene() == self: 
                original_brush = item.data(Qt.UserRole + 1) 
                item.setBrush(original_brush if original_brush else QBrush(self.note_color)); item.setZValue(0)
        key_to_highlight = (line_idx, segment_idx_in_line)
        if key_to_highlight in self.lyric_note_map:
            for item_to_highlight in self.lyric_note_map[key_to_highlight]: # Iterate if multiple notes mapped (though usually one)
                if item_to_highlight.scene() == self: 
                    item_to_highlight.setData(Qt.UserRole + 1, item_to_highlight.brush()) 
                    item_to_highlight.setBrush(QBrush(self.highlight_color)); item_to_highlight.setZValue(1); self.highlighted_items.append(item_to_highlight)
        self.update()
    def clear_scene_notes_and_highlights(self):
        for item_list in [self.note_items, self.highlighted_items]: # Clear both lists
            for item in item_list: 
                if item.scene() == self: 
                    if item in self.highlighted_items: # If it was highlighted, restore brush
                        original_brush = item.data(Qt.UserRole + 1)
                        item.setBrush(original_brush if original_brush else QBrush(self.note_color)); item.setZValue(0)
                    if item in self.note_items: self.removeItem(item) # Only remove if it's a note item
            item_list.clear() # Clear the Python list
        self.lyric_note_map.clear()
    def clear_all_custom_items(self): # Clears notes, highlights, lyrics, AND redraws grid
        self.clear_scene_notes_and_highlights() 
        for item_list in [self.lyrics_display_items, self.lyric_segment_lines]:
            for item in item_list: 
                if item.scene() == self: self.removeItem(item)
            item_list.clear()
        self.draw_grid() 
    
    def clear_completely(self): # Clears EVERYTHING, no grid, back to initial state
        items_to_remove = list(self.items())
        for item in items_to_remove:
            if item.scene() == self:
                self.removeItem(item)
        
        self.note_items.clear()
        self.highlighted_items.clear()
        self.lyrics_display_items.clear()
        self.lyric_segment_lines.clear()
        self.lyric_note_map.clear()
        
        self.total_duration_sec = 10.0 
        self.min_pitch = 21
        self.max_pitch = 108
        self.setSceneRect(QRectF()) # Reset to default empty rect
        self.setBackgroundBrush(self.initial_background_color) 


class PianoRollView(QGraphicsView):
    def __init__(self, scene: PianoRollScene, parent=None):
        super().__init__(scene, parent); self.setRenderHint(QPainter.Antialiasing); self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse); self.setResizeAnchor(QGraphicsView.AnchorViewCenter); self.scale_factor=1.15
    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            if event.angleDelta().y() > 0: self.scale(self.scale_factor, self.scale_factor)
            else: self.scale(1.0 / self.scale_factor, 1.0 / self.scale_factor)
        else: super().wheelEvent(event)

class VideoGenThread(QThread):
    log_message = Signal(str); progress_update = Signal(int, int); generation_finished = Signal(bool, str)
    def __init__(self, params): super().__init__(); self.params = params
    class GuiLoggerAdapter(ILogger):
        def __init__(self, log_signal: Signal): self.log_signal = log_signal
        def info(self, msg): self.log_signal.emit(f"{msg}")
        def warning(self, msg): self.log_signal.emit(f"警告: {msg}")
        def error(self, msg): self.log_signal.emit(f"エラー: {msg}")
    def run(self):
        try:
            logger = self.GuiLoggerAdapter(self.log_message)
            # Use **self.params to unpack dictionary into keyword arguments
            generate_lyric_video_v2(logger=logger, progress_callback=self.progress_update.emit, **self.params)
            self.generation_finished.emit(True, self.params['output_video_path'])
        except Exception as e: self.log_message(f"ビデオ生成スレッド致命的エラー: {e}\n{traceback.format_exc()}"); self.generation_finished.emit(False, str(e))

SETTINGS_ORGANIZATION_NAME = "MySoft"; SETTINGS_APPLICATION_NAME = "MidT2M" # Changed AppName
SETTINGS_LAST_MIDI_DIR = "lastMidiDir"; SETTINGS_LAST_LYRICS_DIR = "lastLyricsDir"
SETTINGS_LAST_OUTPUT_DIR = "lastOutputDir"; SETTINGS_LAST_PROJECT_DIR = "lastProjectDir" # Added
PROJECT_FILE_EXTENSION = "mt2m"; PROJECT_FILE_FILTER = f"MidT2M Project (*.{PROJECT_FILE_EXTENSION})" # Changed extension

class MainWindow(QMainWindow):
    # (Constructor and UI setup methods remain the same as previous correct version)
    def __init__(self):
        super().__init__(); self.setWindowTitle("MidT2M"); self.setGeometry(50,50,1600,950) 
        self.midi_path: Optional[str] = None; self.output_video_path: Optional[str] = None
        self.detailed_midi_notes_for_roll: List[Dict[str, Any]] = [] 
        self.raw_note_events_for_mapping_with_duration: List[Dict[str,Any]] = [] 
        self.midi_total_duration_sec: float = 0.0; self.parsed_lyrics_structure: List[List[str]] = []
        self.final_events_for_mapping: List[Dict[str, Any]] = []
        self.midi_ticks_per_beat: int = 480 
        self.lyrics_edit_debouncer = QTimer(self); self.lyrics_edit_debouncer.setSingleShot(True); self.lyrics_edit_debouncer.setInterval(300); self.lyrics_edit_debouncer.timeout.connect(self._on_lyrics_debounced_change)
        self.midi_load_thread: Optional[MidiLoadThread] = None; self.video_gen_thread: Optional[VideoGenThread] = None
        self.temp_lyrics_file_path: Optional[str] = None; self.available_fonts: Dict[str, str] = {} 
        self.current_lyrics_text_for_roll: List[str] = []; self.current_segment_times_for_roll: List[float] = []
        self.current_highlight_key: Tuple[int, int] = (-1, -1)
        self.settings = QSettings(SETTINGS_ORGANIZATION_NAME, SETTINGS_APPLICATION_NAME)
        self.current_project_path: Optional[str] = None; self.project_modified = False; self.loading_project_or_midi = False
        self._create_menu_bar()
        main_widget = QWidget(); self.setCentralWidget(main_widget); self.root_layout = QHBoxLayout(main_widget)
        self.splitter = QSplitter(Qt.Horizontal); self.root_layout.addWidget(self.splitter)
        self.left_pane_widget = QWidget(); left_pane_layout = QVBoxLayout(self.left_pane_widget)
        self.editor_roll_splitter = QSplitter(Qt.Vertical)
        self.lyrics_editor_group = DropTargetGroupBox("歌詞エディター (.txt をドラッグ＆ドロップ)", [".txt"]); lyrics_editor_layout = QVBoxLayout()
        self.lyrics_edit = LyricsTextEdit(); self.lyrics_edit.setPlaceholderText("改行ごとに文字列表示は区切られます。 文/字/列 のように文字列を/で区切ることでmidiノートごとの対応関係を作ります。\n\n以下の記法が使えます\n\n文字列//文字列　　　　　空文字とすることでノートを飛ばせます。\n---文字列　　　　　　　　文字列を時間的に均等に分割して順次表示します。\nも|じ|れ|つ||文字列　　　　　|ごとのそれぞれの文字をノード内で順番に表示します。\n```/////```　　　　　　　　```で囲うことで記号を文字列として扱えます\n\n/---かえる|蛙/ のように混合して使えます。")
        self.lyrics_edit.textChanged.connect(self.on_lyrics_text_changed_schedule_debounce); self.lyrics_edit.cursorPositionChanged.connect(self.on_cursor_position_changed_debounced)
        self.lyrics_edit.file_dropped.connect(self._handle_lyrics_file_drop); self.lyrics_edit.request_focus.connect(lambda: self.lyrics_edit.setFocus())
        lyrics_editor_layout.addWidget(self.lyrics_edit)
        self.lyrics_load_button = QPushButton("歌詞をファイルからロード"); self.lyrics_load_button.clicked.connect(self._browse_lyrics_file_action); lyrics_editor_layout.addWidget(self.lyrics_load_button)
        self.lyrics_editor_group.setLayout(lyrics_editor_layout); self.lyrics_editor_group.file_dropped.connect(self._handle_lyrics_file_drop) 
        self.editor_roll_splitter.addWidget(self.lyrics_editor_group)
        self.piano_roll_group = DropTargetGroupBox("ピアノロール (.mid/.midi をドラッグ＆ドロップ)", [".mid", ".midi"]); piano_roll_layout = QVBoxLayout()
        self.piano_scene = PianoRollScene(); self.piano_view = PianoRollView(self.piano_scene); self.piano_view.setAcceptDrops(False) 
        piano_roll_layout.addWidget(self.piano_view); self.piano_roll_group.setLayout(piano_roll_layout); self.piano_roll_group.file_dropped.connect(self._handle_midi_file_drop)
        self.editor_roll_splitter.addWidget(self.piano_roll_group); self.editor_roll_splitter.setSizes([300, 500]); left_pane_layout.addWidget(self.editor_roll_splitter)
        self.splitter.addWidget(self.left_pane_widget)
        self.right_pane_widget = QScrollArea(); self.right_pane_widget.setWidgetResizable(True)
        self.right_pane_content_widget = QWidget(); self.right_pane_layout = QVBoxLayout(self.right_pane_content_widget); self.right_pane_widget.setWidget(self.right_pane_content_widget)
        self._create_parameter_widgets(); self._connect_param_widgets_to_modified_signal()
        self.right_pane_layout.addStretch(1); self.splitter.addWidget(self.right_pane_widget); self.splitter.setSizes([750, 550]) 
        self._load_system_fonts_to_combo(); self._update_ui_states()
        self.cursor_change_debouncer = QTimer(self); self.cursor_change_debouncer.setSingleShot(True); self.cursor_change_debouncer.setInterval(150); self.cursor_change_debouncer.timeout.connect(self._process_cursor_position_changed)
        self._update_window_title()
    def _create_menu_bar(self):
        menu_bar = self.menuBar(); file_menu = menu_bar.addMenu("ファイル")
        actions = [
            ("新規プロジェクト", QKeySequence(Qt.CTRL | Qt.Key_N), self._new_project_action, 'new_project_action'),
            ("プロジェクトをロード...", QKeySequence.Open, self._load_project_action, 'load_project_action'),
            ("プロジェクトを保存", QKeySequence.Save, self._save_project_action, 'save_project_action'),
            ("名前を付けてプロジェクトを保存...", QKeySequence.SaveAs, self._save_project_as_action, 'save_project_as_action')
        ]
        for text, shortcut, slot, attr_name in actions:
            action = QAction(text, self); action.setShortcut(shortcut); action.triggered.connect(slot); file_menu.addAction(action); setattr(self, attr_name, action)
        file_menu.addSeparator(); exit_action = QAction("終了", self); exit_action.setShortcut(QKeySequence.Quit); exit_action.triggered.connect(self.close); file_menu.addAction(exit_action)
    
    def _connect_param_widgets_to_modified_signal(self):
        widgets_to_connect = [self.midi_path_edit, self.output_video_path_edit, self.width_spin, self.height_spin, self.fps_spin,
                              self.min_char_render_size_spin, self.max_char_render_size_spin, self.font_size_base_spin, self.char_spacing_spin,
                              self.line_anchor_x_spin, self.line_anchor_y_spin, self.reference_pitch_spin, self.reference_velocity_spin,
                              self.pitch_offset_scale_spin, self.pitch_size_scale_spin, self.velocity_size_scale_spin,
                              self.duration_padding_threshold_ticks_spin, self.duration_padding_scale_per_tick_spin,
                              self.font_combo, self.line_placement_mode_combo, self.line_h_align_combo, self.text_v_align_combo]
        for widget in widgets_to_connect:
            if isinstance(widget, (QLineEdit, QTextEdit)): widget.textChanged.connect(self._mark_project_as_modified)
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)): widget.valueChanged.connect(self._mark_project_as_modified)
            elif isinstance(widget, QComboBox): widget.currentIndexChanged.connect(self._mark_project_as_modified)
    def _mark_project_as_modified(self, *args):
        if self.loading_project_or_midi: return
        if not self.project_modified: self.project_modified = True; self.setWindowModified(True); self._update_window_title()
    def _set_project_modified_status(self, modified: bool): self.project_modified = modified; self.setWindowModified(modified); self._update_window_title()
    def _update_window_title(self):
        title = "MidT2M" + (f" - {os.path.basename(self.current_project_path)}" if self.current_project_path else " - 新規プロジェクト") + ("[*]" if self.project_modified else "")
        self.setWindowTitle(title)

    def _get_last_dir(self, key: str) -> str: return self.settings.value(key, QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation))
    def _set_last_dir(self, key: str, directory: str): self.settings.setValue(key, directory)
    def _create_parameter_widgets(self): self._create_file_param_widgets(); self._create_video_param_widgets(); self._create_text_style_param_widgets(); self._create_dynamic_effects_param_widgets(); self._create_color_param_widgets(); self._create_action_widgets(); self._create_log_widgets()
    def _create_parameter_groupbox(self, title: str) -> Tuple[QGroupBox, QFormLayout]: group_box = QGroupBox(title); layout = QFormLayout(group_box); self.right_pane_layout.addWidget(group_box); return group_box, layout
    def _add_file_picker(self, layout: QFormLayout, lbl_txt: str, le: QLineEdit, cb, flt: str, settings_key: str):
        btn = QPushButton("参照..."); btn.clicked.connect(lambda: cb(le, flt, settings_key))
        if lbl_txt == "MIDI:": self.midi_browse_button = btn 
        cH = QHBoxLayout(); cH.addWidget(le); cH.addWidget(btn); layout.addRow(lbl_txt, cH)
    def _create_file_param_widgets(self):
        _, layout = self._create_parameter_groupbox("ファイル設定")
        self.midi_path_edit = QLineEdit(); self._add_file_picker(layout, "MIDI:", self.midi_path_edit, self._browse_midi_file_action, "MIDI (*.mid *.midi)", SETTINGS_LAST_MIDI_DIR)
        self.loaded_lyrics_path_display = QLineEdit(); self.loaded_lyrics_path_display.setReadOnly(True); layout.addRow("ロードされた歌詞ファイル:", self.loaded_lyrics_path_display)
        self.output_video_path_edit = QLineEdit(); self._add_file_picker(layout, "ビデオ出力パス:", self.output_video_path_edit, self._browse_output_video_file_action, "Video (*.mp4)", SETTINGS_LAST_OUTPUT_DIR)
    def _create_video_param_widgets(self):
        _, layout = self._create_parameter_groupbox("ビデオ設定")
        self.width_spin = QSpinBox(); self.width_spin.setRange(100,7680); self.width_spin.setValue(1920); layout.addRow("幅(px):",self.width_spin)
        self.height_spin = QSpinBox(); self.height_spin.setRange(100,4320); self.height_spin.setValue(1080); layout.addRow("高さ(px):",self.height_spin)
        self.fps_spin = QSpinBox(); self.fps_spin.setRange(1,120); self.fps_spin.setValue(30); layout.addRow("FPS:",self.fps_spin)
        self.min_char_render_size_spin = QSpinBox(); self.min_char_render_size_spin.setRange(1,100); self.min_char_render_size_spin.setValue(8); layout.addRow("最小文字サイズ(px):",self.min_char_render_size_spin)
        self.max_char_render_size_spin = QSpinBox(); self.max_char_render_size_spin.setRange(50,1000); self.max_char_render_size_spin.setValue(300); layout.addRow("最大文字サイズ(px):",self.max_char_render_size_spin)
    def _create_text_style_param_widgets(self): 
        _, layout = self._create_parameter_groupbox("テキストスタイル設定")
        self.font_combo = QComboBox(); layout.addRow("フォント:", self.font_combo); self.font_combo.setEnabled(True) 
        self.font_size_base_spin=QSpinBox();self.font_size_base_spin.setRange(10,500);self.font_size_base_spin.setValue(30);layout.addRow("基本フォントサイズ:",self.font_size_base_spin) 
        self.char_spacing_spin=QSpinBox();self.char_spacing_spin.setRange(-50,100);self.char_spacing_spin.setValue(10);layout.addRow("文字セグメント間隔(px):",self.char_spacing_spin)
        self.line_placement_mode_combo = QComboBox(); self.line_placement_mode_combo.addItems(["動的配置", "固定配置"]); self.line_placement_mode_combo.setCurrentText("動的配置"); layout.addRow("文字配置モード:", self.line_placement_mode_combo)
        self.line_h_align_combo = QComboBox(); self.line_h_align_combo.addItems(["左揃え", "中央揃え", "右揃え"]); self.line_h_align_combo.setCurrentText("中央揃え"); layout.addRow("行の水平揃え:", self.line_h_align_combo)
        self.line_anchor_x_spin = QSpinBox(); self.line_anchor_x_spin.setRange(-7680, 15360); self.line_anchor_x_spin.setValue(self.width_spin.value() // 2); layout.addRow("行アンカーX:", self.line_anchor_x_spin)
        # Use a more robust way to store previous value for anchor auto-update
        self.width_spin.valueChanged.connect(lambda val: self.line_anchor_x_spin.setValue(val // 2) if self.line_anchor_x_spin.value() == (getattr(self.width_spin, "_previousValueForAnchor", val)//2) else None)
        self.width_spin.editingFinished.connect(lambda: setattr(self.width_spin, "_previousValueForAnchor", self.width_spin.value()))
        setattr(self.width_spin, "_previousValueForAnchor", self.width_spin.value()) # Initialize
        self.line_anchor_y_spin = QSpinBox(); self.line_anchor_y_spin.setRange(-4320, 8640); self.line_anchor_y_spin.setValue(self.height_spin.value() // 2); layout.addRow("行アンカーY:", self.line_anchor_y_spin)
        self.height_spin.valueChanged.connect(lambda val: self.line_anchor_y_spin.setValue(val // 2) if self.line_anchor_y_spin.value() == (getattr(self.height_spin, "_previousValueForAnchor", val)//2) else None)
        self.height_spin.editingFinished.connect(lambda: setattr(self.height_spin, "_previousValueForAnchor", self.height_spin.value()))
        setattr(self.height_spin, "_previousValueForAnchor", self.height_spin.value()) # Initialize
        self.text_v_align_combo=QComboBox();self.text_v_align_combo.addItems(["中央揃え","上揃え","下揃え", "ベースライン"]); self.text_v_align_combo.setCurrentText("中央揃え");layout.addRow("行内垂直揃え:",self.text_v_align_combo)
    def _add_slider_for_spinbox(self, layout: QFormLayout, label_text: str, spinbox: QWidget, s_min: int, s_max: int, factor: float = 1.0, is_double: bool = False):
        slider = QSlider(Qt.Horizontal); slider.setRange(s_min, s_max); cH = QHBoxLayout(); cH.addWidget(spinbox, 1); cH.addWidget(slider, 3); layout.addRow(label_text, cH)
        update_slider = lambda v: slider.setValue(int(v * factor))
        update_spinbox = lambda v: spinbox.setValue(float(v)/factor if is_double else int(float(v)/factor)) # type: ignore
        spinbox.valueChanged.connect(update_slider) # type: ignore
        slider.valueChanged.connect(update_spinbox); update_slider(spinbox.value()); slider.valueChanged.connect(self._mark_project_as_modified) # type: ignore
    def _create_dynamic_effects_param_widgets(self):
        _, layout = self._create_parameter_groupbox("動的テキストエフェクト設定")
        self.reference_pitch_spin = QSpinBox(); self.reference_pitch_spin.setRange(0,127); self.reference_pitch_spin.setValue(60); self._add_slider_for_spinbox(layout, "基準ピッチ:", self.reference_pitch_spin, 0, 127)
        self.reference_velocity_spin = QSpinBox(); self.reference_velocity_spin.setRange(1,127); self.reference_velocity_spin.setValue(64); self._add_slider_for_spinbox(layout, "基準ベロシティ:", self.reference_velocity_spin, 1, 127)
        self.pitch_offset_scale_spin = QDoubleSpinBox(); self.pitch_offset_scale_spin.setRange(-20.0, 20.0); self.pitch_offset_scale_spin.setSingleStep(0.1); self.pitch_offset_scale_spin.setValue(0.0); self.pitch_offset_scale_spin.setDecimals(2); self._add_slider_for_spinbox(layout, "ピッチYオフセット強度:", self.pitch_offset_scale_spin, -2000, 2000, 100.0, True)
        self.pitch_size_scale_spin = QDoubleSpinBox(); self.pitch_size_scale_spin.setRange(-0.5, 2.0); self.pitch_size_scale_spin.setSingleStep(0.01); self.pitch_size_scale_spin.setValue(0.0); self.pitch_size_scale_spin.setDecimals(3); self._add_slider_for_spinbox(layout, "ピッチサイズ変化率:", self.pitch_size_scale_spin, -500, 2000, 1000.0, True)
        self.velocity_size_scale_spin = QDoubleSpinBox(); self.velocity_size_scale_spin.setRange(-0.5, 2.0); self.velocity_size_scale_spin.setSingleStep(0.01); self.velocity_size_scale_spin.setValue(0.0); self.velocity_size_scale_spin.setDecimals(3); self._add_slider_for_spinbox(layout, "ベロシティサイズ変化率:", self.velocity_size_scale_spin, -500, 2000, 1000.0, True)
        self.duration_padding_threshold_ticks_spin = QSpinBox(); self.duration_padding_threshold_ticks_spin.setRange(0, 2000); self.duration_padding_threshold_ticks_spin.setSingleStep(10); self.duration_padding_threshold_ticks_spin.setValue(240); self._add_slider_for_spinbox(layout, "ノート長パディング閾値(tick):", self.duration_padding_threshold_ticks_spin, 0, 2000, 1)
        self.duration_padding_scale_per_tick_spin = QDoubleSpinBox(); self.duration_padding_scale_per_tick_spin.setRange(0.0, 5.0); self.duration_padding_scale_per_tick_spin.setSingleStep(0.01); self.duration_padding_scale_per_tick_spin.setValue(0.1); self.duration_padding_scale_per_tick_spin.setDecimals(3); self._add_slider_for_spinbox(layout, "ノート長パディング強度(px/tick):", self.duration_padding_scale_per_tick_spin, 0, 5000, 1000.0, True)
    def _load_system_fonts_to_combo(self):
        self.available_fonts.clear() 
        if IS_WINDOWS:
            try: self.available_fonts.update(get_system_fonts_windows()); self.available_fonts.update(get_user_fonts_windows())
            except Exception as e: self.log_message(f"システムフォント読込エラー: {e}", "error")
        font_path_placeholder = "カスタムフォントパス..."; self.available_fonts[font_path_placeholder] = "" 
        sorted_font_names = sorted(self.available_fonts.keys(), key=lambda x: x.lower())
        current_selection = self.font_combo.currentText(); self.font_combo.clear(); self.font_combo.addItems(sorted_font_names)
        if current_selection and current_selection in self.available_fonts: self.font_combo.setCurrentText(current_selection)
        else:
            default_font = next((f for f in ["Yu Mincho", "MS Mincho", "TakaoMincho", "Arial"] if f in self.available_fonts), None)
            if default_font: self.font_combo.setCurrentText(default_font)
            elif font_path_placeholder in self.available_fonts: self.font_combo.setCurrentText(font_path_placeholder)
            elif sorted_font_names: self.font_combo.setCurrentIndex(0)
        if not IS_WINDOWS and len(self.available_fonts) <= 1: self.log_message("このOSではシステムフォント自動検出非対応。カスタムパスを使用してください。", "info")
        try: self.font_combo.currentIndexChanged.disconnect(self._on_font_combo_changed)
        except RuntimeError: pass
        self.font_combo.currentIndexChanged.connect(self._on_font_combo_changed)
    def _on_font_combo_changed(self, index: int):
        selected = self.font_combo.itemText(index)
        if selected == "カスタムフォントパス...":
            font_path, _ = QFileDialog.getOpenFileName(self, "フォントファイルを選択", self._get_last_dir("lastFontDir"), "Font files (*.ttf *.otf)")
            if font_path:
                name = Path(font_path).stem
                if name not in self.available_fonts or self.available_fonts[name] != font_path: # New or different path for same stem
                    self.available_fonts[name] = font_path; self._load_system_fonts_to_combo(); self.font_combo.setCurrentText(name)
                else: self.font_combo.setCurrentText(name) # Just re-select if path is identical
                self._set_last_dir("lastFontDir", os.path.dirname(font_path))
            else: # User cancelled custom font selection
                prev_valid = next((self.font_combo.itemText(i) for i in range(self.font_combo.count()) if self.font_combo.itemText(i) != "カスタムフォントパス..." and self.available_fonts.get(self.font_combo.itemText(i))), None)
                if prev_valid: self.font_combo.setCurrentText(prev_valid)
                elif "カスタムフォントパス..." in self.available_fonts : self.font_combo.setCurrentText("カスタムフォントパス...")
                elif self.font_combo.count() > 0: self.font_combo.setCurrentIndex(0)
        self._mark_project_as_modified()
    def _create_color_button(self,initial_rgb:Tuple[int,int,int])->QPushButton: btn=QPushButton();btn.setFixedSize(QSize(100,25));self._update_color_button_style(btn,QColor.fromRgb(*initial_rgb));btn.clicked.connect(lambda:self._pick_color(btn));return btn
    def _update_color_button_style(self,btn:QPushButton,qc:QColor): btn.setText(qc.name());pal=btn.palette();pal.setColor(QPalette.Button,qc);txt_c=Qt.white if(qc.redF()*0.299+qc.greenF()*0.587+qc.blueF()*0.114)<0.5 else Qt.black;pal.setColor(QPalette.ButtonText,txt_c);btn.setPalette(pal);btn.setAutoFillBackground(True);btn.update()
    def _pick_color(self,btn_to_update:QPushButton): 
        initial_color = QColor(btn_to_update.text()) if QColor.isValidColor(btn_to_update.text()) else Qt.white
        color=QColorDialog.getColor(initial_color,self,"色を選択");
        if color.isValid(): self._update_color_button_style(btn_to_update,color); self._mark_project_as_modified() 
    def _get_color_from_button(self,btn:QPushButton)->Tuple[int,int,int]:return QColor(btn.text()).getRgb()[:3] 
    def _create_color_param_widgets(self): _,layout=self._create_parameter_groupbox("色設定");self.bg_color_button=self._create_color_button((0,0,0));layout.addRow("背景色:",self.bg_color_button);self.text_color_button=self._create_color_button((255,255,255));layout.addRow("文字色:",self.text_color_button)
    def _create_action_widgets(self): ag=QGroupBox("アクション");al=QVBoxLayout(ag);self.generate_button=QPushButton("ビデオを生成");self.generate_button.clicked.connect(self.start_video_generation);al.addWidget(self.generate_button);self.progress_bar=QProgressBar();self.progress_bar.setVisible(False);al.addWidget(self.progress_bar);self.right_pane_layout.addWidget(ag)
    def _create_log_widgets(self): lg=QGroupBox("ログ");ll=QVBoxLayout(lg);self.log_browser=QTextEdit();self.log_browser.setReadOnly(True);ll.addWidget(self.log_browser);self.right_pane_layout.addWidget(lg)
    def _browse_file(self, le: Optional[QLineEdit], cap: str, flt: str, key: str, save=False) -> Optional[str]:
        s_dir = self._get_last_dir(key); p_text = le.text() if le and le.text() else ""
        c_dir = None
        if p_text:
            if os.path.isdir(p_text): c_dir = p_text
            elif os.path.isfile(p_text): c_dir = os.path.dirname(p_text)
            elif os.path.dirname(p_text) and os.path.exists(os.path.dirname(p_text)): c_dir = os.path.dirname(p_text)
        s_dir_use = c_dir if c_dir and os.path.exists(c_dir) else s_dir
        path_fn = QFileDialog.getSaveFileName if save else QFileDialog.getOpenFileName
        path, _ = path_fn(self, cap, s_dir_use, flt)
        if path:
            if le: le.setText(path) # This will trigger _mark_project_as_modified if not loading_project_or_midi
            self._set_last_dir(key, os.path.dirname(path)); return path
        return None
    def _load_midi_file(self, midi_file_path: str):
        if not midi_file_path or not os.path.exists(midi_file_path): 
            self.log_message(f"MIDIパス無効または空: '{midi_file_path}'。ロードスキップ。", "warning")
            self.midi_path = None
            self.midi_path_edit.setText(midi_file_path or "") 
            self.detailed_midi_notes_for_roll.clear()
            self.raw_note_events_for_mapping_with_duration.clear()
            self.midi_total_duration_sec = 0.0
            self.midi_ticks_per_beat = 480
            self.piano_scene.clear_completely() # Use clear_completely here
            self._recalculate_final_events_and_update_mapping() 
            self._process_cursor_position_changed()
            # If this was part of _apply_project_data, we need to ensure loading_project_or_midi is reset correctly
            if self.loading_project_or_midi and (not self.midi_load_thread or not self.midi_load_thread.isRunning()):
                self.loading_project_or_midi = False
                self._update_ui_states()
            return

        if self.midi_path == midi_file_path and not self.loading_project_or_midi: self.log_message(f"MIDI '{os.path.basename(midi_file_path)}' ロード済。", "info"); return
        self.midi_path = midi_file_path; self.midi_path_edit.setText(midi_file_path); self.log_message(f"MIDI '{os.path.basename(midi_file_path)}' 選択。読込中...", "info")
        self._update_ui_states(is_loading_midi=True)
        if self.midi_load_thread and self.midi_load_thread.isRunning(): self.midi_load_thread.terminate(); self.midi_load_thread.wait()
        self.midi_load_thread = MidiLoadThread(midi_file_path); self.midi_load_thread.finished.connect(self._on_midi_load_finished); self.midi_load_thread.start()
    def _browse_midi_file_action(self, le_target: QLineEdit, filt: str, key: str): 
        path = self._browse_file(le_target, "MIDIファイルを選択", filt, key)
        if path: self._load_midi_file(path) # le_target.setText is handled by _browse_file
    @Slot(str)
    def _handle_midi_file_drop(self, filePath: str): self.log_message(f"MIDIドラッグ＆ドロップ: {filePath}", "info"); self._load_midi_file(filePath)
    def _load_lyrics_from_file(self, lyrics_file_path: str):
        if not lyrics_file_path or not os.path.exists(lyrics_file_path): self.log_message(f"歌詞ファイルパス無効: {lyrics_file_path}", "error"); return
        try:
            with open(lyrics_file_path, 'r', encoding='utf-8') as f: content = f.read()
            try: self.lyrics_edit.textChanged.disconnect(self.on_lyrics_text_changed_schedule_debounce)
            except RuntimeError: pass # Not connected is fine
            current_editor_text = self.lyrics_edit.toPlainText()
            if current_editor_text != content: 
                self.lyrics_edit.setText(content) # This will trigger _mark_project_as_modified (if not loading) via its own connected signal chain
                self.log_message(f"歌詞を '{os.path.basename(lyrics_file_path)}' からロード。", "info")
            else: self.log_message(f"歌詞ファイル '{os.path.basename(lyrics_file_path)}' 内容はエディタと同じ。", "info")
            self.lyrics_edit.textChanged.connect(self.on_lyrics_text_changed_schedule_debounce) # Reconnect
            self.loaded_lyrics_path_display.setText(lyrics_file_path)
            # If content was same or loading project, ensure parse happens
            if current_editor_text == content or self.loading_project_or_midi: 
                self._on_lyrics_debounced_change() # This will parse, recalc, update roll
        except Exception as e: self.log_message(f"歌詞ファイル '{lyrics_file_path}' 読込エラー: {e}", "error")
        finally: # Ensure reconnected if an error happened before explicit reconnect
            if not self.lyrics_edit.signalsBlocked():
                 try: self.lyrics_edit.textChanged.disconnect(self.on_lyrics_text_changed_schedule_debounce)
                 except RuntimeError: pass 
                 self.lyrics_edit.textChanged.connect(self.on_lyrics_text_changed_schedule_debounce)
    def _browse_lyrics_file_action(self): 
        path = self._browse_file(None, "歌詞ファイルをロード", "Text Files (*.txt)", SETTINGS_LAST_LYRICS_DIR)
        if path: self._load_lyrics_from_file(path); self._mark_project_as_modified() # Explicitly mark
    @Slot(str)
    def _handle_lyrics_file_drop(self, filePath: str): self.log_message(f"歌詞ドラッグ＆ドロップ: {filePath}", "info"); self._load_lyrics_from_file(filePath); self._mark_project_as_modified()
    def _browse_output_video_file_action(self, le_target: QLineEdit, filt: str, key: str): 
        path = self._browse_file(le_target, "出力ビデオファイル名を設定", filt, key, save=True)
        if path: self.output_video_path = path # le_target.setText handled by _browse_file
    @Slot(list, float, str, list, int) 
    def _on_midi_load_finished(self, detailed_notes, total_duration, error_msg, raw_mapping_notes, ticks_per_beat): 
        self._update_ui_states(is_loading_midi=False) # Important: update UI state *before* intensive calcs
        if error_msg:
            self.log_message(f"MIDIロードエラー: {error_msg}", "error"); 
            self.detailed_midi_notes_for_roll.clear(); 
            self.raw_note_events_for_mapping_with_duration.clear(); 
            self.midi_total_duration_sec = 0.0; 
            self.midi_ticks_per_beat = 480; 
            self.piano_scene.clear_completely() # Use clear_completely here
        else:
            self.detailed_midi_notes_for_roll = detailed_notes; self.raw_note_events_for_mapping_with_duration = raw_mapping_notes; self.midi_total_duration_sec = total_duration; self.midi_ticks_per_beat = ticks_per_beat
            self.piano_scene.load_midi_notes(self.detailed_midi_notes_for_roll, self.midi_total_duration_sec) # This will draw grid
            self.log_message(f"MIDI '{os.path.basename(self.midi_path)}' ロード完了。TPB:{self.midi_ticks_per_beat},ノート(表示):{len(detailed_notes)},ノート(Map):{len(raw_mapping_notes)},時間:{total_duration:.2f}s", "info")
        
        self._recalculate_final_events_and_update_mapping(); self._process_cursor_position_changed() # These update internal states and roll
        if not self.loading_project_or_midi: self._mark_project_as_modified() # Mark modified if not part of project loading
        
        # This is the final point for MIDI loading part of _apply_project_data
        if self.loading_project_or_midi: # If this was part of project load
            self.loading_project_or_midi = False # MIDI part of loading is done
            self._update_ui_states() # Re-enable UI

    @Slot()
    def on_lyrics_text_changed_schedule_debounce(self): self.lyrics_edit_debouncer.start();_ = self._mark_project_as_modified() if not self.loading_project_or_midi else None
    @Slot()
    def _on_lyrics_debounced_change(self): 
        new_parsed = []
        for line_text in self.lyrics_edit.toPlainText().splitlines():
            if line_text.strip():
                new_parsed.append(line_text.split('/'))
            else: 
                new_parsed.append([])

        if self.parsed_lyrics_structure != new_parsed: 
            self.parsed_lyrics_structure = new_parsed
        self._recalculate_final_events_and_update_mapping()
        self._process_cursor_position_changed() 


    def _calculate_final_events_for_mapping_optimized(self): 
        if not self.raw_note_events_for_mapping_with_duration or not self.parsed_lyrics_structure:
            self.final_events_for_mapping = []
            return

        note_evs = self.raw_note_events_for_mapping_with_duration
        p_lyrics = self.parsed_lyrics_structure
        final_evs = []
        note_ptr = 0

        for line_idx, line_raw_segments in enumerate(p_lyrics):
            seg_idx_in_line = 0
            for raw_seg_text in line_raw_segments: 
                if note_ptr >= len(note_evs):
                    break 
                
                current_note = note_evs[note_ptr]
                parsed_dyn = parse_dynamic_segment(raw_seg_text)
                
                event_data = {
                    **parsed_dyn, 
                    'velocity': current_note['velocity'], 
                    'pitch': current_note['pitch'], 
                    'duration_ticks': current_note['duration_ticks'], 
                    'note_start_time_sec': current_note['time_sec'], 
                    'note_duration_sec': current_note['duration_sec'],
                    'line_idx': line_idx, 
                    'segment_idx_in_line': seg_idx_in_line
                }
                final_evs.append({'time': current_note['time_sec'], 'type': 'char', 'data': event_data})
                seg_idx_in_line += 1
                note_ptr += 1 
            
            if note_ptr >= len(note_evs) and line_idx < len(p_lyrics) - 1:
                break 
        
        self.final_events_for_mapping = final_evs

    def _recalculate_final_events_and_update_mapping(self): 
        if not self.raw_note_events_for_mapping_with_duration:
            self.final_events_for_mapping = []
        else: 
            self._calculate_final_events_for_mapping_optimized()
        
        self.piano_scene.map_lyrics_to_notes(self.final_events_for_mapping, self.detailed_midi_notes_for_roll or [])
        self._process_cursor_position_changed()

    @Slot()
    def on_cursor_position_changed_debounced(self): self.cursor_change_debouncer.start() 
    @Slot()
    def _process_cursor_position_changed(self): 
        cursor = self.lyrics_edit.textCursor(); blk_num = cursor.blockNumber()
        new_roll_txt = []; new_roll_times = []
        doc_lines_text = self.lyrics_edit.toPlainText().splitlines()
        
        current_line_actual_segments_from_parsed_structure = []
        if 0 <= blk_num < len(self.parsed_lyrics_structure):
            current_line_actual_segments_from_parsed_structure = self.parsed_lyrics_structure[blk_num]
            new_roll_txt = list(current_line_actual_segments_from_parsed_structure)

            line_events_for_roll_times = sorted([
                ev['data'] for ev in self.final_events_for_mapping 
                if ev['type']=='char' and ev['data']['line_idx']==blk_num
            ], key=lambda x:x['segment_idx_in_line'])

            new_roll_times = [0.0]*len(new_roll_txt) 
            for i_ev_data, event_data in enumerate(line_events_for_roll_times):
                seg_idx_from_event = event_data['segment_idx_in_line']
                if 0 <= seg_idx_from_event < len(new_roll_times):
                    new_roll_times[seg_idx_from_event] = event_data['note_start_time_sec']
        
        if self.current_lyrics_text_for_roll != new_roll_txt or self.current_segment_times_for_roll != new_roll_times:
            self.current_lyrics_text_for_roll = new_roll_txt; self.current_segment_times_for_roll = new_roll_times
            self.piano_scene.display_lyrics_on_roll(self.current_lyrics_text_for_roll, self.current_segment_times_for_roll)
        
        new_hl_key = (-1,-1)
        if 0 <= blk_num < len(doc_lines_text):
            segments_for_cursor_logic = self.parsed_lyrics_structure[blk_num] if 0 <= blk_num < len(self.parsed_lyrics_structure) else doc_lines_text[blk_num].split('/')
            cur_pos_in_blk = cursor.positionInBlock(); char_cnt = 0; target_seg_idx_for_highlight = -1
            editor_line_text = doc_lines_text[blk_num]
            
            raw_segments_in_editor_line = editor_line_text.split('/')
            
            char_cnt_local = 0
            for editor_idx, editor_seg_txt in enumerate(raw_segments_in_editor_line):
                seg_len_in_editor_display = len(editor_seg_txt)
                
                if char_cnt_local <= cur_pos_in_blk <= char_cnt_local + seg_len_in_editor_display:
                    target_seg_idx_for_highlight = editor_idx
                    break
                if editor_idx < len(raw_segments_in_editor_line) -1 and cur_pos_in_blk == char_cnt_local + seg_len_in_editor_display + 1:
                    target_seg_idx_for_highlight = editor_idx + 1 
                    break
                
                char_cnt_local += seg_len_in_editor_display + 1 
            
            if target_seg_idx_for_highlight == -1 and cur_pos_in_blk >= char_cnt_local and raw_segments_in_editor_line: 
                 target_seg_idx_for_highlight = len(raw_segments_in_editor_line)-1


            if target_seg_idx_for_highlight != -1: 
                new_hl_key = (blk_num, target_seg_idx_for_highlight)
        
        if self.current_highlight_key != new_hl_key: 
            self.current_highlight_key = new_hl_key
            self.piano_scene.highlight_lyric_segment(*self.current_highlight_key)

    def _validate_inputs(self) -> bool: 
        self.midi_path = self.midi_path_edit.text() 
        if not self.midi_path or not os.path.exists(self.midi_path): self.log_message("MIDIファイルパス無効。","error"); return False
        lyrics_content = self.lyrics_edit.toPlainText()
        try:
            fd, self.temp_lyrics_file_path = tempfile.mkstemp(suffix=".txt", text=True)
            with os.fdopen(fd, 'w', encoding='utf-8') as tf: tf.write(lyrics_content) 
            if not lyrics_content.strip(): self.log_message("歌詞が実質的に空です。", "warning")
        except Exception as e: self.log_message(f"一時歌詞ファイル作成エラー: {e}","error"); return False
        font_name = self.font_combo.currentText(); font_path = self.available_fonts.get(font_name)
        if not font_name or not font_path or (font_name == "カスタムフォントパス..." and not (font_path and os.path.exists(font_path))): # Check if custom path actually chosen and valid
            self.log_message("フォント未選択または無効（カスタムパスの場合はファイル選択要）。","error"); self._cleanup_temp_lyrics(); return False
        if font_path and not os.path.exists(font_path): self.log_message(f"フォント '{font_name}' パス無効: '{font_path}'。","error"); self._cleanup_temp_lyrics(); return False
        try: _=ImageFont.truetype(font_path,10)
        except Exception as e: self.log_message(f"フォント '{font_name}' 読込不可: {e}","error"); self._cleanup_temp_lyrics(); return False
        self.output_video_path = self.output_video_path_edit.text()
        if not self.output_video_path: self.log_message("出力ビデオパス未指定。","error"); self._cleanup_temp_lyrics(); return False
        out_dir = os.path.dirname(self.output_video_path)
        if out_dir and not os.path.exists(out_dir):
            try: os.makedirs(out_dir); self.log_message(f"出力先ディレクトリ '{out_dir}' 作成。","info")
            except Exception as e: self.log_message(f"出力先ディレクトリ作成失敗 '{out_dir}': {e}","error"); self._cleanup_temp_lyrics(); return False
        return True
    def _cleanup_temp_lyrics(self): 
        if self.temp_lyrics_file_path and os.path.exists(self.temp_lyrics_file_path):
            try: os.remove(self.temp_lyrics_file_path); self.temp_lyrics_file_path=None
            except Exception as e: self.log_message(f"一時歌詞ファイル削除エラー: {e}","warning")
    def start_video_generation(self): 
        if self.video_gen_thread and self.video_gen_thread.isRunning(): self.log_message("ビデオ生成実行中。","warning"); return
        if not self._validate_inputs(): return
        font_path = self.available_fonts.get(self.font_combo.currentText())
        if not font_path: self.log_message(f"フォントパス取得エラー: {self.font_combo.currentText()}", "error"); return # Should be caught by validate_inputs
        map_v = {"中央揃え":"center", "上揃え":"top", "下揃え":"bottom", "ベースライン":"baseline"}
        map_p = {"動的配置":"dynamic", "固定配置":"fixed"}; map_h = {"左揃え":"left", "中央揃え":"center", "右揃え":"right"}
        # Simplified param collection
        params = {
            'midi_path':self.midi_path, 'lyrics_path':self.temp_lyrics_file_path, 
            'output_video_path':self.output_video_path, 'font_path':font_path, 
            'width':self.width_spin.value(),'height':self.height_spin.value(),'fps':self.fps_spin.value(),
            'font_size_base_param':self.font_size_base_spin.value(), 'char_spacing':self.char_spacing_spin.value(),
            'bg_color':self._get_color_from_button(self.bg_color_button),'text_color':self._get_color_from_button(self.text_color_button),
            'text_vertical_align': map_v.get(self.text_v_align_combo.currentText(), "center"),
            'line_placement_mode': map_p.get(self.line_placement_mode_combo.currentText(), "dynamic"),
            'line_h_align': map_h.get(self.line_h_align_combo.currentText(), "center"),
            'line_anchor_x': self.line_anchor_x_spin.value(),'line_anchor_y': self.line_anchor_y_spin.value(),
            'pitch_offset_scale': self.pitch_offset_scale_spin.value(),'pitch_size_scale': self.pitch_size_scale_spin.value(),
            'velocity_size_scale': self.velocity_size_scale_spin.value(),'reference_pitch': self.reference_pitch_spin.value(),
            'reference_velocity': self.reference_velocity_spin.value(),
            'duration_padding_threshold_ticks': self.duration_padding_threshold_ticks_spin.value(),
            'duration_padding_scale_per_tick': self.duration_padding_scale_per_tick_spin.value(),
            'min_char_render_size': self.min_char_render_size_spin.value(),'max_char_render_size': self.max_char_render_size_spin.value(),
        }
        self.video_gen_thread = VideoGenThread(params)
        self.video_gen_thread.log_message.connect(lambda msg: self.log_message(msg,"default"))
        self.video_gen_thread.progress_update.connect(self.update_progress); self.video_gen_thread.generation_finished.connect(self.on_generation_finished)
        self.video_gen_thread.start(); self._update_ui_states(is_generating_video=True); self.log_message("ビデオ生成開始...","info")
    @Slot(str,str)
    def log_message(self,message:str,level:str="info"): 
        prefix_map = {"info":"[INFO] ","warning":"警告: ","error":"エラー: "}
        prefix = prefix_map.get(level,"") if not any(message.lower().startswith(p) for p in ["[info]","ビデオ生成プロセス開始:","警告:","エラー:"]) else ""
        self.log_browser.append(f"{prefix}{message}")
    @Slot(int,int)
    def update_progress(self,current:int,total:int):self.progress_bar.setRange(0,total);self.progress_bar.setValue(current) 
    @Slot(bool,str)
    def on_generation_finished(self,success:bool,message:str): 
        self.log_message(f"ビデオ生成完了: {message}" if success else f"ビデオ生成失敗: {message}", "info" if success else "error")
        if not success: QMessageBox.critical(self,"失敗",f"ビデオ生成失敗:\n{message}")
        self._update_ui_states(is_generating_video=False);self._cleanup_temp_lyrics()
    def _update_ui_states(self, is_generating_video=False, is_loading_midi=False): 
        self.loading_project_or_midi = is_generating_video or is_loading_midi 
        busy = self.loading_project_or_midi; self.generate_button.setEnabled(not busy); self.progress_bar.setVisible(is_generating_video)
        
        all_param_controls = []
        for group_name in ["ファイル設定", "ビデオ設定", "テキストスタイル設定", "動的テキストエフェクト設定", "色設定"]:
            group_box = next((gb for gb in self.right_pane_content_widget.findChildren(QGroupBox) if gb.title() == group_name), None)
            if group_box:
                all_param_controls.extend(group_box.findChildren(QWidget)) 

        all_param_controls.append(self.lyrics_load_button) 
        if hasattr(self, 'midi_browse_button'): all_param_controls.append(self.midi_browse_button)
        
        for ctrl in all_param_controls:
            if ctrl != self.generate_button and hasattr(ctrl, 'setEnabled'): 
                ctrl.setEnabled(not busy)
        
        self.lyrics_edit.setReadOnly(busy)
        for action_attr in ['new_project_action', 'load_project_action', 'save_project_action', 'save_project_as_action']:
            if hasattr(self, action_attr): getattr(self, action_attr).setEnabled(not busy)
        
    def _confirm_unsaved_changes(self) -> bool:
        if not self.project_modified: return True 
        reply = QMessageBox.question(self, "未保存の変更", "現在のプロジェクトには未保存の変更があります。保存しますか？", QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel, QMessageBox.Save)
        if reply == QMessageBox.Save: return self._save_project_action() 
        return reply == QMessageBox.Discard
    
    def _get_default_project_data(self) -> Dict[str, Any]:
        default_params = {
            'width': 1920, 'height': 1080, 'fps': 30,
            'min_char_render_size': 8, 'max_char_render_size': 300,
            'font_size_base': 30, 'char_spacing': 10,
            'line_anchor_x': 1920 // 2, 'line_anchor_y': 1080 // 2,
            'reference_pitch': 60, 'reference_velocity': 64,
            'pitch_offset_scale': 0.0, 'pitch_size_scale': 0.0, 'velocity_size_scale': 0.0,
            'duration_padding_threshold_ticks': 240, 'duration_padding_scale_per_tick': 0.1,
            'font_name': "", 'font_path_if_custom': "",
            'line_placement_mode': "動的配置", 'line_h_align': "中央揃え", 'text_v_align': "中央揃え",
            'bg_color_rgb': [0,0,0], 'text_color_rgb': [255,255,255],
            'loaded_lyrics_display_path': ""
        }

        font_path_placeholder = "カスタムフォントパス..."
        default_font_candidates = ["Yu Mincho", "MS Mincho", "TakaoMincho", "Arial"]
        chosen_font_name = ""
        
        for font_cand in default_font_candidates:
            if font_cand in self.available_fonts and self.available_fonts.get(font_cand):
                chosen_font_name = font_cand; break
        if not chosen_font_name:
            for name, path_ in self.available_fonts.items():
                if path_ and name != font_path_placeholder:
                    chosen_font_name = name; break
        if not chosen_font_name and font_path_placeholder in self.available_fonts:
            chosen_font_name = font_path_placeholder
        if not chosen_font_name and self.font_combo.count() > 0:
             chosen_font_name = self.font_combo.itemText(0)

        default_params['font_name'] = chosen_font_name
        default_params['font_path_if_custom'] = self.available_fonts.get(chosen_font_name, "")

        return {
            "version": "1.0", "midi_path": "", "lyrics_content": "", "output_video_path": "",
            "parameters": default_params
        }

    def _new_project_action(self):
        if not self._confirm_unsaved_changes(): return
        self.log_message("新規プロジェクト作成中...", "info")
        default_data = self._get_default_project_data()
        self._apply_project_data(default_data) # This will set loading_project_or_midi temporarily
        self.current_project_path = None
        self._set_project_modified_status(False) 
        # Window title updated by _set_project_modified_status
        # loading_project_or_midi will be reset by _on_midi_load_finished (or lack thereof)
        # called from within _apply_project_data -> _load_midi_file
        self.log_message("新規プロジェクトが作成されました。", "info")


    def _load_project_action(self):
        if not self._confirm_unsaved_changes(): return
        path, _ = QFileDialog.getOpenFileName(self, "プロジェクトをロード", self._get_last_dir(SETTINGS_LAST_PROJECT_DIR), PROJECT_FILE_FILTER)
        if path: self._load_project(path); self._set_last_dir(SETTINGS_LAST_PROJECT_DIR, os.path.dirname(path))
    def _save_project_action(self) -> bool: return self._save_project_as_action() if not self.current_project_path else self._save_project(self.current_project_path)
    def _save_project_as_action(self) -> bool:
        path, _ = QFileDialog.getSaveFileName(self, "名前を付けてプロジェクトを保存", self._get_last_dir(SETTINGS_LAST_PROJECT_DIR), PROJECT_FILE_FILTER)
        if path:
            if not path.lower().endswith(f".{PROJECT_FILE_EXTENSION}"): path += f".{PROJECT_FILE_EXTENSION}"
            if self._save_project(path): self._set_last_dir(SETTINGS_LAST_PROJECT_DIR, os.path.dirname(path)); return True
        return False 
    def _collect_project_data(self) -> Dict[str, Any]:
        data = {"version": "1.0", "midi_path": self.midi_path_edit.text(), "lyrics_content": self.lyrics_edit.toPlainText(), "output_video_path": self.output_video_path_edit.text()}
        params = {}
        spin_map = {'width':self.width_spin, 'height':self.height_spin, 'fps':self.fps_spin, 
                    'min_char_render_size':self.min_char_render_size_spin, 'max_char_render_size':self.max_char_render_size_spin,
                    'font_size_base':self.font_size_base_spin, 'char_spacing':self.char_spacing_spin,
                    'line_anchor_x':self.line_anchor_x_spin, 'line_anchor_y':self.line_anchor_y_spin,
                    'reference_pitch':self.reference_pitch_spin, 'reference_velocity':self.reference_velocity_spin,
                    'pitch_offset_scale':self.pitch_offset_scale_spin, 'pitch_size_scale':self.pitch_size_scale_spin,
                    'velocity_size_scale':self.velocity_size_scale_spin, 
                    'duration_padding_threshold_ticks':self.duration_padding_threshold_ticks_spin,
                    'duration_padding_scale_per_tick':self.duration_padding_scale_per_tick_spin}
        for k, w in spin_map.items(): params[k] = w.value()
        
        combo_map = {'font_name':self.font_combo, 'line_placement_mode':self.line_placement_mode_combo,
                     'line_h_align':self.line_h_align_combo, 'text_v_align':self.text_v_align_combo}
        for k, w in combo_map.items(): params[k] = w.currentText()

        params["font_path_if_custom"] = self.available_fonts.get(self.font_combo.currentText(), "")
        params["bg_color_rgb"] = self._get_color_from_button(self.bg_color_button)
        params["text_color_rgb"] = self._get_color_from_button(self.text_color_button)
        params["loaded_lyrics_display_path"] = self.loaded_lyrics_path_display.text()
        data["parameters"] = params
        return data
    def _apply_project_data(self, data: Dict[str, Any]):
        self.loading_project_or_midi = True 
        self._update_ui_states(is_loading_midi=True) # Indicate general loading
        try:
            self.midi_path = None; self.detailed_midi_notes_for_roll.clear(); self.raw_note_events_for_mapping_with_duration.clear()
            # self.piano_scene.clear_completely() # Will be handled by _load_midi_file if path is empty/invalid

            self.midi_path_edit.setText(data.get("midi_path", "")); self.output_video_path_edit.setText(data.get("output_video_path", ""))
            params = data.get("parameters", {})
            
            default_width = params.get("width",1920)
            default_height = params.get("height",1080)
            spin_map_defaults = {'width':default_width, 'height':default_height, 'fps':30, 'min_char_render_size':8,'max_char_render_size':300,
                                 'font_size_base':30, 'char_spacing':10, 
                                 'line_anchor_x':params.get("line_anchor_x", default_width//2), 
                                 'line_anchor_y':params.get("line_anchor_y", default_height//2), 
                                 'reference_pitch':60, 'reference_velocity':64,
                                 'pitch_offset_scale':0.0, 'pitch_size_scale':0.0, 'velocity_size_scale':0.0,
                                 'duration_padding_threshold_ticks':240, 'duration_padding_scale_per_tick':0.1}
            for k, default_val in spin_map_defaults.items():
                widget = getattr(self, k + "_spin", None) 
                if widget: widget.setValue(params.get(k, default_val))
            
            if hasattr(self.width_spin, "_previousValueForAnchor"): setattr(self.width_spin, "_previousValueForAnchor", self.width_spin.value())
            if hasattr(self.height_spin, "_previousValueForAnchor"): setattr(self.height_spin, "_previousValueForAnchor", self.height_spin.value())

            font_name = params.get("font_name"); font_path_custom = params.get("font_path_if_custom")
            if font_name:
                if font_name not in self.available_fonts and font_path_custom and os.path.exists(font_path_custom):
                    self.available_fonts[font_name] = font_path_custom; self._load_system_fonts_to_combo() 
                if self.font_combo.findText(font_name) != -1: self.font_combo.setCurrentText(font_name)
                elif self.font_combo.count() > 0: self.font_combo.setCurrentIndex(0) 
            
            combo_map_defaults = {'line_placement_mode':"動的配置", 'line_h_align':"中央揃え", 'text_v_align':"中央揃え"}
            for k, default_val in combo_map_defaults.items():
                getattr(self, k + "_combo").setCurrentText(params.get(k, default_val))

            self._update_color_button_style(self.bg_color_button, QColor.fromRgb(*params.get("bg_color_rgb", [0,0,0])))
            self._update_color_button_style(self.text_color_button, QColor.fromRgb(*params.get("text_color_rgb", [255,255,255])))
            self.loaded_lyrics_path_display.setText(params.get("loaded_lyrics_display_path",""))

            lyrics_content = data.get("lyrics_content", "")
            self.lyrics_edit.blockSignals(True)
            self.lyrics_edit.setText(lyrics_content)
            self.lyrics_edit.blockSignals(False)
            self._on_lyrics_debounced_change() 

            midi_to_load = data.get("midi_path")
            self._load_midi_file(midi_to_load or "") # This will handle empty/invalid path correctly and call _on_midi_load_finished

        except Exception as e: 
            self.log_message(f"プロジェクトデータ適用エラー: {e}", "error"); traceback.print_exc()
            self.loading_project_or_midi = False 
            self._update_ui_states() 
    def _save_project(self, filepath: str) -> bool:
        try:
            with open(filepath, 'w', encoding='utf-8') as f: json.dump(self._collect_project_data(), f, indent=4, ensure_ascii=False)
            self.current_project_path = filepath; self._set_project_modified_status(False); self.log_message(f"プロジェクト '{os.path.basename(filepath)}' 保存完了。", "info"); return True
        except Exception as e: self.log_message(f"プロジェクト保存エラー '{filepath}': {e}", "error"); QMessageBox.critical(self, "保存エラー", f"保存失敗:\n{e}"); return False
    def _load_project(self, filepath: str):
        try:
            with open(filepath, 'r', encoding='utf-8') as f: project_data = json.load(f)
            if project_data.get("version") != "1.0": self.log_message("プロジェクトファイルバージョン非互換/不明。", "warning")
            self._apply_project_data(project_data) 
            self.current_project_path = filepath; self._set_project_modified_status(False) 
            self.log_message(f"プロジェクト '{os.path.basename(filepath)}' ロード完了。", "info")
        except FileNotFoundError: self.log_message(f"ファイルが見つかりません: {filepath}", "error"); QMessageBox.critical(self, "ロードエラー", f"ファイル未発見:\n{filepath}"); self.loading_project_or_midi = False; self._update_ui_states()
        except json.JSONDecodeError as e: self.log_message(f"プロジェクトファイル解析エラー '{filepath}': {e}", "error"); QMessageBox.critical(self, "ロードエラー", f"形式無効:\n{e}"); self.loading_project_or_midi = False; self._update_ui_states()
        except Exception as e: self.log_message(f"プロジェクトロードエラー '{filepath}': {e}", "error"); traceback.print_exc(); QMessageBox.critical(self, "ロードエラー", f"予期せぬエラー:\n{e}"); self.loading_project_or_midi = False; self._update_ui_states()
    def closeEvent(self,event): 
        if not self._confirm_unsaved_changes(): event.ignore(); return
        for thread_attr_name, name in [("video_gen_thread","ビデオ生成"),("midi_load_thread","MIDI読込")]:
            thread_attr = getattr(self, thread_attr_name, None)
            if thread_attr and thread_attr.isRunning():
                if QMessageBox.question(self,"確認",f"{name}中です。終了しますか？",QMessageBox.Yes|QMessageBox.No,QMessageBox.No) == QMessageBox.Yes:
                    thread_attr.requestInterruption(); 
                    if not thread_attr.wait(1000): thread_attr.terminate(); thread_attr.wait()
                else: event.ignore(); return
        self._cleanup_temp_lyrics(); event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    QApplication.setOrganizationName(SETTINGS_ORGANIZATION_NAME); QApplication.setApplicationName(SETTINGS_APPLICATION_NAME)
    main_window = MainWindow(); main_window.show(); sys.exit(app.exec())

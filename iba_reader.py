"""
iba .dat File Reader for Python
================================
Reads iba PDA .dat files via the ibaAnalyzer COM interface.
Requirement: ibaAnalyzer must be installed (v8.x, 64-bit, Windows only).

Usage:
    reader = IbaReader(r"path\\to\\file.dat")
    reader.open()
    reader.print_signal_tree()
    df = reader.read_signals(["[60:0]", "[60:1]", "[60:2]"])
    reader.close()

Or as context manager:
    with IbaReader(filepath) as reader:
        df = reader.read_signals(["[60:0]", "[60:1]"])
"""

import fnmatch
import os
import re

import win32com.client
import pythoncom
import numpy as np
import pandas as pd
from datetime import datetime, timedelta


LCID = 0

# Filter types for GetSignalTree
FILTER_ALL_GROUPS = 0       # Groups/modules only
FILTER_ANALOG = 1           # Analog signals (float, with ':')
FILTER_DIGITAL = 2          # Digital signals (bool, with '.')
FILTER_ANALOG_DIGITAL = 3   # Analog + Digital
FILTER_TEXT = 4             # Text signals


class IbaReader:
    """Reads iba .dat files via the ibaAnalyzer COM interface."""

    def __init__(self, filepath):
        self.filepath = filepath
        self._app = None

    def open(self):
        """Opens the connection to ibaAnalyzer and loads the file."""
        self._app = win32com.client.dynamic.Dispatch('IBA.Analyzer')
        self._app.OpenDataFile(0, self.filepath)

    def close(self):
        """Closes the file and releases ibaAnalyzer."""
        if self._app is not None:
            try:
                self._app.CloseDataFile(0)
            except Exception:
                pass
            self._app = None

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    @property
    def version(self):
        """ibaAnalyzer version string."""
        return self._app.GetVersion

    def get_signal_list(self, filter_type=FILTER_ANALOG_DIGITAL):
        """
        Returns a list of all signals.

        Args:
            filter_type: FILTER_ANALOG (1), FILTER_DIGITAL (2),
                         FILTER_ANALOG_DIGITAL (3), FILTER_TEXT (4)

        Returns:
            List of dicts with 'id', 'name', 'group'
        """
        tree = self._app.GetSignalTree(filter_type)
        root = tree.GetRootNode()
        signals = []
        self._walk_tree(root, signals, group="")
        return signals

    def _parse_signal_name(self, tree_text):
        """Extracts the clean signal name from tree node text.

        Tree text format: '60:0: Motor_Speed_setpoint'
        Returns:          'Motor_Speed_setpoint'
        """
        parts = tree_text.split(": ", 1)
        return parts[1] if len(parts) > 1 else tree_text

    def _walk_tree(self, node, signals, group=""):
        """Recursively walk the signal tree."""
        while node is not None:
            text = node.Text
            ch_id = node.channelID
            if ch_id:
                signals.append({
                    'id': ch_id,
                    'name': self._parse_signal_name(text),
                    'group': group
                })
            try:
                node.Expand()
            except Exception:
                pass
            child = node.GetFirstChildNode()
            if child is not None:
                self._walk_tree(child, signals, group=text if not ch_id else group)
            node = node.GetSiblingNode()

    def get_signal_names(self, filter_type=FILTER_ANALOG_DIGITAL):
        """
        Returns a dict mapping signal IDs to their names.

        Args:
            filter_type: FILTER_ANALOG (1), FILTER_DIGITAL (2),
                         FILTER_ANALOG_DIGITAL (3), FILTER_TEXT (4)

        Returns:
            dict, e.g. {"[60:0]": "Motor_Speed_setpoint", ...}
        """
        signals = self.get_signal_list(filter_type)
        return {s['id']: s['name'] for s in signals}

    def signal_name(self, channel_id):
        """
        Returns the signal name for a given channel ID.

        Args:
            channel_id: e.g. "[60:0]" or "[64.0]"

        Returns:
            str, e.g. "Motor_Speed_setpoint"
        """
        # Strip brackets for lookup
        bare_id = channel_id.strip("[]")
        names = self.get_signal_names(FILTER_ANALOG_DIGITAL)
        # Also check text signals
        names.update(self.get_signal_names(FILTER_TEXT))
        for sig_id, name in names.items():
            if sig_id.strip("[]") == bare_id:
                return name
        return None

    def get_file_info(self):
        """
        Returns file metadata: start time, duration, sample rates, signal counts.

        Returns:
            dict with keys: version, start_time, duration_s, duration_h,
                            analog_count, digital_count, text_count,
                            sample_rates (set of unique timebases)
        """
        analog = self.get_signal_list(FILTER_ANALOG)
        digital = self.get_signal_list(FILTER_DIGITAL)
        text = self.get_signal_list(FILTER_TEXT)

        # Read timebase from first analog signal to determine duration
        sample_rates = set()
        duration_s = 0.0
        if analog:
            sig_id = analog[0]['id']
            if not sig_id.startswith('['):
                sig_id = f"[{sig_id}]"
            tb, _, data = self.read_signal(sig_id)
            sample_rates.add(tb)
            duration_s = len(data) * tb

        return {
            'version': self.version,
            'duration_s': duration_s,
            'duration_h': duration_s / 3600,
            'analog_count': len(analog),
            'digital_count': len(digital),
            'text_count': len(text),
            'total_count': len(analog) + len(digital) + len(text),
            'sample_rates': sample_rates,
        }

    def search_signals(self, pattern, filter_type=FILTER_ANALOG_DIGITAL):
        """
        Search signals by name pattern (wildcard or regex).

        Args:
            pattern: Wildcard pattern (e.g. "*Speed*", "*Motor*") or
                     regex pattern (e.g. "Motor_.*_actual")
            filter_type: Signal filter type (default: analog + digital)

        Returns:
            List of dicts with 'id', 'name', 'group' for matching signals
        """
        signals = self.get_signal_list(filter_type)

        # Try as wildcard first, fall back to regex
        if any(c in pattern for c in ['*', '?']):
            return [s for s in signals
                    if fnmatch.fnmatch(s['name'], pattern)]

        regex = re.compile(pattern, re.IGNORECASE)
        return [s for s in signals if regex.search(s['name'])]

    def read_all_signals(self, group=None, filter_type=FILTER_ANALOG):
        """
        Reads all signals (or all from a specific group) as DataFrame.

        Args:
            group: Optional group name filter (e.g. "60. analog_signals_01").
                   Supports wildcard patterns (e.g. "*analog_01*").
            filter_type: Signal filter type (default: analog only)

        Returns:
            pandas DataFrame with time index and one column per signal
        """
        signals = self.get_signal_list(filter_type)

        if group is not None:
            if any(c in group for c in ['*', '?']):
                signals = [s for s in signals
                           if fnmatch.fnmatch(s['group'], group)]
            else:
                signals = [s for s in signals if group in s['group']]

        if not signals:
            return pd.DataFrame()

        expressions = [s['id'] if s['id'].startswith('[') else f"[{s['id']}]"
                       for s in signals]
        return self.read_signals(expressions)

    def read_signal_range(self, expression, start_s, end_s, xtype=0):
        """
        Reads a signal for a specific time range only.

        Args:
            expression: Signal expression, e.g. "[60:0]"
            start_s: Start time in seconds
            end_s: End time in seconds
            xtype: 0 = time-based

        Returns:
            tuple (timebase_seconds, time_array, data_array)
        """
        tb, xoff, full_data = self.read_signal(expression, xtype)
        start_idx = max(0, int(start_s / tb))
        end_idx = min(len(full_data), int(end_s / tb))
        data = full_data[start_idx:end_idx]
        time = np.arange(len(data)) * tb + start_idx * tb
        return tb, time, data

    def read_signals_range(self, expressions, start_s, end_s, xtype=0):
        """
        Reads multiple signals for a specific time range as DataFrame.

        Args:
            expressions: List of signal expressions
            start_s: Start time in seconds
            end_s: End time in seconds
            xtype: 0 = time-based

        Returns:
            pandas DataFrame with time index
        """
        columns = {}
        time_arr = None

        for expr in expressions:
            tb, time, data = self.read_signal_range(expr, start_s, end_s, xtype)
            if time_arr is None:
                time_arr = time

            try:
                meta = self.get_channel_metadata(expr)
                col_name = meta['name']
                if meta['unit']:
                    col_name += f" [{meta['unit']}]"
            except Exception:
                col_name = expr

            columns[col_name] = data[:len(time_arr)]

        if time_arr is None:
            return pd.DataFrame()

        df = pd.DataFrame(columns, index=time_arr)
        df.index.name = 'time_s'
        return df

    def export_csv(self, expressions, path, separator=';'):
        """
        Exports signals to a CSV file.

        Args:
            expressions: List of signal expressions, e.g. ["[60:0]", "[60:1]"]
            path: Output file path (.csv)
            separator: CSV separator (default: ';')
        """
        df = self.read_signals(expressions)
        df.to_csv(path, sep=separator)

    def export_parquet(self, expressions, path):
        """
        Exports signals to a Parquet file (efficient columnar format).
        Requires pyarrow or fastparquet: pip install pyarrow

        Args:
            expressions: List of signal expressions
            path: Output file path (.parquet)
        """
        df = self.read_signals(expressions)
        try:
            df.to_parquet(path)
        except ImportError:
            raise ImportError(
                "Parquet export requires pyarrow or fastparquet. "
                "Install with: pip install pyarrow"
            )

    def print_signal_tree(self):
        """Prints the complete signal tree to the console."""
        print("=" * 70)
        print("SIGNALS in:", self.filepath)
        print("=" * 70)

        for label, ftype in [("ANALOG", FILTER_ANALOG),
                              ("DIGITAL", FILTER_DIGITAL),
                              ("TEXT", FILTER_TEXT)]:
            signals = self.get_signal_list(ftype)
            print(f"\n--- {label} ({len(signals)} signals) ---")
            for s in signals:
                print(f"  {s['id']} = \"{s['name']}\"  ({s['group']})")

    def get_channel_metadata(self, channel_id):
        """
        Returns metadata for a channel.

        Args:
            channel_id: e.g. "[60:0]"

        Returns:
            dict with name, unit, comment1, comment2
        """
        meta = self._app.GetChannelMetaData(channel_id)
        return {
            'name': meta.name,
            'unit': meta.Unit,
            'comment1': meta.Comment1,
            'comment2': meta.Comment2,
        }

    def read_signal(self, expression, xtype=0):
        """
        Reads a single signal as a numpy array.

        Args:
            expression: Signal expression, e.g. "[60:0]" or an
                        ibaAnalyzer expression like "Max([60:0])"
            xtype: 0 = time-based, 1 = length-based

        Returns:
            tuple (timebase_seconds, xoffset, data_array)
        """
        result = self._app._oleobj_.InvokeTypes(
            82,     # DispID for EvaluateToArray
            LCID,
            1,      # DISPATCH_METHOD
            (24, 0),  # Return: void
            ((8, 0), (3, 0), (16389, 2), (16389, 2), (16396, 2)),
            expression,
            xtype,
            0.0,    # pTimebase (output)
            0.0,    # xoffset (output)
            None    # pData (output)
        )
        timebase = result[0]
        xoffset = result[1]
        data = np.array(result[2], dtype=np.float32)
        return timebase, xoffset, data

    def read_signals(self, expressions, xtype=0):
        """
        Reads multiple signals and returns a pandas DataFrame.

        Args:
            expressions: List of signal expressions, e.g. ["[60:0]", "[60:1]"]
            xtype: 0 = time-based

        Returns:
            pandas DataFrame with time index and one column per signal
        """
        columns = {}
        timebase = None
        max_len = 0

        for expr in expressions:
            tb, xoff, data = self.read_signal(expr, xtype)
            if timebase is None:
                timebase = tb

            try:
                meta = self.get_channel_metadata(expr)
                col_name = meta['name']
                if meta['unit']:
                    col_name += f" [{meta['unit']}]"
            except Exception:
                col_name = expr

            columns[col_name] = data
            max_len = max(max_len, len(data))

        if timebase and timebase > 0:
            time_seconds = np.arange(max_len) * timebase
        else:
            time_seconds = np.arange(max_len)

        for key in columns:
            if len(columns[key]) < max_len:
                columns[key] = np.pad(
                    columns[key],
                    (0, max_len - len(columns[key])),
                    constant_values=np.nan
                )

        df = pd.DataFrame(columns, index=time_seconds)
        df.index.name = 'time_s'
        return df

    def read_text_signal(self, expression, xtype=0):
        """
        Reads a text signal.

        Args:
            expression: e.g. "[2:0]"

        Returns:
            tuple (timestamps_array, strings_list)
        """
        result = self._app._oleobj_.InvokeTypes(
            85,     # DispID for EvaluateToStringArray
            LCID,
            1,
            (24, 0),
            ((8, 0), (3, 0), (16396, 2), (16396, 2)),
            expression,
            xtype,
            None,   # pTimeStamps (output)
            None    # pStrings (output)
        )
        timestamps = np.array(result[0]) if result[0] else np.array([])
        strings = list(result[1]) if result[1] else []
        return timestamps, strings

    def evaluate(self, expression, xtype=0):
        """
        Evaluates an ibaAnalyzer expression and returns a scalar value.

        Args:
            expression: e.g. "Max([60:0])", "Min([60:0])", "Average([60:0])"

        Returns:
            float
        """
        return self._app.Evaluate(expression, xtype)

    def get_video_channels(self):
        """
        Returns a list of video (CaptureCam) channels in the file.

        Returns:
            List of dicts with 'id', 'name', 'group', 'frame_count',
            'duration_s', 'fps'
        """
        all_signals = self.get_signal_list(FILTER_ANALOG)
        video_channels = []

        for sig in all_signals:
            if 'CaptureCam' not in sig.get('group', ''):
                continue
            sig_id = sig['id']
            if not sig_id.startswith('['):
                sig_id = f"[{sig_id}]"

            try:
                tb, _, data = self.read_signal(sig_id)
                frame_count = len(data)
                duration_s = frame_count * tb
                fps = 1.0 / tb if tb > 0 else 0
                video_channels.append({
                    'id': sig_id,
                    'name': sig['name'],
                    'group': sig['group'],
                    'frame_count': frame_count,
                    'duration_s': duration_s,
                    'fps': fps,
                })
            except Exception:
                video_channels.append({
                    'id': sig_id,
                    'name': sig['name'],
                    'group': sig['group'],
                })

        return video_channels

    def export_video(self, output_path, channel_index=0):
        """
        Extracts embedded video from the .dat file to an MP4 file.

        The video data (ibaCapture/CaptureCam) is stored as a valid MP4 stream
        inside the PDA3 .dat file. This method locates and extracts it directly
        — no ibaCapture server or ibaAnalyzer GUI needed.

        Args:
            output_path: Output file path (should end in .mp4)
            channel_index: Which video channel to extract if multiple exist
                           (default: 0 = first channel)

        Returns:
            dict with 'path', 'size', 'name' of the exported video

        Raises:
            RuntimeError: If no video data is found in the .dat file
        """
        import struct

        fsize = os.path.getsize(self.filepath)

        # Locate the MP4 ftyp atom preceded by "ibaCaptureCAM" marker
        marker = b'ibaCaptureCAM'
        ftyp_sig = b'ftyp'
        chunk_size = 64 * 1024 * 1024
        found_videos = []

        with open(self.filepath, 'rb') as f:
            offset = 0
            while offset < fsize:
                f.seek(offset)
                chunk = f.read(chunk_size + 32)
                if not chunk:
                    break

                search_pos = 0
                while True:
                    idx = chunk.find(ftyp_sig, search_pos)
                    if idx == -1:
                        break
                    # Verify: 4 bytes before ftyp should be the atom size
                    if idx >= 4:
                        atom_size = struct.unpack('>I', chunk[idx-4:idx])[0]
                        if 16 <= atom_size <= 64:
                            abs_offset = offset + idx - 4
                            # Extract embedded filename
                            name = ''
                            if idx >= 22:
                                pre = chunk[idx-22:idx-4]
                                cam_idx = pre.find(marker)
                                if cam_idx != -1:
                                    name = pre[cam_idx:].split(b'\x00')[0].decode('ascii', errors='replace')
                            found_videos.append((abs_offset, name))
                    search_pos = idx + 4

                offset += chunk_size

        if not found_videos:
            raise RuntimeError("No embedded video (MP4) data found in this .dat file.")

        if channel_index >= len(found_videos):
            raise RuntimeError(
                f"Video channel index {channel_index} out of range. "
                f"Found {len(found_videos)} video(s): {[v[1] for v in found_videos]}"
            )

        video_start, video_name = found_videos[channel_index]

        # Find the end of the MP4 data by scanning atoms (ftyp, mdat*, moov)
        video_end = None
        with open(self.filepath, 'rb') as f:
            pos = video_start
            while pos < fsize:
                f.seek(pos)
                hdr = f.read(16)
                if len(hdr) < 8:
                    break
                size = struct.unpack('>I', hdr[0:4])[0]
                tag = hdr[4:8]
                if size == 1 and len(hdr) >= 16:
                    size = struct.unpack('>Q', hdr[8:16])[0]
                if size < 8:
                    break
                if tag == b'moov':
                    video_end = pos + size
                    break
                if tag not in (b'ftyp', b'mdat', b'free', b'skip', b'wide'):
                    break
                pos += size

        if video_end is None:
            raise RuntimeError("Could not determine end of video data (no moov atom found).")

        total = video_end - video_start
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

        with open(self.filepath, 'rb') as src:
            src.seek(video_start)
            with open(output_path, 'wb') as dst:
                copied = 0
                while copied < total:
                    to_read = min(chunk_size, total - copied)
                    data = src.read(to_read)
                    if not data:
                        break
                    dst.write(data)
                    copied += len(data)

        return {
            'path': output_path,
            'size': total,
            'name': video_name or f'video_{channel_index}',
        }


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python iba_reader.py <path_to_dat_file>")
        sys.exit(1)

    DAT_FILE = sys.argv[1]

    with IbaReader(DAT_FILE) as reader:
        print(f"ibaAnalyzer: {reader.version}\n")
        reader.print_signal_tree()

        # Read first analog signal as demo
        signals = reader.get_signal_list(FILTER_ANALOG)
        if signals:
            first_id = f"[{signals[0]['id']}]"
            print(f"\n--- Reading signal {first_id} ---")
            tb, xoff, data = reader.read_signal(first_id)
            print(f"  Sample rate: {tb}s ({1/tb:.0f} Hz)")
            print(f"  Data points: {len(data):,}")
            print(f"  Duration:    {len(data) * tb / 3600:.1f} hours")
            print(f"  Min: {np.nanmin(data):.2f}, Max: {np.nanmax(data):.2f}, Mean: {np.nanmean(data):.2f}")

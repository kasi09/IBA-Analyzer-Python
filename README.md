# IBA-Analyzer-Python

Python library to read **iba PDA .dat files** via the ibaAnalyzer COM interface.

## Requirements

- **Windows** (COM interface)
- **ibaAnalyzer** v8.x (64-bit) installed
- **Python 3.8+** (64-bit)
- `pywin32`, `numpy`, `pandas`

```bash
pip install pywin32 numpy pandas
```

For Parquet export (optional):
```bash
pip install pyarrow
```


## Quick Start

```python
from iba_reader import IbaReader

with IbaReader(r"D:\data\measurement.dat") as reader:
    # File overview
    info = reader.get_file_info()
    print(f"Duration: {info['duration_h']:.1f}h, Signals: {info['total_count']}")

    # List all signals
    reader.print_signal_tree()

    # Search signals by pattern
    results = reader.search_signals("*Speed*")

    # Read signals as pandas DataFrame
    df = reader.read_signals(["[60:0]", "[60:1]", "[60:2]"])

    # Read all signals from a group
    df = reader.read_all_signals(group="*analog_01*")

    # Read only a time range (e.g. first 60 seconds)
    tb, time, data = reader.read_signal_range("[60:0]", start_s=0, end_s=60)

    # Read text signals
    timestamps, strings = reader.read_text_signal("[2:0]")

    # Evaluate ibaAnalyzer expressions
    max_val = reader.evaluate("Max([60:0])")

    # Export to CSV or Parquet
    reader.export_csv(["[60:0]", "[60:1]"], "output.csv")
    reader.export_parquet(["[60:0]", "[60:1]"], "output.parquet")  # needs pyarrow
```

## CLI Usage

```bash
python iba_reader.py path\to\file.dat
```

Prints the ibaAnalyzer version, the full signal tree, and basic statistics for the first analog signal.

## API Reference

### `IbaReader(filepath)`

Create a reader instance for a .dat file. Use as context manager or call `open()`/`close()` manually.

```python
# Context manager (recommended)
with IbaReader(r"path\to\file.dat") as reader:
    ...

# Manual open/close
reader = IbaReader(r"path\to\file.dat")
reader.open()
# ... work with reader ...
reader.close()
```

---

### File Information

#### `version`

Returns the ibaAnalyzer version string.

```python
print(reader.version)  # e.g. "8.3.4"
```

#### `get_file_info()`

Returns file metadata as a dict.

```python
info = reader.get_file_info()
# {
#     'version': '8.3.4',
#     'duration_s': 86400.0,
#     'duration_h': 24.0,
#     'analog_count': 150,
#     'digital_count': 80,
#     'text_count': 5,
#     'total_count': 235,
#     'sample_rates': {0.01}
# }
```

---

### Signal Discovery

#### `get_signal_list(filter_type=3)`

Returns a list of all signals as dicts with `id`, `name`, `group`.

```python
# All analog + digital signals (default)
signals = reader.get_signal_list()

# Only analog signals
analog = reader.get_signal_list(filter_type=1)

# Only digital signals
digital = reader.get_signal_list(filter_type=2)

# Only text signals
text = reader.get_signal_list(filter_type=4)

for s in signals[:5]:
    print(f"{s['id']} = {s['name']}  (Group: {s['group']})")
```

Filter type constants:
| Value | Constant | Description |
|-------|----------|-------------|
| 0 | `FILTER_ALL_GROUPS` | Groups/modules only |
| 1 | `FILTER_ANALOG` | Analog signals (float) |
| 2 | `FILTER_DIGITAL` | Digital signals (bool) |
| 3 | `FILTER_ANALOG_DIGITAL` | Analog + Digital |
| 4 | `FILTER_TEXT` | Text signals |

#### `get_signal_names(filter_type=3)`

Returns a dict mapping signal IDs to their names.

```python
names = reader.get_signal_names()
# {"[60:0]": "Motor_Speed_setpoint", ...}
```

#### `signal_name(channel_id)`

Returns the signal name for a single channel ID.

```python
name = reader.signal_name("[60:0]")
# "Motor_Speed_setpoint"
```

#### `search_signals(pattern, filter_type=3)`

Search signals by wildcard pattern or regex.

```python
# Wildcard search
results = reader.search_signals("*Speed*")
results = reader.search_signals("*Motor*")

# Regex search (if no wildcard characters)
results = reader.search_signals("Motor_.*_actual")

for s in results:
    print(f"{s['id']} = {s['name']}")
```

#### `print_signal_tree()`

Prints all signals grouped by type (analog, digital, text) to the console.

```python
reader.print_signal_tree()
# ======================================================================
# SIGNALS in: D:\data\measurement.dat
# ======================================================================
#
# --- ANALOG (150 signals) ---
#   [60:0] = "Motor_Speed_setpoint"  (60. analog_signals_01)
#   [60:1] = "Motor_Speed_actual"   (60. analog_signals_01)
#   ...
```

#### `get_channel_metadata(channel_id)`

Returns metadata for a channel: name, unit, and comments.

```python
meta = reader.get_channel_metadata("[60:0]")
# {
#     'name': 'Motor_Speed_setpoint',
#     'unit': 'mm',
#     'comment1': '...',
#     'comment2': '...'
# }
```

---

### Reading Signals

#### `read_signal(expression, xtype=0)`

Reads a single signal as a numpy array. Returns `(timebase_seconds, xoffset, data_array)`.

```python
tb, xoff, data = reader.read_signal("[60:0]")
print(f"Sample rate: {tb}s ({1/tb:.0f} Hz)")
print(f"Data points: {len(data):,}")
print(f"Duration: {len(data) * tb / 3600:.1f} hours")
print(f"Min: {data.min():.2f}, Max: {data.max():.2f}")
```

#### `read_signals(expressions, xtype=0)`

Reads multiple signals and returns a pandas DataFrame with a time index.

```python
df = reader.read_signals(["[60:0]", "[60:1]", "[60:2]"])
print(df.head())
#          Speed_setpoint [mm]  Speed_actual [mm]  Velocity [mm/s]
# time_s
# 0.00                  0.0               0.0                 0.0
# 0.01                  0.0               0.0                 0.0
# ...

print(df.describe())
```

Column names are automatically resolved from channel metadata (name + unit).

#### `read_all_signals(group=None, filter_type=1)`

Reads all signals (or all from a specific group) as a DataFrame.

```python
# Read ALL analog signals
df = reader.read_all_signals()

# Read signals from a specific group (wildcard supported)
df = reader.read_all_signals(group="*analog_01*")

# Read signals with exact group name match
df = reader.read_all_signals(group="60. analog_signals_01")
```

#### `read_signal_range(expression, start_s, end_s, xtype=0)`

Reads a signal for a specific time range only. Useful for large files.

```python
# Read only the first 60 seconds
tb, time, data = reader.read_signal_range("[60:0]", start_s=0, end_s=60)

# Read from minute 5 to minute 10
tb, time, data = reader.read_signal_range("[60:0]", start_s=300, end_s=600)
```

#### `read_signals_range(expressions, start_s, end_s, xtype=0)`

Reads multiple signals for a specific time range as a DataFrame.

```python
df = reader.read_signals_range(
    ["[60:0]", "[60:1]", "[60:2]"],
    start_s=0,
    end_s=60
)
```

#### `read_text_signal(expression, xtype=0)`

Reads a text signal. Returns `(timestamps_array, strings_list)`.

```python
timestamps, strings = reader.read_text_signal("[2:0]")
for i in range(min(5, len(strings))):
    print(f"t={timestamps[i]:.1f}s: {strings[i]}")
```

---

### Expressions & Calculations

#### `evaluate(expression, xtype=0)`

Evaluates any ibaAnalyzer expression and returns a scalar float value.

```python
max_val = reader.evaluate("Max([60:0])")
min_val = reader.evaluate("Min([60:0])")
avg_val = reader.evaluate("Average([60:0])")
diff    = reader.evaluate("[60:0] - [60:1]")
```

You can also use expressions directly in `read_signal()`:

```python
tb, xoff, data = reader.read_signal("[60:0] - [60:1]")
```

---

### Export

#### `export_csv(expressions, path, separator=';')`

Exports signals to a CSV file.

```python
reader.export_csv(["[60:0]", "[60:1]"], "output.csv")
reader.export_csv(["[60:0]", "[60:1]"], "output.csv", separator=',')
```

#### `export_parquet(expressions, path)`

Exports signals to a Parquet file (efficient columnar format). Requires `pyarrow`.

```python
reader.export_parquet(["[60:0]", "[60:1]"], "output.parquet")
```

---

### Video

#### `get_video_channels()`

Returns a list of video (CaptureCam) channels found in the .dat file.

```python
channels = reader.get_video_channels()
for ch in channels:
    print(f"{ch['id']} = {ch['name']}")
    print(f"  Frames: {ch['frame_count']:,}, Duration: {ch['duration_s']/3600:.1f}h, FPS: {ch['fps']:.0f}")
```

#### `export_video(output_path, channel_index=0)`

Extracts embedded video (ibaCapture/CaptureCam) from the .dat file to an MP4 file.
The video data is stored as a valid MP4 stream inside the PDA3 file and is extracted
directly — no ibaCapture server or ibaAnalyzer GUI needed.

```python
result = reader.export_video("output.mp4")
print(f"Exported: {result['name']} ({result['size'] / 1024**3:.1f} GB)")
```

---

### Signal ID Format

- **Analog**: `[module:channel]` e.g. `[60:0]`
- **Digital**: `[module.channel]` e.g. `[64.0]`
- **Expressions**: any ibaAnalyzer expression, e.g. `Max([60:0])`, `[60:0] - [60:1]`

## How It Works

This library uses the **ibaAnalyzer COM automation interface** (`IBA.Analyzer`) to read .dat files. ibaAnalyzer runs as a background COM server — no GUI window is opened.

The COM interface was reverse-engineered from the type library embedded in `ibaAnalyzer.exe`.

## Support

<a href="https://www.buymeacoffee.com/kasi09" target="_blank"><img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" alt="Buy Me A Coffee" style="height: 41px !important;width: 174px !important;box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;-webkit-box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;" ></a>

## License

MIT

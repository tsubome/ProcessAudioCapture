# ProcessAudioCapture

A Windows DLL for capturing audio from specific processes (applications).

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Windows](https://img.shields.io/badge/platform-Windows%2010%202004+-blue.svg)](https://www.microsoft.com/windows)
[![PyPI](https://img.shields.io/pypi/v/process-audio-capture.svg)](https://pypi.org/project/process-audio-capture/)

## Overview

ProcessAudioCapture uses the `AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS` API introduced in Windows 10 version 2004 (Build 19041) to capture audio output from a specific process only.

Unlike system-wide loopback capture (which records all audio), this library allows you to:
- Record audio from a single application (e.g., browser, game, media player)
- Exclude a specific application's audio from recording
- Monitor real-time audio levels from a specific process

Based on Microsoft's official [ApplicationLoopback](https://github.com/microsoft/windows-classic-samples/tree/main/Samples/ApplicationLoopback) sample.

## Features

- **Process-specific capture** - Record audio from only the specified process (INCLUDE mode)
- **Process exclusion** - Record all audio except from the specified process (EXCLUDE mode)
- **Real-time level monitoring** - Get audio level in dB via callback
- **Pause/Resume support** - Pause and resume recording without stopping
- **WAV file output** - Direct output to WAV file
- **Process enumeration** - List all processes currently outputting audio
- **Python bindings included** - Ready-to-use Python wrapper

## Requirements

- **OS**: Windows 10 version 2004 (Build 19041) or later, or Windows 11
- **Build Environment** (if building from source):
  - Visual Studio 2019/2022 with C++ Desktop Development
  - Windows 10 SDK (10.0.19041.0 or later)
  - CMake 3.16 or later

## Installation

### Python (Recommended)

```bash
pip install process-audio-capture
```

### Using Prebuilt Binaries

Download the latest release from the [Releases](https://github.com/tsubome/ProcessAudioCapture/releases) page.

Extract and use:
- `ProcessAudioCapture.dll` - The main DLL

### Building from Source

```bash
cd ProcessAudioCapture
mkdir build
cd build
cmake .. -G "Visual Studio 17 2022" -A x64
cmake --build . --config Release
```

Output files:
- `build/bin/Release/ProcessAudioCapture.dll`
- `build/bin/Release/TestCapture.exe`

## Quick Start

### Python

```python
from process_audio_capture import ProcessAudioCapture

# Check if Windows version is supported
if not ProcessAudioCapture.is_supported():
    print("Windows 10 2004 or later required")
    exit()

# List processes outputting audio
processes = ProcessAudioCapture.enumerate_audio_processes()
for p in processes:
    print(f"PID: {p.pid}, Name: {p.name}, Window: {p.window_title}")

# Capture audio from a specific process
def on_level(level_db):
    print(f"Level: {level_db:.1f} dB")

with ProcessAudioCapture(
    pid=target_pid,
    output_path="output.wav",
    level_callback=on_level
) as capture:
    capture.start()
    time.sleep(10)  # Record for 10 seconds
    # capture.stop() is called automatically
```

### C/C++

```cpp
#include "ProcessAudioCapture.h"

// Check Windows version
if (!PacIsSupported()) {
    printf("Windows 10 2004 or later required\n");
    return;
}

// List audio processes
PacProcessInfo processes[64];
int count;
PacEnumerateAudioProcesses(processes, 64, &count);

// Start capture
PacHandle handle;
PacErrorCode err = PacStartCapture(
    targetPid,           // Target process ID
    PAC_MODE_INCLUDE,    // Capture mode
    L"output.wav",       // Output file (NULL for monitoring only)
    levelCallback,       // Level callback (can be NULL)
    nullptr,             // User data
    &handle
);

// During capture...
float levelDb;
PacGetLevel(handle, &levelDb);

// Pause/Resume
PacPauseCapture(handle);
// ... paused ...
PacResumeCapture(handle);

// Stop capture
PacStopCapture(handle);
```

## API Reference

### Functions

| Function | Description |
|----------|-------------|
| `PacIsSupported()` | Check if Windows version is supported |
| `PacEnumerateAudioProcesses()` | Get list of processes outputting audio |
| `PacStartCapture()` | Start capturing audio from a process |
| `PacStopCapture()` | Stop capturing |
| `PacPauseCapture()` | Pause capturing (audio data not written) |
| `PacResumeCapture()` | Resume capturing |
| `PacIsPaused()` | Check if capture is paused |
| `PacGetLevel()` | Get current audio level in dB |
| `PacIsCapturing()` | Check if currently capturing |
| `PacGetLastErrorMessage()` | Get last error message |
| `PacGetVersion()` | Get DLL version |

### Error Codes

| Code | Value | Description |
|------|-------|-------------|
| `PAC_SUCCESS` | 0 | Success |
| `PAC_ERROR_INVALID_PARAM` | -1 | Invalid parameter |
| `PAC_ERROR_NOT_SUPPORTED` | -2 | Windows version not supported |
| `PAC_ERROR_PROCESS_NOT_FOUND` | -3 | Process not found |
| `PAC_ERROR_AUDIO_INIT_FAILED` | -4 | Audio initialization failed |
| `PAC_ERROR_FILE_CREATE_FAILED` | -5 | File creation failed |
| `PAC_ERROR_ALREADY_RECORDING` | -6 | Already recording |
| `PAC_ERROR_NOT_RECORDING` | -7 | Not recording |

### Capture Modes

| Mode | Description |
|------|-------------|
| `PAC_MODE_INCLUDE` | Capture only the specified process's audio |
| `PAC_MODE_EXCLUDE` | Capture all audio except the specified process |

## Directory Structure

```
ProcessAudioCapture/
├── CMakeLists.txt          # Build configuration
├── LICENSE                 # MIT License
├── README.md               # This file
├── src/
│   ├── ProcessAudioCapture.h   # Public API header
│   ├── ProcessAudioCapture.cpp # DLL implementation
│   ├── LoopbackCapture.h       # Internal capture header
│   └── LoopbackCapture.cpp     # Internal capture implementation
├── python/
│   └── process_audio_capture.py  # Python wrapper
├── test/
│   └── test_capture.cpp    # C++ test application
├── prebuilt/               # Prebuilt binaries
│   └── x64/
│       └── ProcessAudioCapture.dll
└── build/                  # Build output (CMake generated)
```

## Limitations

1. **DRM-protected content**: Cannot capture (Windows restriction)
2. **Audio format**: Fixed at 48kHz or device default, 32-bit float, stereo
3. **Child processes**: Audio from child processes of the target is also included/excluded

## Use Cases

- Recording game audio without microphone
- Capturing browser audio (meetings, videos, music)
- Audio processing pipelines
- Accessibility tools
- Streaming/broadcasting specific application audio

## License

MIT License - see [LICENSE](LICENSE) file.

Based on Microsoft's [ApplicationLoopback](https://github.com/microsoft/windows-classic-samples/tree/main/Samples/ApplicationLoopback) sample.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- Microsoft for the ApplicationLoopback sample code
- Windows Audio Session API (WASAPI) documentation

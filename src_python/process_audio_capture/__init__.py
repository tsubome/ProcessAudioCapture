"""
ProcessAudioCapture — クロスプラットフォーム音声キャプチャライブラリ

Windows: WASAPI AudioClient Process Loopback (DLL)
macOS:   ScreenCaptureKit (pyobjc)
"""
import sys

if sys.platform == "win32":
    from ._windows import (
        ProcessAudioCapture,
        AudioProcess,
        ProcessAudioCaptureError,
        PacCaptureMode,
        PacErrorCode,
    )
elif sys.platform == "darwin":
    from ._macos import (
        ProcessAudioCapture,
        AudioProcess,
        ProcessAudioCaptureError,
        PacCaptureMode,
        PacErrorCode,
    )
else:
    # Linux等: 未対応
    from dataclasses import dataclass

    @dataclass
    class AudioProcess:
        pid: int = 0
        name: str = ""
        window_title: str = ""

    class ProcessAudioCaptureError(Exception):
        pass

    class PacCaptureMode:
        INCLUDE = 0
        EXCLUDE = 1

    class PacErrorCode:
        SUCCESS = 0
        NOT_SUPPORTED = -1

    class ProcessAudioCapture:
        INCLUDE = 0
        EXCLUDE = 1

        @staticmethod
        def is_supported():
            return False

        @staticmethod
        def enumerate_audio_processes():
            return []

        def __init__(self, *args, **kwargs):
            raise ProcessAudioCaptureError("ProcessAudioCapture is not supported on this platform")

__all__ = [
    "ProcessAudioCapture",
    "AudioProcess",
    "ProcessAudioCaptureError",
    "PacCaptureMode",
    "PacErrorCode",
]

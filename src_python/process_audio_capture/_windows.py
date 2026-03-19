"""
ProcessAudioCapture - Windows DLL for capturing audio from specific processes

Usage:
    from process_audio_capture import ProcessAudioCapture, AudioProcess

    # Check if Windows version is supported
    if not ProcessAudioCapture.is_supported():
        print("Windows 10 2004 or later required")
        exit()

    # List processes outputting audio
    processes = ProcessAudioCapture.enumerate_audio_processes()
    for p in processes:
        print(f"PID: {p.pid}, Name: {p.name}")

    # Capture audio from a specific process
    with ProcessAudioCapture(pid=target_pid, output_path="output.wav") as capture:
        capture.start()
        time.sleep(10)  # Record for 10 seconds
"""

__version__ = "1.0.0"
__author__ = "tsubome"

import ctypes
from ctypes import wintypes
from dataclasses import dataclass
from typing import List, Optional, Callable
from pathlib import Path
import threading


class PacErrorCode:
    """Error codes returned by the DLL"""
    SUCCESS = 0
    INVALID_PARAM = -1
    NOT_SUPPORTED = -2
    PROCESS_NOT_FOUND = -3
    AUDIO_INIT_FAILED = -4
    FILE_CREATE_FAILED = -5
    ALREADY_RECORDING = -6
    NOT_RECORDING = -7
    UNKNOWN = -100


class PacCaptureMode:
    """Capture mode constants"""
    INCLUDE = 0  # Capture only the specified process's audio
    EXCLUDE = 1  # Capture all audio except the specified process


class _PacProcessInfo(ctypes.Structure):
    """Internal structure for process info"""
    _fields_ = [
        ("processId", wintypes.DWORD),
        ("processName", ctypes.c_wchar * 260),
        ("windowTitle", ctypes.c_wchar * 260),
    ]


_LevelCallbackFunc = ctypes.CFUNCTYPE(None, ctypes.c_float, ctypes.c_void_p)


@dataclass
class AudioProcess:
    """Information about a process outputting audio"""
    pid: int
    name: str
    window_title: str


class ProcessAudioCaptureError(Exception):
    """Exception raised by ProcessAudioCapture operations"""
    def __init__(self, code: int, message: str):
        self.code = code
        self.message = message
        super().__init__(f"Error {code}: {message}")


class ProcessAudioCapture:
    """
    Capture audio from a specific Windows process.

    This class uses the AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS API introduced
    in Windows 10 version 2004 to capture audio from individual processes.

    Example:
        # List processes outputting audio
        processes = ProcessAudioCapture.enumerate_audio_processes()
        for p in processes:
            print(f"{p.pid}: {p.name} - {p.window_title}")

        # Capture audio from a specific process
        with ProcessAudioCapture(pid=1234, output_path="output.wav") as capture:
            capture.start()
            time.sleep(10)  # Record for 10 seconds
            capture.stop()

    Attributes:
        is_capturing (bool): Whether capture is currently active
        is_paused (bool): Whether capture is paused
        level_db (float): Current audio level in dB
    """

    _dll: Optional[ctypes.WinDLL] = None
    _dll_path: Optional[Path] = None

    @classmethod
    def _load_dll(cls, dll_path: Optional[Path] = None) -> ctypes.WinDLL:
        """Load the DLL"""
        if cls._dll is not None and (dll_path is None or dll_path == cls._dll_path):
            return cls._dll

        if dll_path is None:
            # Default path: same directory as this module
            module_dir = Path(__file__).parent
            dll_path = module_dir / "ProcessAudioCapture.dll"

            if not dll_path.exists():
                raise FileNotFoundError(
                    f"ProcessAudioCapture.dll not found at {dll_path}. "
                    "Make sure the DLL is installed with the package."
                )

        cls._dll = ctypes.WinDLL(str(dll_path))
        cls._dll_path = dll_path
        cls._setup_function_signatures(cls._dll)
        return cls._dll

    @classmethod
    def _setup_function_signatures(cls, dll: ctypes.WinDLL):
        """Set up DLL function signatures"""
        dll.PacIsSupported.argtypes = []
        dll.PacIsSupported.restype = wintypes.BOOL

        dll.PacEnumerateAudioProcesses.argtypes = [
            ctypes.POINTER(_PacProcessInfo),
            ctypes.c_int,
            ctypes.POINTER(ctypes.c_int)
        ]
        dll.PacEnumerateAudioProcesses.restype = ctypes.c_int

        dll.PacStartCapture.argtypes = [
            wintypes.DWORD,
            ctypes.c_int,
            ctypes.c_wchar_p,
            _LevelCallbackFunc,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p)
        ]
        dll.PacStartCapture.restype = ctypes.c_int

        dll.PacStopCapture.argtypes = [ctypes.c_void_p]
        dll.PacStopCapture.restype = ctypes.c_int

        dll.PacGetLevel.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_float)]
        dll.PacGetLevel.restype = ctypes.c_int

        dll.PacIsCapturing.argtypes = [ctypes.c_void_p]
        dll.PacIsCapturing.restype = wintypes.BOOL

        dll.PacGetLastErrorMessage.argtypes = [ctypes.c_wchar_p, ctypes.c_int]
        dll.PacGetLastErrorMessage.restype = None

        dll.PacGetVersion.argtypes = []
        dll.PacGetVersion.restype = ctypes.c_char_p

        dll.PacPauseCapture.argtypes = [ctypes.c_void_p]
        dll.PacPauseCapture.restype = ctypes.c_int

        dll.PacResumeCapture.argtypes = [ctypes.c_void_p]
        dll.PacResumeCapture.restype = ctypes.c_int

        dll.PacIsPaused.argtypes = [ctypes.c_void_p]
        dll.PacIsPaused.restype = wintypes.BOOL

    @classmethod
    def is_supported(cls, dll_path: Optional[Path] = None) -> bool:
        """
        Check if the current Windows version supports process-specific capture.

        Returns:
            bool: True if Windows 10 version 2004 or later
        """
        try:
            dll = cls._load_dll(dll_path)
            return bool(dll.PacIsSupported())
        except FileNotFoundError:
            return False

    @classmethod
    def get_version(cls, dll_path: Optional[Path] = None) -> str:
        """
        Get the DLL version string.

        Returns:
            str: Version string (e.g., "1.0.0")
        """
        dll = cls._load_dll(dll_path)
        return dll.PacGetVersion().decode('utf-8')

    @classmethod
    def enumerate_audio_processes(cls, dll_path: Optional[Path] = None) -> List[AudioProcess]:
        """
        Get a list of processes currently outputting audio.

        Returns:
            List[AudioProcess]: List of processes with audio output
        """
        dll = cls._load_dll(dll_path)

        max_count = 64
        processes = (_PacProcessInfo * max_count)()
        actual_count = ctypes.c_int(0)

        err = dll.PacEnumerateAudioProcesses(processes, max_count, ctypes.byref(actual_count))

        if err != PacErrorCode.SUCCESS:
            raise ProcessAudioCaptureError(err, cls._get_last_error())

        result = []
        for i in range(actual_count.value):
            p = processes[i]
            result.append(AudioProcess(
                pid=p.processId,
                name=p.processName,
                window_title=p.windowTitle
            ))

        return result

    @classmethod
    def _get_last_error(cls) -> str:
        """Get the last error message from the DLL"""
        if cls._dll is None:
            return "DLL not loaded"

        buffer = ctypes.create_unicode_buffer(512)
        cls._dll.PacGetLastErrorMessage(buffer, 512)
        return buffer.value

    def __init__(
        self,
        pid: int,
        output_path: Optional[str] = None,
        mode: int = PacCaptureMode.INCLUDE,
        level_callback: Optional[Callable[[float], None]] = None,
        dll_path: Optional[Path] = None
    ):
        """
        Initialize a capture session for a specific process.

        Args:
            pid: Target process ID
            output_path: Path to output WAV file (None for monitoring only)
            mode: Capture mode (INCLUDE or EXCLUDE)
            level_callback: Callback function receiving audio level in dB
            dll_path: Custom path to the DLL (optional)
        """
        self._dll = self._load_dll(dll_path)
        self._pid = pid
        self._output_path = output_path
        self._mode = mode
        self._user_callback = level_callback
        self._handle: Optional[ctypes.c_void_p] = None
        self._callback_ref: Optional[_LevelCallbackFunc] = None
        self._lock = threading.Lock()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
        return False

    def start(self):
        """
        Start capturing audio.

        Raises:
            ProcessAudioCaptureError: If capture fails to start
        """
        with self._lock:
            if self._handle is not None:
                raise ProcessAudioCaptureError(
                    PacErrorCode.ALREADY_RECORDING,
                    "Already capturing"
                )

            if self._user_callback:
                def _callback(level_db: float, user_data: ctypes.c_void_p):
                    if self._user_callback:
                        self._user_callback(level_db)

                self._callback_ref = _LevelCallbackFunc(_callback)
            else:
                self._callback_ref = _LevelCallbackFunc(0)

            handle = ctypes.c_void_p()
            err = self._dll.PacStartCapture(
                wintypes.DWORD(self._pid),
                ctypes.c_int(self._mode),
                self._output_path,
                self._callback_ref,
                None,
                ctypes.byref(handle)
            )

            if err != PacErrorCode.SUCCESS:
                raise ProcessAudioCaptureError(err, self._get_last_error())

            self._handle = handle

    def stop(self):
        """Stop capturing audio."""
        with self._lock:
            if self._handle is None:
                return

            err = self._dll.PacStopCapture(self._handle)
            self._handle = None
            self._callback_ref = None

            if err != PacErrorCode.SUCCESS:
                raise ProcessAudioCaptureError(err, self._get_last_error())

    def pause(self):
        """
        Pause capturing (audio data not written to file).

        Raises:
            ProcessAudioCaptureError: If not currently capturing
        """
        with self._lock:
            if self._handle is None:
                raise ProcessAudioCaptureError(
                    PacErrorCode.NOT_RECORDING,
                    "Not capturing"
                )

            err = self._dll.PacPauseCapture(self._handle)
            if err != PacErrorCode.SUCCESS:
                raise ProcessAudioCaptureError(err, self._get_last_error())

    def resume(self):
        """
        Resume capturing after pause.

        Raises:
            ProcessAudioCaptureError: If not currently capturing
        """
        with self._lock:
            if self._handle is None:
                raise ProcessAudioCaptureError(
                    PacErrorCode.NOT_RECORDING,
                    "Not capturing"
                )

            err = self._dll.PacResumeCapture(self._handle)
            if err != PacErrorCode.SUCCESS:
                raise ProcessAudioCaptureError(err, self._get_last_error())

    @property
    def is_capturing(self) -> bool:
        """Check if currently capturing."""
        with self._lock:
            if self._handle is None:
                return False
            return bool(self._dll.PacIsCapturing(self._handle))

    @property
    def is_paused(self) -> bool:
        """Check if capture is paused."""
        with self._lock:
            if self._handle is None:
                return False
            return bool(self._dll.PacIsPaused(self._handle))

    @property
    def level_db(self) -> float:
        """Get current audio level in dB."""
        with self._lock:
            if self._handle is None:
                return -60.0

            level = ctypes.c_float()
            err = self._dll.PacGetLevel(self._handle, ctypes.byref(level))

            if err != PacErrorCode.SUCCESS:
                return -60.0

            return level.value


# Convenience exports
__all__ = [
    "ProcessAudioCapture",
    "AudioProcess",
    "ProcessAudioCaptureError",
    "PacCaptureMode",
    "PacErrorCode",
    "__version__",
]

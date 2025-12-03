"""
ProcessAudioCapture Python Wrapper
プロセス別音声キャプチャ DLL の Python バインディング
"""

import ctypes
from ctypes import wintypes
from dataclasses import dataclass
from typing import List, Optional, Callable
from pathlib import Path
import threading


# エラーコード
class PacErrorCode:
    SUCCESS = 0
    INVALID_PARAM = -1
    NOT_SUPPORTED = -2
    PROCESS_NOT_FOUND = -3
    AUDIO_INIT_FAILED = -4
    FILE_CREATE_FAILED = -5
    ALREADY_RECORDING = -6
    NOT_RECORDING = -7
    UNKNOWN = -100


# キャプチャモード
class PacCaptureMode:
    INCLUDE = 0  # 指定プロセスの音声のみ
    EXCLUDE = 1  # 指定プロセス以外の音声


# プロセス情報構造体
class PacProcessInfo(ctypes.Structure):
    _fields_ = [
        ("processId", wintypes.DWORD),
        ("processName", ctypes.c_wchar * 260),
        ("windowTitle", ctypes.c_wchar * 260),
    ]


# コールバック関数型
LevelCallbackFunc = ctypes.CFUNCTYPE(None, ctypes.c_float, ctypes.c_void_p)


@dataclass
class AudioProcess:
    """音声を出力しているプロセスの情報"""
    pid: int
    name: str
    window_title: str


class ProcessAudioCaptureError(Exception):
    """ProcessAudioCapture エラー"""
    def __init__(self, code: int, message: str):
        self.code = code
        self.message = message
        super().__init__(f"Error {code}: {message}")


class ProcessAudioCapture:
    """
    プロセス別音声キャプチャ

    使用例:
        # 音声を出力しているプロセス一覧を取得
        processes = ProcessAudioCapture.enumerate_audio_processes()
        for p in processes:
            print(f"{p.pid}: {p.name} - {p.window_title}")

        # 特定プロセスの音声をキャプチャ
        with ProcessAudioCapture(pid=1234, output_path="output.wav") as capture:
            capture.start()
            time.sleep(10)  # 10秒録音
            capture.stop()
    """

    _dll: Optional[ctypes.WinDLL] = None
    _dll_path: Optional[Path] = None

    @classmethod
    def _load_dll(cls, dll_path: Optional[Path] = None) -> ctypes.WinDLL:
        """DLLをロード"""
        if cls._dll is not None and (dll_path is None or dll_path == cls._dll_path):
            return cls._dll

        if dll_path is None:
            # デフォルトパス（このファイルと同じディレクトリ or build/bin）
            search_paths = [
                Path(__file__).parent / "ProcessAudioCapture.dll",
                Path(__file__).parent.parent / "build" / "bin" / "Release" / "ProcessAudioCapture.dll",
                Path(__file__).parent.parent / "build" / "bin" / "Debug" / "ProcessAudioCapture.dll",
                Path(__file__).parent.parent / "build" / "Release" / "ProcessAudioCapture.dll",
                Path(__file__).parent.parent / "build" / "Debug" / "ProcessAudioCapture.dll",
            ]

            for path in search_paths:
                if path.exists():
                    dll_path = path
                    break
            else:
                raise FileNotFoundError(
                    f"ProcessAudioCapture.dll not found. Searched: {search_paths}"
                )

        cls._dll = ctypes.WinDLL(str(dll_path))
        cls._dll_path = dll_path

        # 関数シグネチャを設定
        cls._setup_function_signatures(cls._dll)

        return cls._dll

    @classmethod
    def _setup_function_signatures(cls, dll: ctypes.WinDLL):
        """DLL関数のシグネチャを設定"""
        # PacIsSupported
        dll.PacIsSupported.argtypes = []
        dll.PacIsSupported.restype = wintypes.BOOL

        # PacEnumerateAudioProcesses
        dll.PacEnumerateAudioProcesses.argtypes = [
            ctypes.POINTER(PacProcessInfo),
            ctypes.c_int,
            ctypes.POINTER(ctypes.c_int)
        ]
        dll.PacEnumerateAudioProcesses.restype = ctypes.c_int

        # PacStartCapture
        dll.PacStartCapture.argtypes = [
            wintypes.DWORD,
            ctypes.c_int,
            ctypes.c_wchar_p,
            LevelCallbackFunc,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p)
        ]
        dll.PacStartCapture.restype = ctypes.c_int

        # PacStopCapture
        dll.PacStopCapture.argtypes = [ctypes.c_void_p]
        dll.PacStopCapture.restype = ctypes.c_int

        # PacGetLevel
        dll.PacGetLevel.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_float)]
        dll.PacGetLevel.restype = ctypes.c_int

        # PacIsCapturing
        dll.PacIsCapturing.argtypes = [ctypes.c_void_p]
        dll.PacIsCapturing.restype = wintypes.BOOL

        # PacGetLastErrorMessage
        dll.PacGetLastErrorMessage.argtypes = [ctypes.c_wchar_p, ctypes.c_int]
        dll.PacGetLastErrorMessage.restype = None

        # PacGetVersion
        dll.PacGetVersion.argtypes = []
        dll.PacGetVersion.restype = ctypes.c_char_p

        # PacPauseCapture
        dll.PacPauseCapture.argtypes = [ctypes.c_void_p]
        dll.PacPauseCapture.restype = ctypes.c_int

        # PacResumeCapture
        dll.PacResumeCapture.argtypes = [ctypes.c_void_p]
        dll.PacResumeCapture.restype = ctypes.c_int

        # PacIsPaused
        dll.PacIsPaused.argtypes = [ctypes.c_void_p]
        dll.PacIsPaused.restype = wintypes.BOOL

    @classmethod
    def is_supported(cls, dll_path: Optional[Path] = None) -> bool:
        """
        Windows バージョンがサポートされているか確認

        Returns:
            bool: Windows 10 version 2004以降ならTrue
        """
        try:
            dll = cls._load_dll(dll_path)
            return bool(dll.PacIsSupported())
        except FileNotFoundError:
            return False

    @classmethod
    def get_version(cls, dll_path: Optional[Path] = None) -> str:
        """DLLバージョンを取得"""
        dll = cls._load_dll(dll_path)
        return dll.PacGetVersion().decode('utf-8')

    @classmethod
    def enumerate_audio_processes(cls, dll_path: Optional[Path] = None) -> List[AudioProcess]:
        """
        音声を出力しているプロセス一覧を取得

        Returns:
            List[AudioProcess]: プロセス情報のリスト
        """
        dll = cls._load_dll(dll_path)

        max_count = 64
        processes = (PacProcessInfo * max_count)()
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
        """最後のエラーメッセージを取得"""
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
        Args:
            pid: 対象プロセスID
            output_path: 出力WAVファイルパス（Noneの場合はモニターのみ）
            mode: キャプチャモード（INCLUDE または EXCLUDE）
            level_callback: 音量レベルコールバック関数（引数: dB値）
            dll_path: DLLのパス（省略時は自動検索）
        """
        self._dll = self._load_dll(dll_path)
        self._pid = pid
        self._output_path = output_path
        self._mode = mode
        self._user_callback = level_callback
        self._handle: Optional[ctypes.c_void_p] = None
        self._callback_ref: Optional[LevelCallbackFunc] = None
        self._lock = threading.Lock()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
        return False

    def start(self):
        """キャプチャを開始"""
        with self._lock:
            if self._handle is not None:
                raise ProcessAudioCaptureError(
                    PacErrorCode.ALREADY_RECORDING,
                    "Already capturing"
                )

            # コールバック関数をラップ
            if self._user_callback:
                def _callback(level_db: float, user_data: ctypes.c_void_p):
                    if self._user_callback:
                        self._user_callback(level_db)

                self._callback_ref = LevelCallbackFunc(_callback)
            else:
                self._callback_ref = LevelCallbackFunc(0)  # NULL

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
        """キャプチャを停止"""
        with self._lock:
            if self._handle is None:
                return

            err = self._dll.PacStopCapture(self._handle)
            self._handle = None
            self._callback_ref = None

            if err != PacErrorCode.SUCCESS:
                raise ProcessAudioCaptureError(err, self._get_last_error())

    @property
    def is_capturing(self) -> bool:
        """キャプチャ中かどうか"""
        with self._lock:
            if self._handle is None:
                return False
            return bool(self._dll.PacIsCapturing(self._handle))

    @property
    def level_db(self) -> float:
        """現在の音量レベル（dB）"""
        with self._lock:
            if self._handle is None:
                return -60.0

            level = ctypes.c_float()
            err = self._dll.PacGetLevel(self._handle, ctypes.byref(level))

            if err != PacErrorCode.SUCCESS:
                return -60.0

            return level.value

    def pause(self):
        """Pause capturing (audio data not written to file)"""
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
        """Resume capturing"""
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
    def is_paused(self) -> bool:
        """Check if capture is paused"""
        with self._lock:
            if self._handle is None:
                return False
            return bool(self._dll.PacIsPaused(self._handle))


def main():
    """テスト用メイン関数"""
    import time

    print(f"ProcessAudioCapture Python Wrapper")
    print(f"=" * 40)

    # サポートチェック
    if not ProcessAudioCapture.is_supported():
        print("ERROR: Windows 10 version 2004 or later is required.")
        return

    print(f"Version: {ProcessAudioCapture.get_version()}")
    print()

    # プロセス一覧を取得
    print("Audio processes:")
    print("-" * 40)

    processes = ProcessAudioCapture.enumerate_audio_processes()

    if not processes:
        print("No audio processes found.")
        return

    for i, p in enumerate(processes, 1):
        print(f"  [{i}] PID: {p.pid:5d}  Name: {p.name[:30]:<30}  Window: {p.window_title[:30] if p.window_title else '(no window)'}")

    print()
    print(f"Total: {len(processes)} processes")

    # ユーザーに選択させる
    print()
    choice = input("Select process number to capture (or Enter to skip): ").strip()

    if not choice:
        print("Skipped.")
        return

    try:
        idx = int(choice) - 1
        if idx < 0 or idx >= len(processes):
            print("Invalid selection.")
            return

        selected = processes[idx]
        print(f"\nCapturing from: {selected.name} (PID: {selected.pid})")
        print("Press Ctrl+C to stop.\n")

        def on_level(level_db: float):
            bar_len = int((level_db + 60) / 60 * 50)
            bar_len = max(0, min(50, bar_len))
            bar = "=" * bar_len + " " * (50 - bar_len)
            print(f"\r[{bar}] {level_db:6.1f} dB  ", end="", flush=True)

        with ProcessAudioCapture(
            pid=selected.pid,
            output_path="test_output.wav",
            level_callback=on_level
        ) as capture:
            capture.start()

            try:
                while True:
                    time.sleep(0.1)
            except KeyboardInterrupt:
                print("\n\nStopping...")

        print("Done. Output saved to test_output.wav")

    except ValueError:
        print("Invalid input.")
    except ProcessAudioCaptureError as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()

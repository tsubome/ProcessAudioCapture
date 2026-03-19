"""ProcessAudioCapture macOS backend — ScreenCaptureKit"""
import sys
import os
import struct
import threading
import wave
import math
import time
from dataclasses import dataclass
from typing import List, Optional, Callable

# ScreenCaptureKit imports
try:
    from ScreenCaptureKit import (
        SCShareableContent,
        SCContentFilter,
        SCStreamConfiguration,
        SCStream,
    )
    from CoreMedia import CMSampleBufferGetDataBuffer, CMBlockBufferGetDataLength, CMBlockBufferCopyDataBytes
    from Foundation import NSObject, NSRunLoop, NSDate
    import objc
    _SCK_AVAILABLE = True
except ImportError:
    _SCK_AVAILABLE = False


@dataclass
class AudioProcess:
    pid: int
    name: str
    window_title: str


class ProcessAudioCaptureError(Exception):
    def __init__(self, message="", code=0):
        self.code = code
        self.message = message
        super().__init__(message)


class PacCaptureMode:
    INCLUDE = 0
    EXCLUDE = 1


class PacErrorCode:
    SUCCESS = 0
    INVALID_ARGUMENT = -1
    NOT_INITIALIZED = -2
    ALREADY_RECORDING = -3
    NOT_RECORDING = -4
    COM_ERROR = -5
    DEVICE_ERROR = -6
    NOT_SUPPORTED = -7


class ProcessAudioCapture:
    INCLUDE = PacCaptureMode.INCLUDE
    EXCLUDE = PacCaptureMode.EXCLUDE

    @staticmethod
    def is_supported() -> bool:
        """macOS 12.3+ で ScreenCaptureKit が利用可能か"""
        return _SCK_AVAILABLE

    @staticmethod
    def get_version() -> str:
        """バージョン文字列を返す"""
        return "1.0.0-macos (ScreenCaptureKit)"

    @staticmethod
    def _get_audio_output_pids() -> set:
        """Core Audio HAL で音声出力中のプロセスPIDを取得"""
        try:
            import CoreAudio

            addr = CoreAudio.AudioObjectPropertyAddress(
                mSelector=CoreAudio.kAudioHardwarePropertyProcessObjectList,
                mScope=CoreAudio.kAudioObjectPropertyScopeGlobal,
                mElement=0
            )
            err, size = CoreAudio.AudioObjectGetPropertyDataSize(
                CoreAudio.kAudioObjectSystemObject, addr, 0, [], None
            )
            if err != 0 or size <= 0:
                return set()

            err2, _, raw_data = CoreAudio.AudioObjectGetPropertyData(
                CoreAudio.kAudioObjectSystemObject, addr, 0, [], size, None
            )
            if err2 != 0 or not raw_data:
                return set()

            obj_ids = struct.unpack(f'<{len(raw_data)//4}I', raw_data)
            obj_ids = [x for x in obj_ids if x != 0]

            audio_pids = set()
            for obj_id in obj_ids:
                try:
                    # IsRunningOutput チェック
                    out_addr = CoreAudio.AudioObjectPropertyAddress(
                        mSelector=CoreAudio.kAudioProcessPropertyIsRunningOutput,
                        mScope=CoreAudio.kAudioObjectPropertyScopeGlobal, mElement=0)
                    _, _, out_raw = CoreAudio.AudioObjectGetPropertyData(
                        obj_id, out_addr, 0, [], 4, None)
                    is_output = struct.unpack('<I', out_raw)[0]
                    if not is_output:
                        continue

                    # PID 取得
                    pid_addr = CoreAudio.AudioObjectPropertyAddress(
                        mSelector=CoreAudio.kAudioProcessPropertyPID,
                        mScope=CoreAudio.kAudioObjectPropertyScopeGlobal, mElement=0)
                    _, _, pid_raw = CoreAudio.AudioObjectGetPropertyData(
                        obj_id, pid_addr, 0, [], 4, None)
                    pid = struct.unpack('<I', pid_raw)[0]
                    if pid > 0:
                        audio_pids.add(pid)
                except Exception:
                    continue
            return audio_pids
        except Exception:
            return set()

    @staticmethod
    def enumerate_audio_processes() -> List[AudioProcess]:
        """音声を出力中かつキャプチャ可能なアプリケーション一覧を取得

        Core Audio HAL で「音声出力中」のPIDを取得し、
        ScreenCaptureKit で「キャプチャ可能」なアプリと交差させる。
        → DRM 保護アプリ（Apple Music等）は ScreenCaptureKit に現れないため除外される。
        → 音声を出していないアプリは Core Audio HAL に現れないため除外される。
        """
        if not _SCK_AVAILABLE:
            return []

        # Step 1: Core Audio HAL で音声出力中の PID を取得
        audio_pids = ProcessAudioCapture._get_audio_output_pids()

        result = []
        event = threading.Event()

        def handler(content, error):
            if content is not None:
                apps = content.applications()
                windows = content.windows()

                # PID → ウィンドウタイトルのマッピング
                pid_windows = {}
                for w in windows:
                    owner = w.owningApplication()
                    if owner:
                        wpid = owner.processID()
                        if wpid not in pid_windows:
                            pid_windows[wpid] = str(w.title() or "")

                for app in apps:
                    pid = app.processID()
                    name = str(app.applicationName() or "")

                    # フィルタ: Core Audio で音声出力中のプロセスのみ
                    if pid not in audio_pids:
                        continue

                    if not name:
                        continue

                    title = pid_windows.get(pid, "")
                    result.append(AudioProcess(pid=pid, name=name, window_title=title))
            event.set()

        SCShareableContent.getShareableContentWithCompletionHandler_(handler)
        event.wait(timeout=10)
        return result

    def __init__(self, pid: int, output_path: Optional[str] = None,
                 mode: int = PacCaptureMode.INCLUDE,
                 level_callback: Optional[Callable[[float], None]] = None):
        if not _SCK_AVAILABLE:
            raise ProcessAudioCaptureError("ScreenCaptureKit not available", PacErrorCode.NOT_SUPPORTED)

        self._pid = pid
        self._output_path = output_path
        self._mode = mode
        self._level_callback = level_callback
        self._is_capturing = False
        self._is_paused = False
        self._stream = None
        self._lock = threading.Lock()
        self._audio_buffers: list = []
        self._sample_rate = 48000
        self._channels = 2
        self._delegate = None
        self._current_level_db = -60.0

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
        return False

    @property
    def is_capturing(self) -> bool:
        return self._is_capturing

    @property
    def is_paused(self) -> bool:
        return self._is_paused

    @property
    def level_db(self) -> float:
        """現在の音量レベル (dB)"""
        return self._current_level_db

    def start(self):
        """キャプチャ開始"""
        if self._is_capturing:
            raise ProcessAudioCaptureError("Already recording", PacErrorCode.ALREADY_RECORDING)

        # ターゲットアプリを探す
        event = threading.Event()
        target_app = [None]
        all_content = [None]

        def handler(content, error):
            if content:
                all_content[0] = content
                for app in content.applications():
                    if app.processID() == self._pid:
                        target_app[0] = app
                        break
            event.set()

        SCShareableContent.getShareableContentWithCompletionHandler_(handler)
        event.wait(timeout=10)

        if target_app[0] is None:
            raise ProcessAudioCaptureError(f"Process {self._pid} not found", PacErrorCode.INVALID_ARGUMENT)

        # ディスプレイを取得
        display = None
        if all_content[0] and all_content[0].displays():
            display = all_content[0].displays()[0]

        # SCContentFilter を設定
        if self._mode == PacCaptureMode.INCLUDE:
            # 特定アプリの音声のみキャプチャ
            # アプリベースのフィルタ — 対象アプリのウィンドウを含める
            app_windows = []
            if all_content[0]:
                for w in all_content[0].windows():
                    if w.owningApplication() and w.owningApplication().processID() == self._pid:
                        app_windows.append(w)

            if app_windows and display is not None:
                content_filter = SCContentFilter.alloc().initWithDisplay_includingWindows_(
                    display,
                    app_windows
                )
            else:
                # ウィンドウがない場合、アプリベースでフィルタ
                content_filter = SCContentFilter.alloc().initWithDisplay_includingApplications_exceptingWindows_(
                    display,
                    [target_app[0]],
                    []
                )
        else:
            # EXCLUDE モード: 対象アプリ以外をキャプチャ
            content_filter = SCContentFilter.alloc().initWithDisplay_excludingWindows_(
                display,
                [target_app[0]]
            )

        # SCStreamConfiguration
        config = SCStreamConfiguration.alloc().init()
        config.setCapturesAudio_(True)
        config.setExcludesCurrentProcessAudio_(True)
        config.setSampleRate_(self._sample_rate)
        config.setChannelCount_(self._channels)

        # 映像は不要 — 最小サイズに設定
        config.setWidth_(2)
        config.setHeight_(2)

        # SCStream を作成
        self._stream = SCStream.alloc().initWithFilter_configuration_delegate_(
            content_filter, config, None
        )

        # Stream output delegate
        self._delegate = _StreamOutputDelegate.alloc().init()
        self._delegate._capture = self

        error_holder = objc.nil
        success = self._stream.addStreamOutput_type_sampleHandlerQueue_error_(
            self._delegate,
            1,  # SCStreamOutputTypeAudio
            None,  # use default queue
            error_holder,
        )

        # 開始
        start_event = threading.Event()
        start_error = [None]

        def start_handler(error):
            if error:
                start_error[0] = str(error)
            start_event.set()

        self._stream.startCaptureWithCompletionHandler_(start_handler)
        start_event.wait(timeout=10)

        if start_error[0]:
            raise ProcessAudioCaptureError(f"Failed to start: {start_error[0]}", PacErrorCode.DEVICE_ERROR)

        self._is_capturing = True

    def stop(self):
        """キャプチャ停止"""
        if not self._is_capturing:
            return

        stop_event = threading.Event()

        def stop_handler(error):
            stop_event.set()

        if self._stream:
            self._stream.stopCaptureWithCompletionHandler_(stop_handler)
            stop_event.wait(timeout=10)

        self._is_capturing = False
        self._is_paused = False

        # WAV ファイルに書き出し
        if self._output_path and self._audio_buffers:
            self._write_wav()

    def pause(self):
        """一時停止"""
        self._is_paused = True

    def resume(self):
        """再開"""
        self._is_paused = False

    def _on_audio_buffer(self, data: bytes, num_samples: int = 0):
        """オーディオバッファ受信時のコールバック

        ScreenCaptureKit は非インターリーブ（プレーナー）float32 で出力する:
          [L0 L1 L2 ... Ln] [R0 R1 R2 ... Rn]
        WAV はインターリーブ形式が必要:
          [L0 R0 L1 R1 L2 R2 ... Ln Rn]
        ここでプレーナー → インターリーブ変換を行う。
        """
        if self._is_paused:
            return

        # プレーナー → インターリーブ変換
        if self._channels == 2 and num_samples > 0:
            try:
                floats = struct.unpack(f'<{len(data)//4}f', data)
                half = num_samples
                if len(floats) >= half * 2:
                    left = floats[:half]
                    right = floats[half:half*2]
                    # インターリーブ: L0 R0 L1 R1 ...
                    interleaved = []
                    for l, r in zip(left, right):
                        interleaved.append(l)
                        interleaved.append(r)
                    data = struct.pack(f'<{len(interleaved)}f', *interleaved)
            except Exception:
                pass

        self._audio_buffers.append(data)

        # レベル計算
        if self._level_callback and len(data) >= 4:
            try:
                samples = struct.unpack(f'<{len(data)//4}f', data)
                if samples:
                    rms = math.sqrt(sum(s*s for s in samples) / len(samples))
                    db = 20 * math.log10(max(rms, 1e-10))
                    self._current_level_db = max(db, -60.0)
                    self._level_callback(self._current_level_db)
            except Exception:
                pass

    def _write_wav(self):
        """蓄積したバッファをWAVファイルに書き出し"""
        try:
            all_data = b''.join(self._audio_buffers)
            # float32 → int16 変換
            float_samples = struct.unpack(f'<{len(all_data)//4}f', all_data)
            int16_data = struct.pack(
                f'<{len(float_samples)}h',
                *(max(-32768, min(32767, int(s * 32767))) for s in float_samples)
            )

            with wave.open(self._output_path, 'wb') as wf:
                wf.setnchannels(self._channels)
                wf.setsampwidth(2)  # 16-bit
                wf.setframerate(self._sample_rate)
                wf.writeframes(int16_data)
        except Exception:
            pass


# SystemAudioCapture — システム全体の音声をキャプチャ（WASAPI Loopback の Mac 代替）
class SystemAudioCapture:
    """macOS: ScreenCaptureKit でシステム音声全体をキャプチャ"""

    def __init__(self, output_path: Optional[str] = None,
                 level_callback: Optional[Callable[[float], None]] = None):
        if not _SCK_AVAILABLE:
            raise ProcessAudioCaptureError("ScreenCaptureKit not available", PacErrorCode.NOT_SUPPORTED)
        self._output_path = output_path
        self._level_callback = level_callback
        self._is_capturing = False
        self._is_paused = False
        self._stream = None
        self._audio_buffers: list = []
        self._sample_rate = 48000
        self._channels = 2
        self._delegate = None
        self._current_level_db = -60.0

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
        return False

    @staticmethod
    def is_supported() -> bool:
        return _SCK_AVAILABLE

    @property
    def is_capturing(self) -> bool:
        return self._is_capturing

    @property
    def is_paused(self) -> bool:
        return self._is_paused

    @property
    def level_db(self) -> float:
        return self._current_level_db

    def start(self):
        """システム全体の音声キャプチャ開始"""
        if self._is_capturing:
            raise ProcessAudioCaptureError("Already recording", PacErrorCode.ALREADY_RECORDING)

        event = threading.Event()
        all_content = [None]

        def handler(content, error):
            all_content[0] = content
            event.set()

        SCShareableContent.getShareableContentWithCompletionHandler_(handler)
        event.wait(timeout=10)

        if not all_content[0] or not all_content[0].displays():
            raise ProcessAudioCaptureError("No display found", PacErrorCode.DEVICE_ERROR)

        display = all_content[0].displays()[0]

        # 全アプリを含むフィルタ（何も除外しない = 全部キャプチャ）
        content_filter = SCContentFilter.alloc().initWithDisplay_excludingWindows_(
            display, []
        )

        config = SCStreamConfiguration.alloc().init()
        config.setCapturesAudio_(True)
        config.setExcludesCurrentProcessAudio_(True)
        config.setSampleRate_(self._sample_rate)
        config.setChannelCount_(self._channels)
        config.setWidth_(2)
        config.setHeight_(2)

        self._stream = SCStream.alloc().initWithFilter_configuration_delegate_(
            content_filter, config, None
        )

        self._delegate = _StreamOutputDelegate.alloc().init()
        self._delegate._capture = self

        error_holder = objc.nil
        self._stream.addStreamOutput_type_sampleHandlerQueue_error_(
            self._delegate, 1, None, error_holder
        )

        start_event = threading.Event()
        start_error = [None]

        def start_handler(error):
            if error:
                start_error[0] = str(error)
            start_event.set()

        self._stream.startCaptureWithCompletionHandler_(start_handler)
        start_event.wait(timeout=10)

        if start_error[0]:
            raise ProcessAudioCaptureError(f"Failed to start: {start_error[0]}", PacErrorCode.DEVICE_ERROR)

        self._is_capturing = True

    def stop(self):
        if not self._is_capturing:
            return
        stop_event = threading.Event()

        def stop_handler(error):
            stop_event.set()

        if self._stream:
            self._stream.stopCaptureWithCompletionHandler_(stop_handler)
            stop_event.wait(timeout=10)
        self._is_capturing = False
        self._is_paused = False
        if self._output_path and self._audio_buffers:
            self._write_wav()

    def pause(self):
        self._is_paused = True

    def resume(self):
        self._is_paused = False

    def _on_audio_buffer(self, data: bytes, num_samples: int = 0):
        """ProcessAudioCapture と同じバッファ処理（プレーナー→インターリーブ変換含む）"""
        if self._is_paused:
            return

        # プレーナー → インターリーブ変換
        if self._channels == 2 and num_samples > 0:
            try:
                floats = struct.unpack(f'<{len(data)//4}f', data)
                half = num_samples
                if len(floats) >= half * 2:
                    left = floats[:half]
                    right = floats[half:half*2]
                    interleaved = []
                    for l, r in zip(left, right):
                        interleaved.append(l)
                        interleaved.append(r)
                    data = struct.pack(f'<{len(interleaved)}f', *interleaved)
            except Exception:
                pass

        self._audio_buffers.append(data)
        if self._level_callback and len(data) >= 4:
            try:
                samples = struct.unpack(f'<{len(data)//4}f', data)
                if samples:
                    rms = math.sqrt(sum(s*s for s in samples) / len(samples))
                    db = 20 * math.log10(max(rms, 1e-10))
                    self._current_level_db = max(db, -60.0)
                    self._level_callback(self._current_level_db)
            except Exception:
                pass

    def _write_wav(self):
        try:
            all_data = b''.join(self._audio_buffers)
            float_samples = struct.unpack(f'<{len(all_data)//4}f', all_data)
            int16_data = struct.pack(
                f'<{len(float_samples)}h',
                *(max(-32768, min(32767, int(s * 32767))) for s in float_samples)
            )
            with wave.open(self._output_path, 'wb') as wf:
                wf.setnchannels(self._channels)
                wf.setsampwidth(2)
                wf.setframerate(self._sample_rate)
                wf.writeframes(int16_data)
        except Exception:
            pass


# --- Objective-C Delegate for SCStream output ---

if _SCK_AVAILABLE:
    class _StreamOutputDelegate(NSObject):
        """SCStreamOutput プロトコル準拠デリゲート"""

        def stream_didOutputSampleBuffer_ofType_(self, stream, sample_buffer, output_type):
            """音声サンプルバッファ受信"""
            if output_type != 1:  # SCStreamOutputTypeAudio
                return

            capture = getattr(self, '_capture', None)
            if capture is None:
                return

            try:
                # CMSampleBuffer から実際のオーディオフォーマットを推定（初回のみ）
                if not getattr(capture, '_format_detected', False):
                    try:
                        from CoreMedia import CMSampleBufferGetNumSamples
                        num_samples = CMSampleBufferGetNumSamples(sample_buffer)
                        block_buf = CMSampleBufferGetDataBuffer(sample_buffer)
                        if block_buf is not None and num_samples > 0:
                            data_len = CMBlockBufferGetDataLength(block_buf)
                            bytes_per_sample = data_len // num_samples
                            # float32 stereo = 8, float32 mono = 4
                            if bytes_per_sample == 8:
                                capture._channels = 2
                            elif bytes_per_sample == 4:
                                capture._channels = 1
                    except Exception:
                        pass
                    # 2つ目のバッファでサンプルレートを推定
                    if not getattr(capture, '_first_timestamp', None):
                        try:
                            from CoreMedia import CMSampleBufferGetPresentationTimeStamp
                            ts = CMSampleBufferGetPresentationTimeStamp(sample_buffer)
                            from CoreMedia import CMSampleBufferGetNumSamples
                            ns = CMSampleBufferGetNumSamples(sample_buffer)
                            capture._first_timestamp = (ts.value / ts.timescale if ts.timescale else 0, ns)
                        except Exception:
                            capture._format_detected = True
                    else:
                        try:
                            from CoreMedia import CMSampleBufferGetPresentationTimeStamp, CMSampleBufferGetNumSamples
                            ts = CMSampleBufferGetPresentationTimeStamp(sample_buffer)
                            t = ts.value / ts.timescale if ts.timescale else 0
                            prev_t, prev_ns = capture._first_timestamp
                            dt = t - prev_t
                            if dt > 0:
                                estimated_rate = int(round(prev_ns / dt))
                                # 一般的なサンプルレートに丸める
                                for standard in [44100, 48000, 96000, 22050, 16000]:
                                    if abs(estimated_rate - standard) < standard * 0.05:
                                        capture._sample_rate = standard
                                        break
                        except Exception:
                            pass
                        capture._format_detected = True

                # CMSampleBuffer からオーディオデータを取得
                block_buffer = CMSampleBufferGetDataBuffer(sample_buffer)
                if block_buffer is None:
                    return

                length = CMBlockBufferGetDataLength(block_buffer)
                if length <= 0:
                    return

                # サンプル数を取得（プレーナー→インターリーブ変換に必要）
                from CoreMedia import CMSampleBufferGetNumSamples
                num_samples = CMSampleBufferGetNumSamples(sample_buffer)

                data = bytearray(length)
                CMBlockBufferCopyDataBytes(block_buffer, 0, length, data)

                capture._on_audio_buffer(bytes(data), num_samples=num_samples)
            except Exception:
                pass

else:
    # _SCK_AVAILABLE が False の場合のスタブ（ImportError 時に定義だけ通す）
    class _StreamOutputDelegate:  # type: ignore[no-redef]
        pass


__all__ = [
    "ProcessAudioCapture",
    "SystemAudioCapture",
    "AudioProcess",
    "ProcessAudioCaptureError",
    "PacCaptureMode",
    "PacErrorCode",
]

#!/usr/bin/env python3
"""
ProcessAudioCapture CLI — Cross-platform test tool
Windows: WASAPI Process Loopback / macOS: ScreenCaptureKit
"""
import sys
import os
import time

# src_python をパスに追加
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src_python"))

from process_audio_capture import (
    ProcessAudioCapture,
    AudioProcess,
    ProcessAudioCaptureError,
)


def main():
    platform = "macOS" if sys.platform == "darwin" else "Windows"
    print(f"ProcessAudioCapture CLI ({platform})")
    print("=" * 50)

    # サポートチェック
    if not ProcessAudioCapture.is_supported():
        if sys.platform == "darwin":
            print("ERROR: ScreenCaptureKit not available.")
            print("  macOS 12.3+ required. Install: pip install pyobjc-framework-ScreenCaptureKit")
        else:
            print("ERROR: Windows 10 version 2004 or later required.")
        return

    if hasattr(ProcessAudioCapture, 'get_version'):
        print(f"Version: {ProcessAudioCapture.get_version()}")
    print()

    # プロセス一覧を取得
    print("Enumerating audio processes...")
    print("-" * 50)

    processes = ProcessAudioCapture.enumerate_audio_processes()

    if not processes:
        print("No audio processes found.")
        print("  Make sure some app is playing audio (YouTube, Spotify, etc.)")
        return

    for i, p in enumerate(processes, 1):
        title = p.window_title[:30] if p.window_title else "(no window)"
        print(f"  [{i:2d}] PID: {p.pid:5d}  {p.name[:25]:<25}  {title}")

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
        output_file = "test_output.wav"
        print(f"\nCapturing from: {selected.name} (PID: {selected.pid})")
        print(f"Output: {output_file}")
        print("Press Ctrl+C to stop.\n")

        def on_level(level_db: float):
            bar_len = int((level_db + 60) / 60 * 50)
            bar_len = max(0, min(50, bar_len))
            bar = "=" * bar_len + " " * (50 - bar_len)
            print(f"\r[{bar}] {level_db:6.1f} dB  ", end="", flush=True)

        capture = ProcessAudioCapture(
            pid=selected.pid,
            output_path=output_file,
            level_callback=on_level,
        )
        capture.start()

        try:
            while capture.is_capturing:
                time.sleep(0.1)
        except KeyboardInterrupt:
            print("\n\nStopping...")

        capture.stop()

        if os.path.exists(output_file):
            size = os.path.getsize(output_file)
            print(f"Done. Output saved to {output_file} ({size:,} bytes)")
        else:
            print("Done. (no output file created)")

    except ValueError:
        print("Invalid input.")
    except ProcessAudioCaptureError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")


if __name__ == "__main__":
    main()

// test_capture.cpp
// ProcessAudioCapture DLL test application

#include <Windows.h>
#include <objbase.h>
#include <stdio.h>
#include <conio.h>
#include "../src/ProcessAudioCapture.h"

#pragma comment(lib, "ole32.lib")

// Audio level callback
void LevelCallback(float levelDb, void* userData)
{
    // Display volume bar (-60dB to 0dB mapped to 0-50 characters)
    int barLength = (int)((levelDb + 60.0f) / 60.0f * 50.0f);
    if (barLength < 0) barLength = 0;
    if (barLength > 50) barLength = 50;

    printf("\r[");
    for (int i = 0; i < 50; i++)
    {
        if (i < barLength)
            printf("=");
        else
            printf(" ");
    }
    printf("] %6.1f dB  ", levelDb);
    fflush(stdout);
}

void PrintUsage()
{
    printf("Usage:\n");
    printf("  test_capture.exe list           - List audio processes\n");
    printf("  test_capture.exe <pid> [output] - Capture audio from process\n");
    printf("\nExamples:\n");
    printf("  test_capture.exe list\n");
    printf("  test_capture.exe 1234\n");
    printf("  test_capture.exe 1234 output.wav\n");
}

int wmain(int argc, wchar_t* argv[])
{
    // Initialize COM as MTA (required for process loopback capture)
    HRESULT hrCom = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hrCom) && hrCom != RPC_E_CHANGED_MODE)
    {
        printf("ERROR: Failed to initialize COM: 0x%08X\n", hrCom);
        return 1;
    }

    printf("ProcessAudioCapture Test v%s\n", PacGetVersion());
    printf("================================\n\n");

    // Support check
    if (!PacIsSupported())
    {
        printf("ERROR: Windows 10 version 2004 or later is required.\n");
        if (SUCCEEDED(hrCom)) CoUninitialize();
        return 1;
    }
    printf("Windows version: Supported\n\n");

    if (argc < 2)
    {
        PrintUsage();
        if (SUCCEEDED(hrCom)) CoUninitialize();
        return 1;
    }

    // "list" command
    if (wcscmp(argv[1], L"list") == 0)
    {
        printf("Audio processes:\n");
        printf("----------------\n");

        PacProcessInfo processes[64];
        int count = 0;
        PacErrorCode err = PacEnumerateAudioProcesses(processes, 64, &count);

        if (err != PAC_SUCCESS)
        {
            wchar_t errMsg[512];
            PacGetLastErrorMessage(errMsg, 512);
            wprintf(L"Error: %s\n", errMsg);
            if (SUCCEEDED(hrCom)) CoUninitialize();
            return 1;
        }

        if (count == 0)
        {
            printf("No audio processes found.\n");
            if (SUCCEEDED(hrCom)) CoUninitialize();
            return 0;
        }

        for (int i = 0; i < count; i++)
        {
            wprintf(L"  [%d] PID: %5d  Name: %-30s  Window: %s\n",
                i + 1,
                processes[i].processId,
                processes[i].processName,
                processes[i].windowTitle[0] ? processes[i].windowTitle : L"(no window)"
            );
        }

        printf("\nTotal: %d processes\n", count);
        if (SUCCEEDED(hrCom)) CoUninitialize();
        return 0;
    }

    // Capture command
    DWORD pid = (DWORD)_wtoi(argv[1]);
    if (pid == 0)
    {
        printf("ERROR: Invalid process ID.\n");
        PrintUsage();
        if (SUCCEEDED(hrCom)) CoUninitialize();
        return 1;
    }

    const wchar_t* outputPath = nullptr;
    if (argc >= 3)
    {
        outputPath = argv[2];
    }

    printf("Target process ID: %d\n", pid);
    if (outputPath)
    {
        wprintf(L"Output file: %s\n", outputPath);
    }
    else
    {
        printf("Output file: (none - monitor only)\n");
    }
    printf("\n");

    // Start capture
    PacHandle handle = nullptr;
    PacErrorCode err = PacStartCapture(
        pid,
        PAC_MODE_INCLUDE,
        outputPath,
        LevelCallback,
        nullptr,
        &handle
    );

    if (err != PAC_SUCCESS)
    {
        wchar_t errMsg[512];
        PacGetLastErrorMessage(errMsg, 512);
        wprintf(L"Failed to start capture: %s (code: %d)\n", errMsg, err);
        if (SUCCEEDED(hrCom)) CoUninitialize();
        return 1;
    }

    printf("Capturing... Press any key to stop.\n\n");

    // Wait for key input
    while (!_kbhit())
    {
        Sleep(100);
    }
    _getch();

    printf("\n\nStopping capture...\n");

    // Stop capture
    err = PacStopCapture(handle);
    if (err != PAC_SUCCESS)
    {
        wchar_t errMsg[512];
        PacGetLastErrorMessage(errMsg, 512);
        wprintf(L"Failed to stop capture: %s\n", errMsg);
        if (SUCCEEDED(hrCom)) CoUninitialize();
        return 1;
    }

    printf("Done.\n");
    if (SUCCEEDED(hrCom)) CoUninitialize();
    return 0;
}

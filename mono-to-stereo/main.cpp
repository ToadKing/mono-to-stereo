// main.cpp

#include "common.h"

int do_everything(int argc, LPCWSTR argv[]);

int _cdecl wmain(int argc, LPCWSTR argv[]) {
    HRESULT hr = S_OK;

    hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        ERR(L"CoInitialize failed: hr = 0x%08x", hr);
        return -__LINE__;
    }
    CoUninitializeOnExit cuoe;

    return do_everything(argc, argv);
}

int do_everything(int argc, LPCWSTR argv[]) {
    HRESULT hr = S_OK;

    // parse command line
    CPrefs prefs(argc, argv, hr);
    if (FAILED(hr)) {
        ERR(L"CPrefs::CPrefs constructor failed: hr = 0x%08x", hr);
        return -__LINE__;
    }
    if (S_FALSE == hr) {
        // nothing to do
        return 0;
    }

    // create a "loopback capture has started" event
    HANDLE hStartedEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (NULL == hStartedEvent) {
        ERR(L"CreateEvent failed: last error is %u", GetLastError());
        return -__LINE__;
    }
    CloseHandleOnExit closeStartedEvent(hStartedEvent);

    // create a "stop capturing now" event
    HANDLE hStopEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (NULL == hStopEvent) {
        ERR(L"CreateEvent failed: last error is %u", GetLastError());
        return -__LINE__;
    }
    CloseHandleOnExit closeStopEvent(hStopEvent);

    // create arguments for loopback capture thread
    LoopbackCaptureThreadFunctionArguments threadArgs;
    threadArgs.hr = E_UNEXPECTED; // thread will overwrite this
    threadArgs.pMMInDevice = prefs.m_pMMInDevice;
    threadArgs.pMMOutDevice = prefs.m_pMMOutDevice;
    threadArgs.iBufferMs = prefs.m_iBufferMs;
    threadArgs.bSkipFirstSample = prefs.m_bSkipFirstSample;
    threadArgs.bForceMonoToStereo = prefs.m_bForceMonoToStereo;
    threadArgs.bCaptureRenderer = prefs.m_bCaptureRenderer;
    threadArgs.bDuplicateChannels = prefs.m_bDuplicateChannels;
    threadArgs.preProcess = prefs.m_preProcess;
    threadArgs.hStartedEvent = hStartedEvent;
    threadArgs.hStopEvent = hStopEvent;
    threadArgs.nFrames = 0;

    LOG(L"<< %ls", prefs.m_szInName);
    LOG(L">> %ls", prefs.m_szOutName);

    HANDLE hThread = CreateThread(
        NULL, 0,
        LoopbackCaptureThreadFunction, &threadArgs,
        0, NULL
    );
    if (NULL == hThread) {
        ERR(L"CreateThread failed: last error is %u", GetLastError());
        return -__LINE__;
    }
    CloseHandleOnExit closeThread(hThread);

    // wait for either capture to start or the thread to end
    HANDLE waitArray[2] = { hStartedEvent, hThread };
    DWORD dwWaitResult;
    dwWaitResult = WaitForMultipleObjects(
        ARRAYSIZE(waitArray), waitArray,
        FALSE, INFINITE
    );

    if (WAIT_OBJECT_0 + 1 == dwWaitResult) {
        ERR(L"Thread aborted before starting to capture: hr = 0x%08x", threadArgs.hr);
        return -__LINE__;
    }

    if (WAIT_OBJECT_0 != dwWaitResult) {
        ERR(L"Unexpected WaitForMultipleObjects return value %u", dwWaitResult);
        return -__LINE__;
    }

    // at this point capture is running
    // wait for the user to press a key or for capture to error out
    {
        WaitForSingleObjectOnExit waitForThread(hThread);
        SetEventOnExit setStopEvent(hStopEvent);
        HANDLE hStdIn = GetStdHandle(STD_INPUT_HANDLE);

        if (INVALID_HANDLE_VALUE == hStdIn) {
            ERR(L"GetStdHandle returned INVALID_HANDLE_VALUE: last error is %u", GetLastError());
            return -__LINE__;
        }

        LOG(L"%s", L"Press Enter to quit...\n");

        FlushConsoleInputBuffer(hStdIn);

        HANDLE rhHandles[2] = { hThread, hStdIn };

        bool bKeepWaiting = true;
        while (bKeepWaiting) {

            dwWaitResult = WaitForMultipleObjects(2, rhHandles, FALSE, INFINITE);

            switch (dwWaitResult) {

            case WAIT_OBJECT_0: // hThread
                ERR(L"%s", L"The thread terminated early - something bad happened");
                bKeepWaiting = false;
                break;

            case WAIT_OBJECT_0 + 1: // hStdIn
                // see if any of them was an Enter key-up event
                INPUT_RECORD rInput[128];
                DWORD nEvents;
                if (!ReadConsoleInput(hStdIn, rInput, ARRAYSIZE(rInput), &nEvents)) {
                    ERR(L"ReadConsoleInput failed: last error is %u", GetLastError());
                    bKeepWaiting = false;
                }
                else {
                    for (DWORD i = 0; i < nEvents; i++) {
                        if (
                            KEY_EVENT == rInput[i].EventType &&
                            VK_RETURN == rInput[i].Event.KeyEvent.wVirtualKeyCode &&
                            !rInput[i].Event.KeyEvent.bKeyDown
                            ) {
                            LOG(L"%s", L"Stopping capture...");
                            bKeepWaiting = false;
                            break;
                        }
                    }
                    // if none of them were Enter key-up events,
                    // continue waiting
                }
                break;

            default:
                ERR(L"WaitForMultipleObjects returned unexpected value 0x%08x", dwWaitResult);
                bKeepWaiting = false;
                break;
            } // switch
        } // while
    } // naked scope

    // at this point the thread is definitely finished

    DWORD exitCode;
    if (!GetExitCodeThread(hThread, &exitCode)) {
        ERR(L"GetExitCodeThread failed: last error is %u", GetLastError());
        return -__LINE__;
    }

    if (0 != exitCode) {
        ERR(L"Capture thread exit code is %u; expected 0", exitCode);
        return -__LINE__;
    }

    if (S_OK != threadArgs.hr) {
        ERR(L"Thread HRESULT is 0x%08x", threadArgs.hr);
        return -__LINE__;
    }

    // let prefs' destructor call mmioClose

    return 0;
}

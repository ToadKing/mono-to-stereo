// mono-to-stereo.cpp

#include "common.h"

HRESULT LoopbackCapture(
    IMMDevice *pMMInDevice,
    IMMDevice *pMMOutDevice,
    int iBufferMs,
    bool bSwapChannels,
    HANDLE hStartedEvent,
    HANDLE hStopEvent,
    PUINT32 pnFrames
);

DWORD WINAPI LoopbackCaptureThreadFunction(LPVOID pContext) {
    LoopbackCaptureThreadFunctionArguments *pArgs =
        (LoopbackCaptureThreadFunctionArguments*)pContext;

    pArgs->hr = CoInitialize(NULL);
    if (FAILED(pArgs->hr)) {
        ERR(L"CoInitialize failed: hr = 0x%08x", pArgs->hr);
        return 0;
    }
    CoUninitializeOnExit cuoe;

    pArgs->hr = LoopbackCapture(
        pArgs->pMMInDevice,
        pArgs->pMMOutDevice,
        pArgs->iBufferMs,
        pArgs->bSwapChannels,
        pArgs->hStartedEvent,
        pArgs->hStopEvent,
        &pArgs->nFrames
    );

    return 0;
}

void swapMemcpy(void *_dst, void *_src, size_t size, size_t chunkSize) {
    size_t blockSize = chunkSize * 2;

    if (size % blockSize != 0) {
        ERR("bad swapMemcpy size (size %zu, chunkSize %zu)", size, chunkSize);
        return;
    }

    BYTE* dst = (BYTE*)_dst;
    BYTE* src = (BYTE*)_src;

    for (size_t i = 0; i < size / blockSize; i++) {
        memcpy(dst + i * blockSize, src + i * blockSize + chunkSize, chunkSize);
        memcpy(dst + i * blockSize + chunkSize, src + i * blockSize, chunkSize);
    }
}

HRESULT LoopbackCapture(
    IMMDevice *pMMInDevice,
    IMMDevice *pMMOutDevice,
    int iBufferMs,
    bool bSwapChannels,
    HANDLE hStartedEvent,
    HANDLE hStopEvent,
    PUINT32 pnFrames
) {
    HRESULT hr;

    // activate an IAudioClient
    IAudioClient *pAudioClient;
    hr = pMMInDevice->Activate(
        __uuidof(IAudioClient),
        CLSCTX_ALL, NULL,
        (void**)&pAudioClient
    );
    if (FAILED(hr)) {
        ERR(L"IMMDevice::Activate(IAudioClient) failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseAudioClient(pAudioClient);

    // get the default device periodicity
    REFERENCE_TIME hnsDefaultDevicePeriod;
    hr = pAudioClient->GetDevicePeriod(&hnsDefaultDevicePeriod, NULL);
    if (FAILED(hr)) {
        ERR(L"IAudioClient::GetDevicePeriod failed: hr = 0x%08x", hr);
        return hr;
    }

    // get the default device format
    WAVEFORMATEX *pwfx;
    hr = pAudioClient->GetMixFormat(&pwfx);
    if (FAILED(hr)) {
        ERR(L"IAudioClient::GetMixFormat failed: hr = 0x%08x", hr);
        return hr;
    }
    CoTaskMemFreeOnExit freeMixFormat(pwfx);

    if (pwfx->nChannels != 1) {
        ERR(L"device doesn't have 1 channel, has %d", pwfx->nChannels);
        return hr;
    }

    pwfx->nBlockAlign = pwfx->nChannels * pwfx->wBitsPerSample / 8;

    // create a periodic waitable timer
    HANDLE hWakeUp = CreateWaitableTimer(NULL, FALSE, NULL);
    if (NULL == hWakeUp) {
        DWORD dwErr = GetLastError();
        ERR(L"CreateWaitableTimer failed: last error = %u", dwErr);
        return HRESULT_FROM_WIN32(dwErr);
    }
    CloseHandleOnExit closeWakeUp(hWakeUp);

    UINT32 nBlockAlign = pwfx->nBlockAlign;
    *pnFrames = 0;

    // call IAudioClient::Initialize
    // note that AUDCLNT_STREAMFLAGS_LOOPBACK and AUDCLNT_STREAMFLAGS_EVENTCALLBACK
    // do not work together...
    // the "data ready" event never gets set
    // so we're going to do a timer-driven loop
    hr = pAudioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        0,
        0, 0, pwfx, 0
    );
    if (FAILED(hr)) {
        ERR(L"IAudioClient::Initialize failed: hr = 0x%08x", hr);
        return hr;
    }

    // activate an IAudioCaptureClient
    IAudioCaptureClient *pAudioCaptureClient;
    hr = pAudioClient->GetService(
        __uuidof(IAudioCaptureClient),
        (void**)&pAudioCaptureClient
    );
    if (FAILED(hr)) {
        ERR(L"IAudioClient::GetService(IAudioCaptureClient) failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseAudioCaptureClient(pAudioCaptureClient);

    // register with MMCSS
    DWORD nTaskIndex = 0;
    HANDLE hTask = AvSetMmThreadCharacteristics(L"Audio", &nTaskIndex);
    if (NULL == hTask) {
        DWORD dwErr = GetLastError();
        ERR(L"AvSetMmThreadCharacteristics failed: last error = %u", dwErr);
        return HRESULT_FROM_WIN32(dwErr);
    }
    AvRevertMmThreadCharacteristicsOnExit unregisterMmcss(hTask);

    // set the waitable timer
    LARGE_INTEGER liFirstFire;
    liFirstFire.QuadPart = -hnsDefaultDevicePeriod / 2; // negative means relative time
    LONG lTimeBetweenFires = (LONG)hnsDefaultDevicePeriod / 2 / (10 * 1000); // convert to milliseconds
    BOOL bOK = SetWaitableTimer(
        hWakeUp,
        &liFirstFire,
        lTimeBetweenFires,
        NULL, NULL, FALSE
    );
    if (!bOK) {
        DWORD dwErr = GetLastError();
        ERR(L"SetWaitableTimer failed: last error = %u", dwErr);
        return HRESULT_FROM_WIN32(dwErr);
    }
    CancelWaitableTimerOnExit cancelWakeUp(hWakeUp);

    // call IAudioClient::Start
    hr = pAudioClient->Start();
    if (FAILED(hr)) {
        ERR(L"IAudioClient::Start failed: hr = 0x%08x", hr);
        return hr;
    }
    AudioClientStopOnExit stopAudioClient(pAudioClient);

    // update format for stereo conversion
    pwfx->nChannels *= 2;
    pwfx->nSamplesPerSec /= 2;
    pwfx->nBlockAlign *= 2;

    // set up output device
    IAudioClient *pAudioOutClient;
    hr = pMMOutDevice->Activate(
        __uuidof(IAudioClient), CLSCTX_ALL,
        NULL, (void**)&pAudioOutClient);
    if (FAILED(hr)) {
        ERR(L"IMMDevice::Activate(IAudioClient) failed (output): hr = 0x%08x", hr);
        return hr;
    }

    hr = pAudioOutClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        0,
        iBufferMs * 10000,
        iBufferMs * 10000,
        pwfx,
        NULL);
    if (FAILED(hr)) {
        ERR(L"IAudioClient::Initialize failed (output): hr = 0x%08x", hr);
        return hr;
    }

    IAudioRenderClient *pRenderClient;
    hr = pAudioOutClient->GetService(
        __uuidof(IAudioRenderClient),
        (void**)&pRenderClient);

    // Get the actual size of the allocated buffer.
    UINT32 clientBufferFrameCount;
    hr = pAudioOutClient->GetBufferSize(&clientBufferFrameCount);
    if (FAILED(hr)) {
        ERR(L"IAudioClient::GetBufferSize failed (output): hr = 0x%08x", hr);
        return hr;
    }

    // Grab the entire buffer for the initial fill operation.
    BYTE *tmp;
    hr = pRenderClient->GetBuffer(clientBufferFrameCount / 2, &tmp);
    if (FAILED(hr)) {
        ERR(L"IAudioClient::GetBuffer failed (output): hr = 0x%08x", hr);
        return hr;
    }

    hr = pRenderClient->ReleaseBuffer(clientBufferFrameCount / 2, AUDCLNT_BUFFERFLAGS_SILENT);
    if (FAILED(hr)) {
        ERR(L"IAudioCaptureClient::ReleaseBuffer failed (output): hr = 0x%08x", hr);
        return hr;
    }

    hr = pAudioOutClient->Start();
    if (FAILED(hr)) {
        ERR(L"IAudioClient::Start failed (output): hr = 0x%08x", hr);
        return hr;
    }

    SetEvent(hStartedEvent);

    // loopback capture loop
    HANDLE waitArray[2] = { hStopEvent, hWakeUp };
    DWORD dwWaitResult;

    bool bDone = false;
    bool bFirstPacket = true;
    for (UINT32 nPasses = 0; !bDone; nPasses++) {
        // drain data while it is available
        UINT32 nNextPacketSize;
        for (
            hr = pAudioCaptureClient->GetNextPacketSize(&nNextPacketSize);
            SUCCEEDED(hr) && nNextPacketSize > 0;
            hr = pAudioCaptureClient->GetNextPacketSize(&nNextPacketSize)
            ) {
            // get the captured data
            BYTE *pData;
            BYTE *pOutData;
            UINT32 nNumFramesToRead;
            DWORD dwFlags;

            hr = pAudioCaptureClient->GetBuffer(
                &pData,
                &nNumFramesToRead,
                &dwFlags,
                NULL,
                NULL
            );
            if (FAILED(hr)) {
                ERR(L"IAudioCaptureClient::GetBuffer failed on pass %u after %u frames: hr = 0x%08x", nPasses, *pnFrames, hr);
                return hr;
            }

            if (AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY == dwFlags) {
                LOG(L"Probably spurious glitch reported on pass %u after %u frames", nPasses, *pnFrames);
            }
            else if (0 != dwFlags) {
                LOG(L"IAudioCaptureClient::GetBuffer set flags to 0x%08x on pass %u after %u frames", dwFlags, nPasses, *pnFrames);
                return E_UNEXPECTED;
            }

            nNumFramesToRead &= ~1;

            if (0 == nNumFramesToRead) {
                ERR(L"IAudioCaptureClient::GetBuffer said to read 0 frames on pass %u after %u frames", nPasses, *pnFrames);
                return E_UNEXPECTED;
            }

            LONG lBytesToWrite = nNumFramesToRead * nBlockAlign;

            for (;;) {
                hr = pRenderClient->GetBuffer(nNumFramesToRead / 2, &pOutData);
                if (hr == AUDCLNT_E_BUFFER_TOO_LARGE) {
                    ERR(L"%s", L"buffer overflow!");
                    Sleep(1);
                    continue;
                }
                if (FAILED(hr)) {
                    ERR(L"IAudioCaptureClient::GetBuffer failed (output) on pass %u after %u frames: hr = 0x%08x", nPasses, *pnFrames, hr);
                    return hr;
                }
                break;
            }

            if (bSwapChannels) {
                swapMemcpy(pOutData, pData, lBytesToWrite, nBlockAlign);
            }
            else {
                memcpy(pOutData, pData, lBytesToWrite);
            }

            hr = pRenderClient->ReleaseBuffer(nNumFramesToRead / 2, 0);
            if (FAILED(hr)) {
                ERR(L"IAudioCaptureClient::ReleaseBuffer failed (output) on pass %u after %u frames: hr = 0x%08x", nPasses, *pnFrames, hr);
                return hr;
            }

            hr = pAudioCaptureClient->ReleaseBuffer(nNumFramesToRead);
            if (FAILED(hr)) {
                ERR(L"IAudioCaptureClient::ReleaseBuffer failed on pass %u after %u frames: hr = 0x%08x", nPasses, *pnFrames, hr);
                return hr;
            }

            *pnFrames += nNumFramesToRead;

            bFirstPacket = false;
        }

        if (FAILED(hr)) {
            ERR(L"IAudioCaptureClient::GetNextPacketSize failed on pass %u after %u frames: hr = 0x%08x", nPasses, *pnFrames, hr);
            return hr;
        }

        dwWaitResult = WaitForMultipleObjects(
            ARRAYSIZE(waitArray), waitArray,
            FALSE, INFINITE
        );

        if (WAIT_OBJECT_0 == dwWaitResult) {
            LOG(L"Received stop event after %u passes and %u frames", nPasses, *pnFrames);
            bDone = true;
            continue; // exits loop
        }

        if (WAIT_OBJECT_0 + 1 != dwWaitResult) {
            ERR(L"Unexpected WaitForMultipleObjects return value %u on pass %u after %u frames", dwWaitResult, nPasses, *pnFrames);
            return E_UNEXPECTED;
        }
    } // capture loop

    return hr;
}

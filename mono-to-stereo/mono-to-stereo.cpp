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

    UINT32 nBlockAlign = pwfx->nBlockAlign;
    *pnFrames = 0;

    // call IAudioClient::Initialize
    hr = pAudioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
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

    // create the event timer
    HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (hEvent == NULL)
    {
        DWORD dwErr = GetLastError();
        ERR(L"CreateEvent failed: last error = %u", dwErr);
        return HRESULT_FROM_WIN32(dwErr);
    }
    CloseHandleOnExit closeEvent(hEvent);

    pAudioClient->SetEventHandle(hEvent);

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
        (REFERENCE_TIME)iBufferMs * 10000,
        0,
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
    HANDLE waitArray[2] = { hStopEvent, hEvent };
    DWORD dwWaitResult;

    bool bDone = false;

    while (!bDone) {
        dwWaitResult = WaitForMultipleObjects(
            ARRAYSIZE(waitArray), waitArray,
            FALSE, INFINITE
        );

        if (WAIT_OBJECT_0 == dwWaitResult) {
            LOG(L"Received stop event after %u frames", *pnFrames);
            bDone = true;
            continue; // exits loop
        }

        if (WAIT_OBJECT_0 + 1 != dwWaitResult) {
            ERR(L"Unexpected WaitForMultipleObjects return value %u after %u frames", dwWaitResult, *pnFrames);
            return E_UNEXPECTED;
        }

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
            ERR(L"IAudioCaptureClient::GetBuffer failed after %u frames: hr = 0x%08x", *pnFrames, hr);
            return hr;
        }

        if (AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY == dwFlags) {
            LOG(L"Probably spurious glitch reported after %u frames", *pnFrames);
        }
        else if (0 != dwFlags) {
            LOG(L"IAudioCaptureClient::GetBuffer set flags to 0x%08x after %u frames", dwFlags, *pnFrames);
            return E_UNEXPECTED;
        }

        nNumFramesToRead &= ~1;

        if (0 == nNumFramesToRead) {
            LOG(L"IAudioCaptureClient::GetBuffer said to read 0 frames after %u frames", *pnFrames);
            continue;
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
                ERR(L"IAudioCaptureClient::GetBuffer failed (output) after %u frames: hr = 0x%08x", *pnFrames, hr);
                return hr;
            }
            break;
        }

        if (bSwapChannels) {
            swapMemcpy(pOutData, pData, lBytesToWrite, pwfx->wBitsPerSample / 8);
        }
        else {
            memcpy(pOutData, pData, lBytesToWrite);
        }

        hr = pRenderClient->ReleaseBuffer(nNumFramesToRead / 2, 0);
        if (FAILED(hr)) {
            ERR(L"IAudioCaptureClient::ReleaseBuffer failed (output) after %u frames: hr = 0x%08x", *pnFrames, hr);
            return hr;
        }

        hr = pAudioCaptureClient->ReleaseBuffer(nNumFramesToRead);
        if (FAILED(hr)) {
            ERR(L"IAudioCaptureClient::ReleaseBuffer failed after %u frames: hr = 0x%08x", *pnFrames, hr);
            return hr;
        }

        *pnFrames += nNumFramesToRead;
    } // capture loop

    return hr;
}

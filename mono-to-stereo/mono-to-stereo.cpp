// mono-to-stereo.cpp

#include "common.h"
#include <vector>

HRESULT LoopbackCapture(
    IMMDevice* pMMInDevice,
    IMMDevice* pMMOutDevice,
    int iBufferMs,
    bool bSkipFirstSample,
    HANDLE hStartedEvent,
    HANDLE hStopEvent,
    PUINT32 pnFrames
);

DWORD WINAPI LoopbackCaptureThreadFunction(LPVOID pContext) {
    LoopbackCaptureThreadFunctionArguments* pArgs =
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
        pArgs->bSkipFirstSample,
        pArgs->hStartedEvent,
        pArgs->hStopEvent,
        &pArgs->nFrames
    );

    return 0;
}

HRESULT LoopbackCapture(
    IMMDevice* pMMInDevice,
    IMMDevice* pMMOutDevice,
    int iBufferMs,
    bool bSkipFirstSample,
    HANDLE hStartedEvent,
    HANDLE hStopEvent,
    PUINT32 pnFrames
) {
    HRESULT hr;

    // activate an IAudioClient
    IAudioClient* pAudioClient;
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
    WAVEFORMATEX* pwfx;
    hr = pAudioClient->GetMixFormat(&pwfx);
    if (FAILED(hr)) {
        ERR(L"IAudioClient::GetMixFormat failed: hr = 0x%08x", hr);
        return hr;
    }
    CoTaskMemFreeOnExit freeMixFormat(pwfx);

    if (pwfx->nChannels != 1) {
        ERR(L"device doesn't have 1 channel, has %d", pwfx->nChannels);
        return E_UNEXPECTED;
    }

    if (pwfx->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        auto pwfxExtensible = reinterpret_cast<PWAVEFORMATEXTENSIBLE>(pwfx);
        if (!IsEqualGUID(KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, pwfxExtensible->SubFormat) && !IsEqualGUID(KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, pwfxExtensible->SubFormat)) {
            OLECHAR subFormatGUID[39];
            StringFromGUID2(pwfxExtensible->SubFormat, subFormatGUID, _countof(subFormatGUID));
            ERR(L"extensible input format not PCM, got %s", subFormatGUID);
            return E_UNEXPECTED;
        }
    }
    else if (pwfx->wFormatTag != WAVE_FORMAT_PCM && pwfx->wFormatTag != WAVE_FORMAT_IEEE_FLOAT) {
        ERR(L"input format not PCM, got %d", pwfx->wFormatTag);
        return E_UNEXPECTED;
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
    IAudioCaptureClient* pAudioCaptureClient;
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
    UINT32 OutputBlockAlign = pwfx->nBlockAlign;

    // set up output device
    IAudioClient* pAudioOutClient;
    hr = pMMOutDevice->Activate(
        __uuidof(IAudioClient), CLSCTX_ALL,
        NULL, (void**)&pAudioOutClient);
    if (FAILED(hr)) {
        ERR(L"IMMDevice::Activate(IAudioClient) failed (output): hr = 0x%08x", hr);
        return hr;
    }

    WAVEFORMATEX* pSupported = NULL;
    hr = pAudioOutClient->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, pwfx, &pSupported);
    if (hr != S_OK) {
        // S_FALSE or FAILED(hr)
        LOG(L" Channels %ld -> %ld", pwfx->nChannels, pSupported->nChannels);
        LOG(L" BitsPerSample %ld -> %ld", pwfx->wBitsPerSample, pSupported->wBitsPerSample);
        LOG(L" SamplesPerSec %ld -> %ld", pwfx->nSamplesPerSec, pSupported->nSamplesPerSec);
        ERR(L"IAudioClient::IsFormatSupported failed (output): hr = 0x%08x", hr);
        return hr;
    }

    hr = pAudioOutClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        0,
        static_cast<REFERENCE_TIME>(iBufferMs) * 10000,
        0,
        pwfx,
        NULL);
    if (FAILED(hr)) {
        ERR(L"IAudioClient::Initialize failed (output): hr = 0x%08x", hr);
        return hr;
    }

    IAudioRenderClient* pRenderClient;
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

    // Grab half the buffer for the initial fill operation.
    BYTE* tmp;
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

    std::vector<BYTE> lastSample;
    if (bSkipFirstSample) {
        lastSample.resize(nBlockAlign);
    }

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

        for (;;) {
            // get the captured data
            BYTE* pData;
            BYTE* pOutData;
            UINT32 nNextPacketSize;
            UINT32 nNumFramesToRead;
            DWORD dwFlags;

            hr = pAudioCaptureClient->GetNextPacketSize(&nNextPacketSize);
            if (FAILED(hr)) {
                ERR(L"IAudioCaptureClient::GetNextPacketSize failed after %u frames: hr = 0x%08x", *pnFrames, hr);
                return hr;
            }

            if (nNextPacketSize == 0) {
                break;
            }

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

            if (nNextPacketSize != nNumFramesToRead) {
                ERR(L"GetNextPacketSize and GetBuffer values don't match (%u and %u) after %u frames", nNextPacketSize, nNumFramesToRead, *pnFrames);
                return E_UNEXPECTED;
            }

            if (AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY == dwFlags) {
                if (*pnFrames != 0) {
                    LOG(L"Probably spurious glitch reported after %u frames", *pnFrames);
                }
            }
            else if (0 != dwFlags) {
                LOG(L"IAudioCaptureClient::GetBuffer set flags to 0x%08x after %u frames", dwFlags, *pnFrames);
                return E_UNEXPECTED;
            }

            if (nNumFramesToRead % 1 != 0) {
                ERR(L"frames to output is odd (%u), will miss the last sample after %u frames", nNumFramesToRead, *pnFrames);
            }

            UINT32 output_frames_to_write = nNumFramesToRead / 2;

            LONG lBytesToWrite = output_frames_to_write * OutputBlockAlign;

            for (;;) {
                hr = pRenderClient->GetBuffer(output_frames_to_write, &pOutData);
                if (hr == AUDCLNT_E_BUFFER_TOO_LARGE) {
                    ERR(L"%ls", L"buffer overflow!");
                    Sleep(1);
                    continue;
                }
                if (FAILED(hr)) {
                    ERR(L"IAudioCaptureClient::GetBuffer failed (output) after %u frames: hr = 0x%08x", *pnFrames, hr);
                    return hr;
                }
                break;
            }

            if (bSkipFirstSample) {
                memcpy(pOutData, lastSample.data(), nBlockAlign);
                memcpy(pOutData + nBlockAlign, pData, static_cast<size_t>(lBytesToWrite) - nBlockAlign);
                memcpy(lastSample.data(), pData + lBytesToWrite - nBlockAlign, nBlockAlign);
            }
            else {
                memcpy(pOutData, pData, lBytesToWrite);
            }

            hr = pRenderClient->ReleaseBuffer(output_frames_to_write, 0);
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
        }
    } // capture loop

    return hr;
}

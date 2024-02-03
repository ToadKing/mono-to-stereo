// mono-to-stereo.cpp

#include "common.h"
#include <vector>

void sndcpy(void *dst, WAVEFORMATEX* pwfxOut, const void *src, WAVEFORMATEX* pwfx, UINT nNumFrames, bool bDuplicateChannels);

HRESULT LoopbackCapture(
    IMMDevice* pMMInDevice,
    IMMDevice* pMMOutDevice,
    int iBufferMs,
    bool bSkipFirstSample,
    bool bForceMonoToStereo,
    bool bCaptureRenderer,
    bool bDuplicateChannels,
    CPreProcess &preProcess,
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
        pArgs->bForceMonoToStereo,
        pArgs->bCaptureRenderer,
        pArgs->bDuplicateChannels,
        pArgs->preProcess,
        pArgs->hStartedEvent,
        pArgs->hStopEvent,
        &pArgs->nFrames
    );

    return 0;
}

void ProcessSample(BYTE* pData, BYTE* tmp, UINT nBytes, CPreProcess &preProcess) {
    if (preProcess.m_bZeroLeft) {
        memset(pData, 0, nBytes);
    }

    if (preProcess.m_bZeroRight) {
        memset(pData + nBytes, 0, nBytes);
    }

    if (preProcess.m_bSwapChannels) {
        memcpy(tmp, pData, nBytes);
        memcpy(pData, pData + nBytes, nBytes);
        memcpy(pData + nBytes, tmp, nBytes);
    }
    else if (preProcess.m_bCopyToRight) {
        memcpy(pData + nBytes, pData, nBytes);
    }
    else if (preProcess.m_bCopyToLeft) {
        memcpy(pData, pData + nBytes, nBytes);
    }
}

HRESULT LoopbackCapture(
    IMMDevice* pMMInDevice,
    IMMDevice* pMMOutDevice,
    int iBufferMs,
    bool bSkipFirstSample,
    bool bForceMonoToStereo,
    bool bCaptureRenderer,
    bool bDuplicateChannels,
    CPreProcess &preProcess,
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

    if (bForceMonoToStereo && pwfx->nChannels != 1) {
        ERR(L"capture device doesn't have 1 channel, has %d", pwfx->nChannels);
        return E_UNEXPECTED;
    }

    if (pwfx->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        auto pwfxExtensible = reinterpret_cast<PWAVEFORMATEXTENSIBLE>(pwfx);
        if (!IsEqualGUID(KSDATAFORMAT_SUBTYPE_PCM, pwfxExtensible->SubFormat) && !IsEqualGUID(KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, pwfxExtensible->SubFormat)) {
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
    *pnFrames = 0;

    // create a periodic waitable timer
    HANDLE hWakeUp = NULL;
    if (bCaptureRenderer) {
        hWakeUp = CreateWaitableTimer(NULL, FALSE, NULL);
        if (NULL == hWakeUp) {
            DWORD dwErr = GetLastError();
            ERR(L"CreateWaitableTimer failed: last error = %u", dwErr);
            return HRESULT_FROM_WIN32(dwErr);
        }
    }
    CloseHandleOnExit closeWakeUp(hWakeUp);

    // call IAudioClient::Initialize
    hr = pAudioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        (bCaptureRenderer ? AUDCLNT_STREAMFLAGS_LOOPBACK : AUDCLNT_STREAMFLAGS_EVENTCALLBACK),
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
    HANDLE hEvent = NULL;
    if (bCaptureRenderer) {
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
    }
    else {
        hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        if (hEvent == NULL)
        {
            DWORD dwErr = GetLastError();
            ERR(L"CreateEvent failed: last error = %u", dwErr);
            return HRESULT_FROM_WIN32(dwErr);
        }

        pAudioClient->SetEventHandle(hEvent);
    }
    CancelWaitableTimerOnExit cancelWakeUp(hWakeUp);
    CloseHandleOnExit closeEvent(hEvent);

    // call IAudioClient::Start
    hr = pAudioClient->Start();
    if (FAILED(hr)) {
        ERR(L"IAudioClient::Start failed: hr = 0x%08x", hr);
        return hr;
    }
    AudioClientStopOnExit stopAudioClient(pAudioClient);

    if (bForceMonoToStereo) {
        // update format for stereo conversion
        pwfx->nChannels *= 2; // from mono to stereo
        pwfx->nSamplesPerSec /= 2;
        pwfx->nBlockAlign *= 2;
    }

    // set up output device
    IAudioClient* pAudioOutClient;
    hr = pMMOutDevice->Activate(
        __uuidof(IAudioClient),
        CLSCTX_ALL, NULL,
        (void**)&pAudioOutClient
    );
    if (FAILED(hr)) {
        ERR(L"IMMDevice::Activate(IAudioClient) failed (output): hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseAudioOutClient(pAudioOutClient);

    WAVEFORMATEX* pwfxOut = NULL;
    hr = pAudioOutClient->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, pwfx, &pwfxOut);
    if (hr == S_OK) {
        pwfxOut = pwfx;
    }
    LOG(L" Channels %ld -> %ld", pwfx->nChannels, pwfxOut->nChannels);
    LOG(L" BitsPerSample %ld -> %ld", pwfx->wBitsPerSample, pwfxOut->wBitsPerSample);
    LOG(L" SamplesPerSec %ld -> %ld", pwfx->nSamplesPerSec, pwfxOut->nSamplesPerSec);
    if (FAILED(hr))
    {
        ERR(L"IAudioClient::IsFormatSupported failed (output): hr = 0x%08x", hr);
        return hr;
    }

    if (pwfx->nSamplesPerSec != pwfxOut->nSamplesPerSec) {
        ERR(L"cannot convert between sample rates");
        return E_UNEXPECTED;
    }

    hr = pAudioOutClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED, 0,
        static_cast<REFERENCE_TIME>(iBufferMs)* 10000,
        0, pwfxOut, 0
    );
    if (FAILED(hr)) {
        ERR(L"IAudioClient::Initialize failed (output): hr = 0x%08x", hr);
        return hr;
    }

    IAudioRenderClient* pRenderClient;
    hr = pAudioOutClient->GetService(
        __uuidof(IAudioRenderClient),
        (void**)&pRenderClient
    );
    if (FAILED(hr)) {
        ERR(L"IAudioClient::GetService(IAudioRenderClient) failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseRenderClient(pRenderClient);

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
    AudioClientStopOnExit stopAudioOutClient(pAudioOutClient);

    SetEvent(hStartedEvent);

    // loopback capture loop
    HANDLE waitArray[2] = { hStopEvent, (bCaptureRenderer ? hWakeUp : hEvent) };
    DWORD dwWaitResult;

    bool bDone = false;

    std::vector<BYTE> lastBlock;
    if (bSkipFirstSample) {
        lastBlock.resize(pwfx->nBlockAlign);
    }

    UINT nBytes = pwfx->wBitsPerSample / 8;
    std::vector<BYTE> tmpSample(nBytes);

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

            hr = pAudioCaptureClient->GetBuffer(&pData, &nNumFramesToRead, &dwFlags, NULL, NULL);
            if (FAILED(hr)) {
                ERR(L"IAudioCaptureClient::GetBuffer failed after %u frames: hr = 0x%08x", *pnFrames, hr);
                return hr;
            }

            if (nNextPacketSize != nNumFramesToRead) {
                ERR(L"GetNextPacketSize and GetBuffer values don't match (%u and %u) after %u frames", nNextPacketSize, nNumFramesToRead, *pnFrames);
                // release capture buffer before exiting
                pAudioCaptureClient->ReleaseBuffer(nNumFramesToRead);
                return E_UNEXPECTED;
            }

            if (AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY == dwFlags) {
                if (*pnFrames != 0) {
                    LOG(L"Probably spurious glitch reported after %u frames", *pnFrames);
                }
            }
            else if (0 != dwFlags) {
                LOG(L"IAudioCaptureClient::GetBuffer set flags to 0x%08x after %u frames", dwFlags, *pnFrames);
                // release capture buffer before exiting
                pAudioCaptureClient->ReleaseBuffer(nNumFramesToRead);
                return E_UNEXPECTED;
            }

            if (nNumFramesToRead % 2 != 0) {
                ERR(L"frames to output is odd (%u), will miss the last sample after %u frames", nNumFramesToRead, *pnFrames);
            }

            // pre-process audio data
            if (pwfx->nChannels >= 2 && preProcess.IsRequired()) {
                for (UINT i = 0; i < nNumFramesToRead; i++) {
                    BYTE* pOffset = pData + (i * pwfx->nBlockAlign);
                    switch (pwfx->nChannels) {
                    case 8:
                        // ML and MR
                        ProcessSample(pOffset + (nBytes * 6), tmpSample.data(), nBytes, preProcess);
                    case 6:
                    case 4:
                        // RL and RR
                        ProcessSample(pOffset + (nBytes * 2), tmpSample.data(), nBytes, preProcess);
                    case 2:
                        // L and R
                        ProcessSample(pOffset, tmpSample.data(), nBytes, preProcess);
                        break;
                    }
                }
            }

            // this is halved because nNumFramesToRead came from a mono source
            UINT32 nNumFramesToWrite = (bForceMonoToStereo ? nNumFramesToRead / 2 : nNumFramesToRead);

            for (;;) {
                hr = pRenderClient->GetBuffer(nNumFramesToWrite, &pOutData);
                if (hr == AUDCLNT_E_BUFFER_TOO_LARGE) {
                    ERR(L"%ls", L"buffer overflow!");
                    Sleep(1);
                    continue;
                }
                if (FAILED(hr)) {
                    ERR(L"IAudioCaptureClient::GetBuffer failed (output) after %u frames: hr = 0x%08x", *pnFrames, hr);
                    // release capture buffer before exiting
                    pAudioCaptureClient->ReleaseBuffer(nNumFramesToRead);
                    return hr;
                }
                break;
            }

            if (bSkipFirstSample) {
                sndcpy(pOutData, pwfxOut, lastBlock.data(), pwfx, 1, bDuplicateChannels);
                sndcpy(pOutData + pwfxOut->nBlockAlign, pwfxOut, pData, pwfx, nNumFramesToWrite - 1, bDuplicateChannels);
                memcpy(lastBlock.data(), pData + ((nNumFramesToWrite - 1) * pwfx->nBlockAlign), pwfx->nBlockAlign);
            }
            else {
                sndcpy(pOutData, pwfxOut, pData, pwfx, nNumFramesToWrite, bDuplicateChannels);
            }

            hr = pRenderClient->ReleaseBuffer(nNumFramesToWrite, 0);
            if (FAILED(hr)) {
                ERR(L"IAudioCaptureClient::ReleaseBuffer failed (output) after %u frames: hr = 0x%08x", *pnFrames, hr);
                // release capture buffer before exiting
                pAudioCaptureClient->ReleaseBuffer(nNumFramesToRead);
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

void valcpy(BYTE* pOutData, UINT nOutBytes, BYTE* pData, UINT nBytes) {
    if (nOutBytes >= nBytes) {
        for (UINT i = 0; i < nBytes; i++) {
            pOutData[i] = pData[i];
        }
        if (nOutBytes > nBytes) {
            UINT leftShift = nOutBytes - nBytes;
            for (UINT i = 0; i < leftShift; i++) {
                pOutData[leftShift + i] = 0;
            }
        }
    }
    else { // (nOutBytes < nBytes)
        for (UINT i = 0; i < nOutBytes; i++) {
            pOutData[i] = pData[i];
        }
    }
}

void sndcpy(void *dst, WAVEFORMATEX* pwfxOut, const void *src, WAVEFORMATEX* pwfx, UINT nNumFrames, bool bDuplicateChannels) {

    // does not support converting between sample different sample rates, would require some form of interpolation??

    if (nNumFrames == 0 || pwfx->nSamplesPerSec != pwfxOut->nSamplesPerSec) {
        // do nothing
        return;
    }
    else if (pwfx->nChannels == pwfxOut->nChannels && pwfx->wBitsPerSample == pwfxOut->wBitsPerSample) {
        // same format, raw copy
        memcpy(dst, src, nNumFrames * pwfx->nBlockAlign);
        return;
    }
    else {
        BYTE *in = (BYTE *)src;
        BYTE *out = (BYTE *)dst;

        UINT nBytes = pwfx->wBitsPerSample / 8;
        UINT nOutBytes = pwfxOut->wBitsPerSample / 8;

        std::vector<BYTE> channel(nOutBytes);

        BYTE *pOutData, *pData;
        UINT nFrameOffset = 0;
        while (nFrameOffset < nNumFrames) {
            pData = &in[nFrameOffset * pwfx->nBlockAlign];
            pOutData = &out[nFrameOffset * pwfxOut->nBlockAlign];

            // copy in L to out L always
            valcpy(pOutData, nOutBytes, pData, nBytes);

            if (pwfxOut->nChannels == 1) { // mono
                // do nothing
            }
            if (pwfxOut->nChannels >= 2) { // stereo
                if (pwfx->nChannels >= 2) {
                    // copy in R to out R
                    valcpy(pOutData + nOutBytes, nOutBytes, pData + nBytes, nBytes);
                }
                else if(bDuplicateChannels) {
                    // copy out L to out R
                    memcpy(pOutData + nOutBytes, pOutData, nOutBytes);
                }
            }
            if (pwfxOut->nChannels >= 4) { // quad
                if (pwfx->nChannels >= 4) {
                    // copy in LR to out LR
                    valcpy(pOutData + (nOutBytes * 2), nOutBytes, pData + (nBytes * 2), nBytes);
                    // copy in RR to out RR
                    valcpy(pOutData + (nOutBytes * 3), nOutBytes, pData + (nBytes * 3), nBytes);
                }
                else if (bDuplicateChannels) {
                    // copy out L and R to out RL and RR
                    memcpy(pOutData + (nOutBytes * 2), pOutData, nOutBytes * 2);
                }
            }
            if (pwfxOut->nChannels >= 6) { // 5.1
                if (pwfx->nChannels >= 6) {
                    // copy in C to out C
                    valcpy(pOutData + (nOutBytes * 4), nOutBytes, pData + (nBytes * 4), nBytes);
                    // copy in SW to out SW
                    valcpy(pOutData + (nOutBytes * 5), nOutBytes, pData + (nBytes * 5), nBytes);
                }
            }
            if (pwfxOut->nChannels >= 8) { // 7.1
                if (pwfx->nChannels >= 8) {
                    // copy in ML to out ML
                    valcpy(pOutData + (nOutBytes * 6), nOutBytes, pData + (nBytes * 6), nBytes);
                    // copy in MR to out MR
                    valcpy(pOutData + (nOutBytes * 7), nOutBytes, pData + (nBytes * 7), nBytes);
                }
                else if (bDuplicateChannels) {
                    // copy out L and R to out RL and RR
                    // NOTE: copy from front speakers, not rear speakers
                    memcpy(pOutData + (nOutBytes * 2), pOutData, nOutBytes * 2);
                }
            }

            nFrameOffset++;
        }
    }
}

// mono-to-stereo.cpp

#include "common.h"
#include <vector>

HRESULT LoopbackCapture(
    IMMDevice* pMMInDevice,
    IMMDevice* pMMOutDevice,
    int iBufferMs,
    bool bDuplicateChannels,
    bool bForceMonoToStereo,
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
        pArgs->bDuplicateChannels,
        pArgs->bForceMonoToStereo,
        pArgs->bSkipFirstSample,
        pArgs->hStartedEvent,
        pArgs->hStopEvent,
        &pArgs->nFrames
    );

    return 0;
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

            // copy first channel always
            valcpy(pOutData, nOutBytes, pData, nBytes);

            switch (pwfx->nChannels) {
            case 1: // mono
                if (!bDuplicateChannels) {
                    break;
                }

                switch (pwfxOut->nChannels) {
                case 2: // stereo
                    // copy out L to out R
                    for (UINT i = 0; i < nOutBytes; i++) {
                        pOutData[nOutBytes + i] = pOutData[i];
                    }
                case 4: // quad
                case 6: // 5.1
                case 8: // 7.1
                    // copy out L to R, RL, RR
                    UINT nTotal = nOutBytes * 3;
                    for (UINT i = 0; i < nTotal; i++) {
                        pOutData[nOutBytes + i] = pOutData[i % nOutBytes];
                    }
                    // fill remainder of channels with blank
                    UINT nZero = nOutBytes * (pwfxOut->nChannels - 4);
                    for (UINT i = 0; i < nZero; i++) {
                        pOutData[nOutBytes + nTotal + i] = 0;
                    }
                    break;
                }
                break;
            case 2: // stereo
                switch (pwfxOut->nChannels) {
                case 1: // mono
                    // average in R to out L, R 
                    valcpy(channel.data(), nOutBytes, pData + nBytes, nBytes);

                    // ? ? ? TODO: take average of both L and temp R value and overwrite out L

                    break;
                case 2: // stereo
                case 4: // quad
                case 6: // 5.1
                case 8: // 7.1
                    // copy in R to out R
                    valcpy(pOutData + nOutBytes, nOutBytes, pData + nBytes, nBytes);

                    if (!bDuplicateChannels || pwfxOut->nChannels < 4) {
                        break;
                    }

                    // copy out L, R to out RL, RR
                    UINT nOffset = nOutBytes * 2;
                    UINT nTotal = nOutBytes * 2;
                    for (UINT i = 0; i < nTotal; i++) {
                        pOutData[nOffset + i] = pOutData[i];
                    }
                    // fill remainder of channels with blank
                    UINT nZero = nOutBytes * (pwfxOut->nChannels - 4);
                    for (UINT i = 0; i < nZero; i++) {
                        pOutData[nOffset + nTotal + i] = 0;
                    }
                    break;
                }
                break;
                break;
            case 4: // quad
                break;
            case 6: // 5.1
                break;
            case 8: // 7.1
                break;
            }

            nFrameOffset++;
        }
    }
}

HRESULT LoopbackCapture(
    IMMDevice* pMMInDevice,
    IMMDevice* pMMOutDevice,
    int iBufferMs,
    bool bDuplicateChannels,
    bool bForceMonoToStereo,
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

    if (pwfx->nChannels >= 4) {
        ERR(L"cannot convert quaud, 5.1 opr 7.1 channel capture audio");
        return E_UNEXPECTED;
    }

    hr = pAudioOutClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED, 0,
        static_cast<REFERENCE_TIME>(iBufferMs) * 10000,
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
    ReleaseOnExit releaseAudioRenderClient(pRenderClient);

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
    HANDLE waitArray[2] = { hStopEvent, hEvent };
    DWORD dwWaitResult;

    bool bDone = false;

    std::vector<BYTE> lastSample;
    if (bSkipFirstSample) {
        lastSample.resize(pwfx->nBlockAlign);
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

            // this is halved because nNumFramesToRead came from a mono source
            UINT32 nNumFramesToWrite = (bForceMonoToStereo ? nNumFramesToRead / 2 : nNumFramesToRead);

            for (;;) {
                hr = pRenderClient->GetBuffer(nNumFramesToWrite, &pOutData);
                if (hr == AUDCLNT_E_BUFFER_TOO_LARGE) {
                    ERR(L"%ls", L"buffer overflow!");
                    Sleep(1);
                    continue;
                }
                else if (FAILED(hr)) {
                    ERR(L"IAudioCaptureClient::GetBuffer failed (output) after %u frames: hr = 0x%08x", *pnFrames, hr);
                    // release capture buffer before exiting
                    pAudioCaptureClient->ReleaseBuffer(nNumFramesToRead);
                    return hr;
                }
                break;
            }

            //static_cast<size_t>(lBytesWeRead - pwfx->nBlockAlign)
            if (bSkipFirstSample) {
                sndcpy(pOutData, pwfxOut, lastSample.data(), pwfx,1, bDuplicateChannels);
                sndcpy(pOutData + pwfxOut->nBlockAlign, pwfxOut, pData, pwfx, nNumFramesToWrite - 1, bDuplicateChannels);
                memcpy(lastSample.data(), pData + ((nNumFramesToWrite - 1) * pwfx->nBlockAlign), pwfx->nBlockAlign);
            }
            else {
                sndcpy(pOutData, pwfxOut, pData, pwfx, nNumFramesToWrite, bDuplicateChannels);
            }

            hr = pRenderClient->ReleaseBuffer(nNumFramesToWrite, 0);
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

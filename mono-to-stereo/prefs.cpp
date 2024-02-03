// prefs.cpp

#pragma warning( disable : 4996 )

#include "common.h"

#define DEFAULT_BUFFER_MS 64

void usage(LPCWSTR exe);
HRESULT get_default_device(EDataFlow direction, IMMDevice **ppMMDevice, LPWSTR szOutName);
HRESULT list_devices();
HRESULT list_devices_with_direction(EDataFlow direction, const wchar_t *direction_label);
HRESULT get_specific_device(LPCWSTR szLongName, EDataFlow direction, IMMDevice **ppMMDevice, LPWSTR szOutName);

void usage(LPCWSTR exe) {
    LOG(
        L"mono-to-stereo v%s\n"
        L"\n"
        L"%ls -?\n"
        L"%ls --list-devices\n"
        L"%ls [--capture-renderer] [--in-device \"Device long name\"|DeviceIndex] [--out-device \"Device long name\"|DeviceIndex] [--buffer-size 128] [--no-skip-first-sample]\n"
        L" [--no-mono-to-stereo] [--duplicate-channels] [--zero-left] [--zero-right] [--swap-channels] [--copy-to-right] [--copy-to-left]\n"
        L"\n"
        L"    -? prints this message.\n"
        L"    --list-devices displays the long names of all active capture and render devices.\n"
        L"    --capture-renderer captures from a render device instead (needs to be set before --in-device)\n"
        L"    --in-device captures from the specified device to capture (\"Digital Audio Interface (USB Digital Audio)\" if omitted)\n"
        L"    --out-device device to stream stereo audio to (default if omitted)\n"
        L"    --buffer-size set the size of the audio buffer, in milliseconds (default to %dms)\n"
        L"    --no-skip-first-sample do not skip the first channel sample\n"
        L"    --no-mono-to-stereo do not force mono capture device to be treated as stereo without conversion if mono device\n"
        L"    --duplicate-channels duplicate left and right audio if render device has more channels\n"
        L"    --zero-left zero out all capture left channels\n"
        L"    --zero-right zero out all capture right channels\n"
        L"    --swap-channels swap all available left and right channels (cannot enable if already copying to right or left)\n"
        L"    --copy-to-right copy left channel data to right channel (cannot enable if already swapping or copying to left)\n"
        L"    --copy-to-left copy right channel data to left channel (cannot enable if already swapping or copying to right)\n",
        VERSION, exe, exe, exe, DEFAULT_BUFFER_MS

        );
}


CPreProcess::CPreProcess()
    : m_bZeroRight(false)
    , m_bZeroLeft(false)
    , m_bSwapChannels(false)
    , m_bCopyToRight(false)
    , m_bCopyToLeft(false)
{ }


bool CPreProcess::IsRequired() {
    return (m_bZeroLeft || m_bZeroRight|| m_bSwapChannels || m_bCopyToRight || m_bCopyToLeft);
}


CPrefs::CPrefs(int argc, LPCWSTR argv[], HRESULT &hr)
    : m_pMMInDevice(NULL)
    , m_pMMOutDevice(NULL)
    , m_iBufferMs(DEFAULT_BUFFER_MS)
    , m_bSkipFirstSample(true)
    , m_bForceMonoToStereo(true)
    , m_bCaptureRenderer(false)
    , m_bDuplicateChannels(false)
{
    switch (argc) {
    case 2:
        if (0 == _wcsicmp(argv[1], L"-?") || 0 == _wcsicmp(argv[1], L"/?")) {
            // print usage but don't actually capture
            hr = S_FALSE;
            usage(argv[0]);
            return;
        }
        else if (0 == _wcsicmp(argv[1], L"--list-devices")) {
            // list the devices but don't actually capture
            hr = list_devices();

            // don't actually play
            if (S_OK == hr) {
                hr = S_FALSE;
                return;
            }
        }
        // intentional fallthrough

    default:
        // loop through arguments and parse them
        for (int i = 1; i < argc; i++) {

            // --in-device
            if (0 == _wcsicmp(argv[i], L"--in-device")) {
                if (NULL != m_pMMInDevice) {
                    ERR(L"%s", L"Only one --device switch is allowed");
                    hr = E_INVALIDARG;
                    return;
                }

                if (i++ == argc) {
                    ERR(L"%s", L"--device switch requires an argument");
                    hr = E_INVALIDARG;
                    return;
                }

                hr = get_specific_device(argv[i], (m_bCaptureRenderer ? eRender : eCapture), &m_pMMInDevice, m_szInName);
                if (FAILED(hr)) {
                    return;
                }

                continue;
            }

            // --out-device
            if (0 == _wcsicmp(argv[i], L"--out-device")) {
                if (NULL != m_pMMOutDevice) {
                    ERR(L"%s", L"Only one --device switch is allowed");
                    hr = E_INVALIDARG;
                    return;
                }

                if (i++ == argc) {
                    ERR(L"%s", L"--device switch requires an argument");
                    hr = E_INVALIDARG;
                    return;
                }

                hr = get_specific_device(argv[i], eRender, &m_pMMOutDevice, m_szOutName);
                if (FAILED(hr)) {
                    return;
                }

                continue;
            }

            // --buffer-size
            if (0 == _wcsicmp(argv[i], L"--buffer-size")) {
                if (i++ == argc) {
                    ERR(L"%s", L"--buffer-size switch requires an argument");
                    hr = E_INVALIDARG;
                    return;
                }

                m_iBufferMs = _wtoi(argv[i]);
                if (m_iBufferMs <= 0) {
                    ERR(L"%s", L"invalid buffer size given");
                    hr = E_INVALIDARG;
                    return;
                }

                continue;
            }

            // --skip-first-sample
            if (0 == _wcsicmp(argv[i], L"--no-skip-first-sample")) {
                m_bSkipFirstSample = false;
                continue;
            }

            // --force-mono-to-stereo
            if (0 == _wcsicmp(argv[i], L"--no-mono-to-stereo")) {
                m_bForceMonoToStereo = false;
                continue;
            }

            // --capture-renderer
            if (0 == _wcsicmp(argv[i], L"--capture-renderer")) {
                if (NULL != m_pMMInDevice) {
                    ERR(L"%s", L"--capture-renderer is required before --in-device is selected");
                    hr = E_INVALIDARG;
                    return;
                }

                m_bCaptureRenderer = true;
                continue;
            }

            // --duplicate-channels
            if (0 == _wcsicmp(argv[i], L"--duplicate-channels")) {
                m_bDuplicateChannels = true;
                continue;
            }

            // --zero-left
            if (0 == _wcsicmp(argv[i], L"--zero-left")) {
                m_preProcess.m_bZeroLeft = true;
                continue;
            }

            // --zero-right
            if (0 == _wcsicmp(argv[i], L"--zero-right")) {
                m_preProcess.m_bZeroRight = true;
                continue;
            }

            // --swap-channels
            if (0 == _wcsicmp(argv[i], L"--swap-channels")) {
                if (m_preProcess.m_bCopyToRight || m_preProcess.m_bCopyToLeft) {
                    ERR(L"%s", L"--copy-to-right or --copy-to-left is already enabled");
                    hr = E_INVALIDARG;
                }

                m_preProcess.m_bSwapChannels = true;
                continue;
            }

            // --copy-to-right
            if (0 == _wcsicmp(argv[i], L"--copy-to-right")) {
                if (m_preProcess.m_bSwapChannels || m_preProcess.m_bCopyToLeft) {
                    ERR(L"%s", L"--swap-channels or --copy-to-left is already enabled");
                    hr = E_INVALIDARG;
                }

                m_preProcess.m_bCopyToRight = true;
                continue;
            }

            // --copy-to-left
            if (0 == _wcsicmp(argv[i], L"--copy-to-left")) {
                if (m_preProcess.m_bSwapChannels || m_preProcess.m_bCopyToRight) {
                    ERR(L"%s", L"--swap-channels or --copy-to-right is already enabled");
                    hr = E_INVALIDARG;
                }

                m_preProcess.m_bCopyToLeft = true;
                continue;
            }

            ERR(L"Invalid argument %ls", argv[i]);
            hr = E_INVALIDARG;
            return;
        }

        // open default device if not specified
        if (NULL == m_pMMInDevice) {
            if (m_bForceMonoToStereo && !m_bCaptureRenderer) {
                hr = get_specific_device(L"Digital Audio Interface (USB Digital Audio)", eCapture, &m_pMMInDevice, m_szInName);
            }
            else {
                hr = get_default_device((m_bCaptureRenderer ? eRender : eCapture), &m_pMMInDevice, m_szInName);
            }
            if (FAILED(hr)) {
                return;
            }
        }

        // open default device if not specified
        if (NULL == m_pMMOutDevice) {
            hr = get_default_device(eRender, &m_pMMOutDevice, m_szOutName);
            if (FAILED(hr)) {
                return;
            }
        }
    }
}

CPrefs::~CPrefs() {
    if (NULL != m_pMMInDevice) {
        m_pMMInDevice->Release();
    }

    if (NULL != m_pMMOutDevice) {
        m_pMMOutDevice->Release();
    }
}

HRESULT get_default_device(EDataFlow direction, IMMDevice **ppMMDevice, LPWSTR szOutName) {
    HRESULT hr = S_OK;
    IMMDeviceEnumerator *pMMDeviceEnumerator;

    // activate a device enumerator
    hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator),
        (void**)&pMMDeviceEnumerator
    );
    if (FAILED(hr)) {
        ERR(L"CoCreateInstance(IMMDeviceEnumerator) failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseMMDeviceEnumerator(pMMDeviceEnumerator);

    // get the default endpoint
    hr = pMMDeviceEnumerator->GetDefaultAudioEndpoint(direction, eConsole, ppMMDevice);
    if (FAILED(hr)) {
        ERR(L"IMMDeviceEnumerator::GetDefaultAudioEndpoint failed: hr = 0x%08x", hr);
        return hr;
    }

    // open the property store on that device
    IPropertyStore *pPropertyStore;
    hr = (*ppMMDevice)->OpenPropertyStore(STGM_READ, &pPropertyStore);
    if (FAILED(hr)) {
        ERR(L"IMMDevice::OpenPropertyStore failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releasePropertyStore(pPropertyStore);

    // get the long name property
    PROPVARIANT pv; PropVariantInit(&pv);
    hr = pPropertyStore->GetValue(PKEY_Device_FriendlyName, &pv);
    if (FAILED(hr)) {
        ERR(L"IPropertyStore::GetValue failed: hr = 0x%08x", hr);
        return hr;
    }
    PropVariantClearOnExit clearPv(&pv);

    if (VT_LPWSTR != pv.vt) {
        ERR(L"PKEY_Device_FriendlyName variant type is %u - expected VT_LPWSTR", pv.vt);
        return E_UNEXPECTED;
    }

    if (szOutName) {
        wcscpy(szOutName, pv.pwszVal);
    }

    return S_OK;
}

HRESULT list_devices() {
    HRESULT hr;

    hr = list_devices_with_direction(eRender, L"render");
    if (FAILED(hr)) {
        return hr;
    }

    LOG(L"");

    hr = list_devices_with_direction(eCapture, L"capture");
    if (FAILED(hr)) {
        return hr;
    }

    return hr;
}

HRESULT list_devices_with_direction(EDataFlow direction, const wchar_t *direction_label) {
    HRESULT hr = S_OK;

    // get an enumerator
    IMMDeviceEnumerator *pMMDeviceEnumerator;

    hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator),
        (void**)&pMMDeviceEnumerator
    );
    if (FAILED(hr)) {
        ERR(L"CoCreateInstance(IMMDeviceEnumerator) failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseMMDeviceEnumerator(pMMDeviceEnumerator);

    IMMDeviceCollection *pMMDeviceCollection;

    // get all the active render endpoints
    hr = pMMDeviceEnumerator->EnumAudioEndpoints(
        direction, DEVICE_STATE_ACTIVE, &pMMDeviceCollection
    );
    if (FAILED(hr)) {
        ERR(L"IMMDeviceEnumerator::EnumAudioEndpoints failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseMMDeviceCollection(pMMDeviceCollection);

    UINT count;
    hr = pMMDeviceCollection->GetCount(&count);
    if (FAILED(hr)) {
        ERR(L"IMMDeviceCollection::GetCount failed: hr = 0x%08x", hr);
        return hr;
    }
    LOG(L"Active %s endpoints found: %u", direction_label, count);

    for (UINT i = 0; i < count; i++) {
        IMMDevice *pMMDevice;

        // get the "n"th device
        hr = pMMDeviceCollection->Item(i, &pMMDevice);
        if (FAILED(hr)) {
            ERR(L"IMMDeviceCollection::Item failed: hr = 0x%08x", hr);
            return hr;
        }
        ReleaseOnExit releaseMMDevice(pMMDevice);

        // open the property store on that device
        IPropertyStore *pPropertyStore;
        hr = pMMDevice->OpenPropertyStore(STGM_READ, &pPropertyStore);
        if (FAILED(hr)) {
            ERR(L"IMMDevice::OpenPropertyStore failed: hr = 0x%08x", hr);
            return hr;
        }
        ReleaseOnExit releasePropertyStore(pPropertyStore);

        // get the long name property
        PROPVARIANT pv; PropVariantInit(&pv);
        hr = pPropertyStore->GetValue(PKEY_Device_FriendlyName, &pv);
        if (FAILED(hr)) {
            ERR(L"IPropertyStore::GetValue failed: hr = 0x%08x", hr);
            return hr;
        }
        PropVariantClearOnExit clearPv(&pv);

        if (VT_LPWSTR != pv.vt) {
            ERR(L"PKEY_Device_FriendlyName variant type is %u - expected VT_LPWSTR", pv.vt);
            return E_UNEXPECTED;
        }

        LOG(L"%u)    '%ls'", i + 1, pv.pwszVal);
    }

    return S_OK;
}

BOOL is_numeric(LPCWSTR szLongName, LPLONG lpOut) {

    WCHAR* p;

    LONG lResult = wcstol(szLongName, &p, 10);

    if (*p == 0) { // at end of wide string
        if (lpOut)
            *lpOut = lResult;
        return TRUE;
    }

    return FALSE;
}

HRESULT get_specific_device(LPCWSTR szLongName, EDataFlow direction, IMMDevice **ppMMDevice, LPWSTR szOutName) {
    HRESULT hr = S_OK;

    *ppMMDevice = NULL;

    // get an enumerator
    IMMDeviceEnumerator *pMMDeviceEnumerator;

    hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator),
        (void**)&pMMDeviceEnumerator
    );
    if (FAILED(hr)) {
        ERR(L"CoCreateInstance(IMMDeviceEnumerator) failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseMMDeviceEnumerator(pMMDeviceEnumerator);

    IMMDeviceCollection *pMMDeviceCollection;

    // get all the active render endpoints
    hr = pMMDeviceEnumerator->EnumAudioEndpoints(
        direction, DEVICE_STATE_ACTIVE, &pMMDeviceCollection
    );
    if (FAILED(hr)) {
        ERR(L"IMMDeviceEnumerator::EnumAudioEndpoints failed: hr = 0x%08x", hr);
        return hr;
    }
    ReleaseOnExit releaseMMDeviceCollection(pMMDeviceCollection);

    LONG lIndex = 0;
    is_numeric(szLongName, &lIndex);

    // limit number of deviecs to 1024 so its not bigger then a UINT
    UINT index = (lIndex > 0L && lIndex <= 1024L) ? (UINT)lIndex : 0;

    UINT count;
    hr = pMMDeviceCollection->GetCount(&count);
    if (FAILED(hr)) {
        ERR(L"IMMDeviceCollection::GetCount failed: hr = 0x%08x", hr);
        return hr;
    }

    for (UINT i = 0; i < count; i++) {
        IMMDevice *pMMDevice;

        // get the "n"th device
        hr = pMMDeviceCollection->Item(i, &pMMDevice);
        if (FAILED(hr)) {
            ERR(L"IMMDeviceCollection::Item failed: hr = 0x%08x", hr);
            return hr;
        }
        ReleaseOnExit releaseMMDevice(pMMDevice);

        // open the property store on that device
        IPropertyStore *pPropertyStore;
        hr = pMMDevice->OpenPropertyStore(STGM_READ, &pPropertyStore);
        if (FAILED(hr)) {
            ERR(L"IMMDevice::OpenPropertyStore failed: hr = 0x%08x", hr);
            return hr;
        }
        ReleaseOnExit releasePropertyStore(pPropertyStore);

        // get the long name property
        PROPVARIANT pv; PropVariantInit(&pv);
        hr = pPropertyStore->GetValue(PKEY_Device_FriendlyName, &pv);
        if (FAILED(hr)) {
            ERR(L"IPropertyStore::GetValue failed: hr = 0x%08x", hr);
            return hr;
        }
        PropVariantClearOnExit clearPv(&pv);

        if (VT_LPWSTR != pv.vt) {
            ERR(L"PKEY_Device_FriendlyName variant type is %u - expected VT_LPWSTR", pv.vt);
            return E_UNEXPECTED;
        }

        // is it a match?
        if (index > 0) {
            if (index == (i + 1)) {
                *ppMMDevice = pMMDevice;
                pMMDevice->AddRef();

                if (szOutName) {
                    wcscpy(szOutName, pv.pwszVal);
                }

                // exit loop if searching by index
                break;
            }
        }
        else if (0 == _wcsicmp(pv.pwszVal, szLongName)) {
            // did we already find it?
            if (NULL == *ppMMDevice) {
                *ppMMDevice = pMMDevice;
                pMMDevice->AddRef();

                if (szOutName) {
                    wcscpy(szOutName, pv.pwszVal);
                }
            }
            else {
                ERR(L"Found (at least) two devices named %ls", szLongName);
                return E_UNEXPECTED;
            }
        }

    }

    if (NULL == *ppMMDevice) {
        ERR(L"Could not find a device named %ls", szLongName);
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    return S_OK;
}

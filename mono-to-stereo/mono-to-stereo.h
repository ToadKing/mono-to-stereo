// mono-to-stereo.h

// call CreateThread on this function
// feed it the address of a LoopbackCaptureThreadFunctionArguments
// it will capture via loopback from the IMMDevice
// and dump output to the HMMIO
// until the stop event is set
// any failures will be propagated back via hr

#define VERSION L"0.3"

struct LoopbackCaptureThreadFunctionArguments {
    IMMDevice *pMMInDevice;
    IMMDevice *pMMOutDevice;
    int iBufferMs;
    bool bSwapChannels;
    HANDLE hStartedEvent;
    HANDLE hStopEvent;
    UINT32 nFrames;
    HRESULT hr;
};

DWORD WINAPI LoopbackCaptureThreadFunction(LPVOID pContext);

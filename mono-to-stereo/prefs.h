// prefs.h

class CPrefs {
public:
    IMMDevice *m_pMMInDevice;
    IMMDevice *m_pMMOutDevice;
    WCHAR szInName[FILENAME_MAX];
    WCHAR szOutName[FILENAME_MAX];
    int m_iBufferMs;
    bool m_bCaptureRenderer;
    bool m_bSwapChannels;
    bool m_bCopyToRight;
    bool m_bCopyToLeft;
    bool m_bDuplicateChannels;
    bool m_bForceMonoToStereo;
    bool m_bSkipFirstSample;

    // set hr to S_FALSE to abort but return success
    CPrefs(int argc, LPCWSTR argv[], HRESULT &hr);
    ~CPrefs();

};

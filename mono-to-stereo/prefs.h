// prefs.h

class CPreProcess {
public:
    bool m_bZeroLeft;
    bool m_bZeroRight;
    bool m_bSwapChannels;
    bool m_bCopyToRight;
    bool m_bCopyToLeft;

    CPreProcess();

    bool IsRequired();

};

class CPrefs {
public:
    IMMDevice *m_pMMInDevice;
    IMMDevice *m_pMMOutDevice;
    WCHAR m_szInName[FILENAME_MAX];
    WCHAR m_szOutName[FILENAME_MAX];
    int m_iBufferMs;
    bool m_bSkipFirstSample;
    bool m_bForceMonoToStereo;
    bool m_bCaptureRenderer;
    bool m_bDuplicateChannels;
    CPreProcess m_preProcess;

    // set hr to S_FALSE to abort but return success
    CPrefs(int argc, LPCWSTR argv[], HRESULT &hr);
    ~CPrefs();

};

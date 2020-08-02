// prefs.h

class CPrefs {
public:
    IMMDevice *m_pMMInDevice;
    IMMDevice *m_pMMOutDevice;
    int m_iBufferMs;
    bool m_bSwapChannels;

    // set hr to S_FALSE to abort but return success
    CPrefs(int argc, LPCWSTR argv[], HRESULT &hr);
    ~CPrefs();

};

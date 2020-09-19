// log.h

#include <Windows.h>
#include <vector>
#include <cstdarg>

static inline void LOG(wchar_t* fmt...) {
    va_list args;
    size_t len;

    va_start(args, fmt);
    len = _vscwprintf(fmt, args);

    std::vector<wchar_t> buffer(len + 1);
    vswprintf_s(buffer.data(), len + 1, fmt, args);
    va_end(args);

    WriteConsoleW(GetStdHandle(STD_OUTPUT_HANDLE), buffer.data(), static_cast<DWORD>(len), nullptr, nullptr);
    WriteConsoleW(GetStdHandle(STD_OUTPUT_HANDLE), L"\n", 1, nullptr, nullptr);
}

#define ERR(...) LOG(L"Error: " __VA_ARGS__)

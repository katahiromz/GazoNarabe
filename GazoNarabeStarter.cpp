#include <windows.h>
#include <shlwapi.h>
#include <string>

INT WINAPI
WinMain(HINSTANCE   hInstance,
        HINSTANCE   hPrevInstance,
        LPSTR       lpCmdLine,
        INT         nCmdShow)
{
    WCHAR szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, MAX_PATH);
    PathRemoveFileSpecW(szPath);
    PathAppendW(szPath, L"data");
    std::wstring dir = szPath;
    PathAppendW(szPath, L"GazoNarabe.exe");

    INT argc, ret = -1;
    if (LPWSTR *wargv = CommandLineToArgvW(GetCommandLineW(), &argc))
    {
        std::wstring params;
        for (INT i = 1; i < argc; ++i)
        {
            if (i != 1)
                params += L" ";

            std::wstring arg = wargv[i];
            if (arg.find(' ') != std::wstring::npos ||
                arg.find('\t') != std::wstring::npos)
            {
                params += L"\"";
                params += arg;
                params += L"\"";
            }
            else
            {
                params += arg;
            }
        }

        INT_PTR inst = (INT_PTR)ShellExecuteW(NULL, NULL, szPath, params.c_str(), dir.c_str(), nCmdShow);
        if (inst > 32)
            ret = 0;

        LocalFree(wargv);
    }

    return ret;
}

#include <windows.h>

#include <cwctype>
#include <filesystem>
#include <string>

namespace fs = std::filesystem;

// 运行环境是否完整：python.exe、openpyxl、tkinter/Tcl/Tk 组件都在。
static bool IsRuntimeReady(const fs::path& runtimeDir)
{
    std::error_code ec;
    return fs::exists(runtimeDir / L"python.exe", ec)
        && fs::exists(runtimeDir / L"Lib" / L"site-packages" / L"openpyxl", ec)
        && fs::exists(runtimeDir / L"_tkinter.pyd", ec)
        && fs::exists(runtimeDir / L"tcl" / L"tcl8.6" / L"init.tcl", ec)
        && fs::exists(runtimeDir / L"tcl" / L"tk8.6" / L"tk.tcl", ec)
        && fs::exists(runtimeDir / L"tkinter" / L"__init__.py", ec);
}

// 取命令行里 exe 路径之后的参数，保持 Windows 原有的引号形式，原样转交给 fee_app.py。
static std::wstring GetArgumentsTail()
{
    const std::wstring cmdLine = GetCommandLineW();
    size_t i = 0;

    if (!cmdLine.empty() && cmdLine[0] == L'"') {
        ++i;
        while (i < cmdLine.size() && cmdLine[i] != L'"')
            ++i;
        if (i < cmdLine.size())
            ++i;
    } else {
        while (i < cmdLine.size() && !std::iswspace(cmdLine[i]))
            ++i;
    }

    while (i < cmdLine.size() && std::iswspace(cmdLine[i]))
        ++i;
    return cmdLine.substr(i);
}

static void ShowError(const wchar_t* message)
{
    MessageBoxW(NULL, message, L"饭费管理系统", MB_OK | MB_ICONERROR);
}

// 首次运行（或运行环境损坏）时调用 bootstrap_runtime.ps1。
// 传入 -NoLaunch：只安装环境，安装完成后由启动器自己拉起主程序，避免启动两次。
static bool RunBootstrap(const fs::path& appDir, const fs::path& script)
{
    std::wstring command = L"powershell.exe -NoProfile -ExecutionPolicy Bypass -File \""
                         + script.wstring() + L"\" -NoLaunch";
    const std::wstring workDir = appDir.wstring();

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi = {};
    if (!CreateProcessW(NULL, command.data(), NULL, NULL, FALSE, CREATE_NEW_CONSOLE,
                        NULL, workDir.c_str(), &si, &pi)) {
        ShowError(L"无法启动运行环境安装脚本。");
        return false;
    }

    WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD exitCode = 1;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return exitCode == 0;
}

int main()
{
    wchar_t exeBuffer[MAX_PATH] = {};
    if (!GetModuleFileNameW(NULL, exeBuffer, MAX_PATH)) {
        ShowError(L"无法获取程序路径。");
        return 1;
    }

    const fs::path appDir = fs::path(exeBuffer).parent_path();
    const fs::path runtimeDir = appDir / L"runtime";
    const fs::path pythonExe = runtimeDir / L"python.exe";
    const fs::path mainPy = appDir / L"fee_app.py";
    const fs::path bootstrap = appDir / L"bootstrap_runtime.ps1";

    if (!fs::exists(mainPy)) {
        ShowError(L"缺少主程序 fee_app.py，安装不完整。");
        return 1;
    }

    if (!IsRuntimeReady(runtimeDir)) {
        if (!fs::exists(bootstrap)) {
            ShowError(L"缺少 bootstrap_runtime.ps1，安装不完整。");
            return 1;
        }
        if (!RunBootstrap(appDir, bootstrap) || !IsRuntimeReady(runtimeDir)) {
            ShowError(L"运行环境安装失败，请检查网络连接后重试。\n也可以手动运行安装目录下的 bootstrap_runtime.bat。");
            return 1;
        }
    }

    // 嵌入式 Python 不含 Tcl/Tk，必须显式告诉 tkinter 去哪里找库文件。
    SetEnvironmentVariableW(L"TCL_LIBRARY", (runtimeDir / L"tcl" / L"tcl8.6").wstring().c_str());
    SetEnvironmentVariableW(L"TK_LIBRARY", (runtimeDir / L"tcl" / L"tk8.6").wstring().c_str());

    std::wstring command = L"\"" + pythonExe.wstring() + L"\" \"" + mainPy.wstring() + L"\"";
    const std::wstring tail = GetArgumentsTail();
    if (!tail.empty())
        command += L" " + tail;

    const std::wstring workDir = appDir.wstring();
    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi = {};
    if (!CreateProcessW(NULL, command.data(), NULL, NULL, FALSE, CREATE_NO_WINDOW,
                        NULL, workDir.c_str(), &si, &pi)) {
        ShowError(L"启动主程序失败。");
        return 1;
    }

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return 0;
}
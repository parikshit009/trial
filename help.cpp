/*************************************************************************
*
* ADOBE CONFIDENTIAL
* ___________________
*
* Copyright 2021 Adobe Systems Incorporated
* All Rights Reserved.
*
* NOTICE: All information contained herein is, and remains
* the property of Adobe Systems Incorporated and its suppliers,
* if any. The intellectual and technical concepts contained
* herein are proprietary to Adobe Systems Incorporated and its
* suppliers and may be covered by U.S. and Foreign Patents,
* patents in process, and are protected by trade secret or copyright law.
* Dissemination of this information or reproduction of this material
* is strictly forbidden unless prior written permission is obtained
* from Adobe Inc.
**************************************************************************/


#include "pch.h"
#include <wrl/module.h>
#include <wrl/implements.h>
#include <wrl/client.h>
#include <shobjidl_core.h>
#include <wil\resource.h>
#include <string>
#include <vector>
#include <sstream>
#include <Windows.h>
#include <tchar.h>
#include "resource.h"
#include "..\ContextMenuShim\Resources\resource.h"
#include <atlstr.h>
#include "UIHandler.h"
#include "AcroLaunchGuard.h"
#include <future>
#include <unordered_map>
#include "Logger.h"

#include <Windows.h>
#include <Shlwapi.h>
#include <shellapi.h>
#include "..\..\Distiller\Products\Adobe\Shared\Sources\TrayIconPublic.h"

#include "shobjidl_core.h"
#define ENABLE_LOCAL_DEBUGGING 0// enable this while debugging win11 context menu workflows locallty

#define REGKEY_POLICIES_BASEPATH	L"SOFTWARE\\Policies"
#define REGKEY_ADOBE_ACROBAT        L"\\Adobe\\Adobe Acrobat\\"
#define REGKEY_FEATURE_LOCK_DOWN    L"\\FeatureLockDown\\"
#define ACROBAT_ENTITLEMENT_KEY     L"SOFTWARE\\Adobe\\Adobe Acrobat\\" REGISTRY_KEY_TRACK L"\\AVEntitlement"

#define IsAcroTrayEnabledAsService_KEY		L"bIsAcroTrayEnabledAsService"

typedef enum {
    PROCESS_TYPE_ACTIVATION_ACROBAT,
    PROCESS_TYPE_ACTIVATION_ACROTRAY,
    PROCESS_TYPE_ACTIVATION_ACROELEM,
    PROCESS_TYPE_ACTIVATION_END,
} PROCESS_TYPE_ACTIVATION;

DWORD LaunchProcessViaActivationRoute(PROCESS_TYPE_ACTIVATION proctye, const WCHAR* params)
{
    IApplicationActivationManager* pActivator = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_ApplicationActivationManager,
        nullptr,
        CLSCTX_INPROC,
        IID_IApplicationActivationManager,
        (void**)&pActivator);

    if (!SUCCEEDED(hr)) {
        return false;
    }

#if (IS_DEV_BUILD || ENABLE_LOCAL_DEBUGGING)
    std::wstring appId = L"AdobeAcrobatDCCoreApp_9p443f96enen8";
#else
    std::wstring appId = L"AdobeAcrobatDCCoreApp_pc75e8sa7ep4e";
#endif
    switch (proctye) {
    case PROCESS_TYPE_ACTIVATION::PROCESS_TYPE_ACTIVATION_ACROBAT:
        appId = appId + L"!AdobeAcrobatDCCoreAppAcr";
        break;
    case PROCESS_TYPE_ACTIVATION::PROCESS_TYPE_ACTIVATION_ACROTRAY:
        appId = appId + L"!AdobeAcrobatDCCoreAppAcroTray";
        break;
    case PROCESS_TYPE_ACTIVATION::PROCESS_TYPE_ACTIVATION_ACROELEM:
        appId = appId + L"!AdobeAcrobatDCCoreAppAcroElem";
        break;
    }

    DWORD pid = 0;
    hr = pActivator->ActivateApplication(appId.c_str(), params, AO_NONE, &pid);
    if (!SUCCEEDED(hr)) {
        pActivator->Release();
        return 0;
    }

    pActivator->Release();

    return pid;
}

#include <VersionHelpers.h>
#include <winnt.h>
typedef LONG NTSTATUS;
NTSTATUS RtlGetVersionFunctionProto(_Out_ PRTL_OSVERSIONINFOW lpVersionInformation);
using RtlGetVersionFunction = decltype(RtlGetVersionFunctionProto);

bool IsWindows11OrGreater()
{
    if (RtlGetVersionFunction* RtlGetVersion_f = reinterpret_cast<RtlGetVersionFunction*>(GetProcAddress(::GetModuleHandle(L"ntdll.dll"), "RtlGetVersion")))
    {
        RTL_OSVERSIONINFOW vi = {};
        vi.dwOSVersionInfoSize = sizeof(vi);
        if ((RtlGetVersion_f(&vi)) >= 0)
        {
            if (vi.dwMajorVersion >= 10 && vi.dwBuildNumber >= 22000)
                return true;

        }
    }
    return false;
}


const UINT PRINT_DELAY = 2000;

// Identifies the "Create Adobe PDF" menu item.
const UINT CREATE_ADOBE_PDF = 0;
// Identifies the "Create and Email Adobe PDF" menu item.
const UINT CREATE_AND_EMAIL_ADOBE_PDF = 1;
// Identifies the "Change Default Conversion Settings..." menu item.
//const UINT CHANGE_CONVERSION_SETTINGS = 2;
// Identifies the "Add to Binder..." menu item.
const UINT ADD_TO_ACROBAT_BINDER = 2;
// Identifies the "Edit to Binder..." menu item.
const UINT EDIT_PDF_IN_ACROBAT = 3;

const UINT SHARE_PDF_FOR_REVIEW = 4;
// Indicates the number of commands.
const USHORT COMMAND_COUNT = 5;

#define READ_BUFFER_SIZE	0xFFF0 //64K buffer size
#define NON_OFFICE_FILE		0
#define WORD_FILE			1
#define EXCEL_FILE			2
#define POWER_POINT_FILE	3
#define VISIO_FILE			4
#define AUTOCAD_FILE		5
#define PROJECT_FILE		6
#define PUBLISHER_FILE		7
#define ACCESS_FILE			8
#define INVENTOR_FILE		9
#define ACROBAT_NAME_STR		_T(REGISTRY_KEY_NAMEA)

enum theme {
    LightTheme,
    DarkTheme
};
enum themeIconLoaded
{
    noIcons,
    lightIcons,
    darkIcons
};


void WINAPI TimerprocGetState(
    HWND unnamedParam1,
    UINT unnamedParam2,
    UINT_PTR unnamedParam3,
    DWORD unnamedParam4
);

/*AVGetOSAppTheme() returns the active OS theme for Apps. Context menu theme is same as OS's App theme settings*/
static theme GetOSAppTheme()
{
    theme os_theme = LightTheme;

    HKEY hKey;
    LONG lRes = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, KEY_READ, &hKey);
    if (lRes == ERROR_SUCCESS) {
        DWORD dwType = 0;
        DWORD dwValue = 0;
        DWORD dwSize = sizeof(DWORD);
        if (RegQueryValueEx(hKey, L"AppsUseLightTheme", NULL, &dwType, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS) {
            if (dwValue == 0)
                os_theme = DarkTheme;
        }
        RegCloseKey(hKey);
    }
    return os_theme;
}

extern "C" BOOL IsThisSupportedFileObjects(TCHAR * pBuffer);

extern "C" int GetOfficeFormat(const TCHAR * pBuffer);

extern "C"  BOOL IsThisAdobePDF(LPCTSTR pExt);

extern "C" BOOL IsThisImageFileSupported(LPCTSTR pExt);

static HINSTANCE LoadAdist32Dll(void)
{
    HINSTANCE hInst; //Adist.dll  handle    
    TCHAR szAdistPath[MAX_PATH] = { 0 };
    HKEY hkey;

    if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_DISTILLER_KEY,
        &hkey) == ERROR_SUCCESS)
    {
        DWORD cbValue = sizeof(szAdistPath);
        DWORD typeData = REG_SZ;
        RegQueryValueEx(hkey, L"InstallPath", NULL,
            &typeData, (LPBYTE)szAdistPath, &cbValue);
        RegCloseKey(hkey);
    }

#ifdef _WIN64
    wcscat(szAdistPath, L"\\ADIST64.dll");
#else
    wcscat(szAdistPath, L"\\ADIST.dll");
#endif 
    hInst = LoadLibrary(szAdistPath);
    return hInst;
}

static BOOL IsAcrobatElementsOnlyInstalled()
{
    BOOL bRetVal = TRUE;
    HINSTANCE hInst = LoadAdist32Dll();
    if (hInst)
    {
        typedef BOOL(*PIS_ELEMENTS_INSTALLED)(void);
        PIS_ELEMENTS_INSTALLED pIsElements;
        if (pIsElements = (PIS_ELEMENTS_INSTALLED)GetProcAddress(hInst, (LPCSTR)"IsAcrobatElementsInstalled"))
        {
            bRetVal = pIsElements();
        }
        FreeLibrary(hInst);
    }
    return bRetVal;
}

static bool Is64BitAcrobatInstalled()
{
    static bool sInitialized = false;
    static bool is64Bit = false;
    if (!sInitialized)
    {
        static LPCTSTR kAcrobatInstallerKey = _T("Software\\Adobe\\Adobe Acrobat\\") ACROBAT_NAME_STR _T("\\Installer");
        static LPCTSTR kAcrobat64BitKey = _T("Is64BitProduct");
        sInitialized = true;
        HKEY hKey = nullptr;
        LONG lResult = 0;
        lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, kAcrobatInstallerKey, 0, KEY_QUERY_VALUE, &hKey);

        if (lResult == ERROR_SUCCESS)
        {
            DWORD		keySize = sizeof(DWORD);
            DWORD		keyType = REG_DWORD;
            DWORD		keyValue = 0;
            lResult = RegQueryValueEx(hKey, kAcrobat64BitKey, nullptr, &keyType, (LPBYTE)&keyValue, &keySize);
            if ((ERROR_SUCCESS == lResult) && (REG_DWORD == keyType))
            {
                is64Bit = (keyValue == 1);
            }
            RegCloseKey(hKey);
        }
    }
    return is64Bit;
}

static bool  CheckIfAutoCADIsInstalled(void) {
    static bool sInitialized = FALSE;
    static bool sAutoCADInstalled = FALSE;

    if (sInitialized == FALSE) {
        sInitialized = TRUE;

        LPCTSTR autocadReg = L"SOFTWARE\\Autodesk\\AutoCAD\\R";
        LPCTSTR versionStrs[] = { L"23.0", L"22.0", L"20.0", L"19.1", L"19.0" };

        for (int i = 0; i < sizeof(versionStrs) / sizeof(LPCTSTR); ++i) {
            LPCTSTR versionStr = versionStrs[i];
            wchar_t fullReg[_MAX_PATH + 1] = { 0 };
            if (StringCchCat(fullReg, MAX_PATH + 1, autocadReg) == S_OK && StringCchCat(fullReg, MAX_PATH + 1, versionStr) == S_OK) {
                HKEY hKey = nullptr;
                LONG lResult = 0;
                lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, fullReg, 0, KEY_QUERY_VALUE, &hKey);
                if (lResult == ERROR_SUCCESS) {
                    RegCloseKey(hKey);
                    sAutoCADInstalled = TRUE;
                    break;
                }
            }
        }
    }
    return sAutoCADInstalled;
}

extern "C" static LRESULT SendTrayIconMessage(UINT code, void* pFile)
{
    HWND hwndTrayIcon;
    COPYDATASTRUCT cds;

    hwndTrayIcon = FindWindow(g_szTrayIconClass, g_szTrayIconClass);
    if (!hwndTrayIcon)
    {
        bool areWeLoadedInWin11Menu = ::GetModuleHandleW(L"ContextMenuIExplorerCommandShim.dll") != 0;
        bool success = false;
        if (areWeLoadedInWin11Menu)
        {
            success = LaunchProcessViaActivationRoute(PROCESS_TYPE_ACTIVATION_ACROTRAY, 0);
            if (success)
                Sleep(1000);
        }
        if (!success)
        {
            HKEY hKey;
            if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, DISTILLER_PATH_KEY, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
            {
                TCHAR szPath[MAX_PATH];
                DWORD lData = sizeof(szPath);
                DWORD lType;
                if (RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szPath, &lData) == ERROR_SUCCESS)
                {
                    wcscat(szPath, _T("\\AcroTray.exe"));
                    SHELLEXECUTEINFO shExecInfo = { 0 };
                    shExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
                    shExecInfo.hwnd = HWND_DESKTOP;
                    shExecInfo.lpFile = szPath;
                    shExecInfo.nShow = SW_SHOWNORMAL;
                    shExecInfo.fMask = SEE_MASK_FLAG_DDEWAIT;
                    ShellExecuteEx(&shExecInfo);
                    Sleep(1000);
                }
                RegCloseKey(hKey);
            }
        }


        hwndTrayIcon = FindWindow(g_szTrayIconClass, g_szTrayIconClass);
        if (!hwndTrayIcon)
        {
            int iCount = 0;
            while (iCount++ < 25 && (hwndTrayIcon == NULL))
            {
                Sleep(400); // Citrix presentation server takes more time to initialize let us wait for 10,000 Milliseconds
                hwndTrayIcon = FindWindow(g_szTrayIconClass, g_szTrayIconClass);
            }

            if (!hwndTrayIcon)
                return 0;
        }
    }

    cds.dwData = code;
    cds.lpData = pFile;
    if (code == TI_HLIGHT_LOG_EVENT)
        cds.cbData = sizeof(HEADLIGHT_DATA);
    else
        cds.cbData = (!pFile ? 0 : (DWORD)((wcslen((TCHAR*)pFile) + 1) * sizeof(TCHAR)));

    return SendMessage(hwndTrayIcon, WM_COPYDATA,
        (WPARAM)GetForegroundWindow(), (LPARAM)&cds);
}


#include "Msi.h"

static void GetProductWithUpgradeCode(LPCTSTR sUpgradeCode, LPTSTR sProdCodeList)
{
    DWORD dw = ERROR_SUCCESS;
    for (unsigned int index = 0; dw == ERROR_SUCCESS; ++index)
    {
        TCHAR sProductCode[MAX_PATH] = { 0 };
        dw = MsiEnumRelatedProducts(sUpgradeCode, 0, index, sProductCode);
        if (dw == ERROR_SUCCESS && lstrlen(sProductCode))
        {
            if (_tcslen(sProdCodeList) > 0)
                _tcscat_s(sProdCodeList, MAX_PATH, L";");

            _tcscat_s(sProdCodeList, MAX_PATH, sProductCode);
        }
    }
}

static BOOL IsSCAReader(LPCTSTR sProductCode)
{
    BOOL bRet = false;
    if (sProductCode[25] == 'B')
    {
        if (_tcsstr(sProductCode, L"-1033-0000-") || _tcsstr(sProductCode, L"-1033-FFFF-"))
            bRet = false;
        else
            bRet = true;
    }
    return bRet;
}

static bool IsAcrobatInstalled()
{
    bool isAcrobatInstalled = false;

    //Below are Acrobat and Reader upgrade codes
    //{A6EADE66 - 0000 - 0000 - 484E-7E8A45000000} -UpgradeCode for Reader pre-SCA
    //{AC76BA86 - 0000 - 0000 - 7760 - 7E8A45000000} -UpgradeCode for Acrobat
    TCHAR sUpgradeCode[MAX_PATH] = { _T("{AC76BA86-0000-0000-7760-7E8A45000000}") };
    TCHAR sProductList[MAX_PATH] = { _T("") };
    GetProductWithUpgradeCode(sUpgradeCode, sProductList); //sProductList is ; seperated list of codes.

    if (_tcslen(sProductList) > 0)
    {
        // Acrobat is installed but it could be Reader SCA
        TCHAR* token_list_temp = NULL;

        LPTSTR next_token = _tcstok_s(sProductList, _T(";"), &token_list_temp);
        while (next_token != NULL)
        {
            if (!IsSCAReader(next_token))
            {
                isAcrobatInstalled = true;
                break;
            }
            next_token = _tcstok_s(NULL, _T(";"), &token_list_temp);
        }
    }
    return isAcrobatInstalled;
}

// Check if acrotray not running as services feature is disabled.
// By default it is ON
bool IsAcroTrayEnabledAsService()
{
    static bool b_Once = false;
    static bool bIsAcroTrayEnabledAsService = true;
    static bool bKeyFound = false;

    if (!b_Once)
    {
        DWORD keySize = sizeof(DWORD);
        DWORD dwValue = 0;
        HKEY hKeyCU = NULL;
        bKeyFound = false;

        // Read bIsAcroTrayEnabledAsService from registry, set by Acrobat.
        if ((::RegOpenKeyEx(HKEY_CURRENT_USER, ACROBAT_ENTITLEMENT_KEY, 0, KEY_QUERY_VALUE, &hKeyCU) == ERROR_SUCCESS))
        {
            if ((ERROR_SUCCESS == ::RegQueryValueEx(hKeyCU, IsAcroTrayEnabledAsService_KEY, NULL, NULL, (LPBYTE)&dwValue, &keySize)))
            {
                if (dwValue == 0)
                    bIsAcroTrayEnabledAsService = false;
            }
            ::RegCloseKey(hKeyCU);
        }

        HKEY hKey = NULL;
        std::wstring regFeatureLockDownPath = REGKEY_POLICIES_BASEPATH REGKEY_ADOBE_ACROBAT REGISTRY_KEY_TRACKW REGKEY_FEATURE_LOCK_DOWN;

        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, regFeatureLockDownPath.c_str(), 0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            if (hKey)
            {
                std::wstring isAcroTrayEnabledAsServiceKeyName = IsAcroTrayEnabledAsService_KEY;
                DWORD dwType = REG_NONE;
                dwValue = 0;
                keySize = sizeof(DWORD);

                if (RegQueryValueExW(hKey, isAcroTrayEnabledAsServiceKeyName.c_str(), NULL, &dwType, (LPBYTE)&dwValue, &keySize) == ERROR_SUCCESS)
                {
                    if (dwType == REG_DWORD)
                    {
                        bKeyFound = true;
                        bIsAcroTrayEnabledAsService = (dwValue != 0);
                    }
                }
                RegCloseKey(hKey);
            }
        }

        b_Once = true;
    }

#if IS_TRUNK_BUILD || IS_BETA_BUILD
    if (!bKeyFound)
        return true;
#endif
    return bIsAcroTrayEnabledAsService;
}

static BOOL IsAcrobatStdInstalled()
{
    if (!IsAcroTrayEnabledAsService())
    {
        static BOOL bEntitlementStd = FALSE;

        if (!bEntitlementStd)
        	bEntitlementStd = BOOL(SendTrayIconMessage(TI_CHECK_ACROSTD, NULL));
        return bEntitlementStd;
    }
    else
    {
        static BOOL bIsAcrobatStdInstalled = FALSE;

        if (!bIsAcrobatStdInstalled)
            bIsAcrobatStdInstalled = IsAcrobatInstalled();

        return bIsAcrobatStdInstalled;
    }
}

static BOOL IsAcrobatProInstalled()
{
    if (!IsAcroTrayEnabledAsService())
    {
        static BOOL bEntitlementPro = FALSE;

        if (!bEntitlementPro)
        	bEntitlementPro = BOOL(SendTrayIconMessage(TI_CHECK_ACROPRO, NULL));
        return bEntitlementPro;
    }
    else
    {
        static BOOL bIsAcrobatProInstalled = FALSE;

        if (!bIsAcrobatProInstalled)
            bIsAcrobatProInstalled = IsAcrobatInstalled();

        return bIsAcrobatProInstalled;
    }
}


static BOOL IsFormatSupportedWithStandard(int formatType)
{
    return (formatType != AUTOCAD_FILE && formatType != VISIO_FILE && formatType != PROJECT_FILE);
}

//To fix bug# 2833949, need to call this for those office applications
//for which PDF Maker is yet not supported in 64 bit. 
static bool CheckIfApplication64BitBinary(HKEY hKey)
{
    bool result64BitBinary = false;
    if (hKey)
    {
        TCHAR szValue[MAX_PATH];
        DWORD lData = sizeof(szValue);
        DWORD lType;
        DWORD lBinaryType = SCS_32BIT_BINARY;
        BOOL result = 0;
        if (RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData) == ERROR_SUCCESS)
            result = GetBinaryTypeW(szValue, &lBinaryType);

        if (result != 0 && lBinaryType == SCS_64BIT_BINARY)
            result64BitBinary = true;
    }

    return result64BitBinary;
}

BOOL CheckIfApplicationInstalled(int iOfficeFile)
{
    LPCTSTR pExt, pApp;
    BOOL retVal = FALSE;
    CString szAppPath('\0', 2 * MAX_PATH);

    switch (iOfficeFile)
    {
    case AUTOCAD_FILE:
        HKEY hKey;
        pExt = _T(".dwg");
        pApp = _T("\\ACAD.EXE");
        if (RegOpenKey(HKEY_CLASSES_ROOT, pExt, &hKey) == ERROR_SUCCESS)
        {
            DWORD lData = 2 * MAX_PATH;
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)(LPTSTR)(LPCTSTR)szAppPath, &lData);
            RegCloseKey(hKey);
            szAppPath.Format(_T("%s%s"), szAppPath, _T("\\shell\\open\\command"));
            if (RegOpenKey(HKEY_CLASSES_ROOT, (LPTSTR)(LPCTSTR)szAppPath, &hKey) == ERROR_SUCCESS)
            {
                lData = 2 * MAX_PATH;
                RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)(LPTSTR)(LPCTSTR)szAppPath, &lData);
                RegCloseKey(hKey);
            }
            else
                return FALSE;
        }
        else
            return FALSE;
        break;
    default:
        return FALSE;
    }

    CString szExe('\0', 2 * MAX_PATH);
    DWORD cchOut = 2 * MAX_PATH;
    if (AssocQueryString(ASSOCF_NOUSERSETTINGS,
        ASSOCSTR_COMMAND,
        pExt,
        NULL,
        (LPTSTR)(LPCTSTR)szExe,
        &cchOut) == S_OK)
    {
        szExe.MakeUpper();
        if (szExe.Find(pApp) != -1)
            retVal = TRUE;
    }
    else
    {
        szAppPath.Format(_T("%s"), szAppPath);
        szAppPath.MakeUpper();
        if (szAppPath.Find(pApp) != -1)
            retVal = TRUE;
    }
    return retVal;
}


static BOOL CheckIfPFDMakerIsInstalled(int iOfficeFile)
{
    static int iInstalledOfficeMacros = -1;

    if (iInstalledOfficeMacros == -1)
    {
        iInstalledOfficeMacros = 0;
        HKEY hKey;
        // Check if PDFMaker is installed
        if (RegOpenKey(HKEY_CLASSES_ROOT, _T("AdobePDFMakerX.Word"), &hKey) == ERROR_SUCCESS)
        {
            RegCloseKey(hKey);
            //PDFMaker is installed now check if Word Macro is installed
            if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_WORD_PATH, &hKey) == ERROR_SUCCESS
#ifdef _WIN64
                || RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_WORD_PATH_64, &hKey) == ERROR_SUCCESS
#else
                // Access a 64-bit key from a 32-bit application
                || RegOpenKeyEx(HKEY_LOCAL_MACHINE, LM_MS_WORD_PATH_64, 0, KEY_READ | KEY_WOW64_64KEY, &hKey) == ERROR_SUCCESS
#endif
                )
            {
                RegCloseKey(hKey);
                if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_WORD_PATH, &hKey) == ERROR_SUCCESS)
                {
                    DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                    RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                    if (dwValue)
                        iInstalledOfficeMacros = WORD_FILE;
                    RegCloseKey(hKey);
                }
            }
            //Now check if Excel Macro is installed
            if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_EXCEL_PATH, &hKey) == ERROR_SUCCESS
#ifdef _WIN64
                || RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_EXCEL_PATH_64, &hKey) == ERROR_SUCCESS
#else
                // Access a 64-bit key from a 32-bit application
                || RegOpenKeyEx(HKEY_LOCAL_MACHINE, LM_MS_EXCEL_PATH_64, 0, KEY_READ | KEY_WOW64_64KEY, &hKey) == ERROR_SUCCESS
#endif
                )
            {
                RegCloseKey(hKey);
                if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_EXCEL_PATH, &hKey) == ERROR_SUCCESS)
                {
                    DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                    RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                    if (dwValue)
                        iInstalledOfficeMacros |= (EXCEL_FILE << 0x4);
                    RegCloseKey(hKey);
                }
            }
            //Now check if Power Point is installed
            if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_POWERPOINT_PATH, &hKey) == ERROR_SUCCESS
#ifdef _WIN64
                || RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_POWERPOINT_PATH_64, &hKey) == ERROR_SUCCESS
#else
                // Access a 64-bit key from a 32-bit application
                || RegOpenKeyEx(HKEY_LOCAL_MACHINE, LM_MS_POWERPOINT_PATH_64, 0, KEY_READ | KEY_WOW64_64KEY, &hKey) == ERROR_SUCCESS
#endif
                )
            {
                RegCloseKey(hKey);
                if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_POWERPOINT_PATH, &hKey) == ERROR_SUCCESS)
                {
                    DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                    RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                    if (dwValue)
                        iInstalledOfficeMacros |= (POWER_POINT_FILE << 0x8);
                    RegCloseKey(hKey);
                }
            }
            //Now check if Visio is installed
            if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_VISIO_PATH, &hKey) == ERROR_SUCCESS || // Visio 2002
                RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_VISIO32_PATH, &hKey) == ERROR_SUCCESS	// Visio 2000
#ifdef _WIN64
                || RegOpenKey(HKEY_LOCAL_MACHINE, LM_MS_VISIO_PATH_64, &hKey) == ERROR_SUCCESS
#endif
                )
            {
                RegCloseKey(hKey);
                if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_VISIO_PATH, &hKey) == ERROR_SUCCESS)
                {
                    DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                    RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                    if (dwValue)
                        iInstalledOfficeMacros |= (VISIO_FILE << 0xc);
                    RegCloseKey(hKey);
                }
            }
            //Now check if AUTOCAD is installed
            hKey = 0;
            if (IsAcrobatElementsOnlyInstalled() == false)
                iInstalledOfficeMacros |= (AUTOCAD_FILE << 0x10);
            else if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_AUTOCAD_PATH, &hKey) == ERROR_SUCCESS ||
                CheckIfApplicationInstalled(AUTOCAD_FILE))

            {
                bool is64BitBinary = CheckIfApplication64BitBinary(hKey);
                if (hKey)
                    RegCloseKey(hKey);

                if (!is64BitBinary)
                {
                    if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_AUTOCAD_PATH, &hKey) == ERROR_SUCCESS)
                    {
                        DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                        RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                        if (dwValue)
                            iInstalledOfficeMacros |= (AUTOCAD_FILE << 0x10);
                        RegCloseKey(hKey);
                    }
                }
            }
            //Now check if MSProject is installed
            if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PROJECT_PATH, &hKey) == ERROR_SUCCESS)
            {
                bool is64BitBinary = CheckIfApplication64BitBinary(hKey);
                RegCloseKey(hKey);
                if (!is64BitBinary)
                {
                    if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_PROJECT_PATH, &hKey) == ERROR_SUCCESS)
                    {
                        DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                        RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                        if (dwValue)
                            iInstalledOfficeMacros |= (PROJECT_FILE << 0x14);
                        RegCloseKey(hKey);
                    }
                }
            }
            //Now check if MSPublisher is installed
            if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PUBLISHER_PATH, &hKey) == ERROR_SUCCESS)
            {
                bool is64BitBinary = CheckIfApplication64BitBinary(hKey);
                RegCloseKey(hKey);
                if (!is64BitBinary)
                {
                    if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_PUBLISHER_PATH, &hKey) == ERROR_SUCCESS)
                    {
                        DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                        RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                        if (dwValue)
                            iInstalledOfficeMacros |= (PUBLISHER_FILE << 0x18);
                        RegCloseKey(hKey);
                    }
                }
            }
            //Now check if MSAccess is installed
            if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_ACCESS_PATH, &hKey) == ERROR_SUCCESS)
            {
                bool is64BitBinary = CheckIfApplication64BitBinary(hKey);
                RegCloseKey(hKey);
                if (!is64BitBinary)
                {
                    if (RegOpenKey(HKEY_LOCAL_MACHINE, LM_PDFMAKER_FOR_ACCESS_PATH, &hKey) == ERROR_SUCCESS)
                    {
                        DWORD dwValue = 0, cbData = sizeof(DWORD), dwType = 0;
                        RegQueryValueEx(hKey, (LPCTSTR)NULL, 0, &dwType, (LPBYTE)&dwValue, &cbData);
                        if (dwValue)
                            iInstalledOfficeMacros |= (ACCESS_FILE << 0x1c);
                        RegCloseKey(hKey);
                    }
                }
            }
        }
    }
    switch (iOfficeFile)
    {
    case WORD_FILE:
        if (iInstalledOfficeMacros & WORD_FILE)
            return TRUE;
        break;
    case EXCEL_FILE:
        if (iInstalledOfficeMacros & (EXCEL_FILE << 4))
            return TRUE;
        break;

    case POWER_POINT_FILE:
        if (iInstalledOfficeMacros & (POWER_POINT_FILE << 8))
            return TRUE;
        break;
    case VISIO_FILE:
        if (iInstalledOfficeMacros & (VISIO_FILE << 0xc))
            return TRUE;
        break;
    case AUTOCAD_FILE:
        if (iInstalledOfficeMacros & (AUTOCAD_FILE << 0x10))
            return TRUE;
        break;

    case PROJECT_FILE:
        if (iInstalledOfficeMacros & (PROJECT_FILE << 0x14))
            return TRUE;
        break;
    case PUBLISHER_FILE:
        if (iInstalledOfficeMacros & (PUBLISHER_FILE << 0x18))
            return TRUE;
        break;
    case ACCESS_FILE:
        if (iInstalledOfficeMacros & (ACCESS_FILE << 0x1c))
            return TRUE;
        break;
    }
    return FALSE;
}

static BOOL IsThisSupportedOfficeFormat(int format)
{
    if (!IsAcrobatProInstalled() && !IsFormatSupportedWithStandard(format))
        return FALSE;

    if (format == WORD_FILE || format == EXCEL_FILE || format == POWER_POINT_FILE || format == VISIO_FILE || format == PROJECT_FILE)
        return TRUE;

    if (format != NON_OFFICE_FILE)
        return CheckIfPFDMakerIsInstalled(format);
    return FALSE;
}

static BOOL IsThisSupportedCombineOfficeFormat(int format)
{
    if (!IsAcrobatProInstalled() && !IsFormatSupportedWithStandard(format))
        return false;

    if (format != NON_OFFICE_FILE)
        return CheckIfPFDMakerIsInstalled(format);
    return FALSE;
}

static int GetFeatureEnableThreshold(LPCWSTR featureReg)
{
    int threshold = 500;
    HKEY hKey;
    LONG lRes = ERROR_PATH_NOT_FOUND;
#if _WIN64
    if (Is64BitAcrobatInstalled())
        lRes = RegOpenKeyExW(HKEY_LOCAL_MACHINE, LM_SHARE_ENABLE_THRESHOLD_64BIT, 0, KEY_READ, &hKey);
    else
#endif
        lRes = RegOpenKeyExW(HKEY_LOCAL_MACHINE, LM_SHARE_ENABLE_THRESHOLD, 0, KEY_READ, &hKey);
    if (lRes == ERROR_SUCCESS) {
        DWORD dwType = 0;
        DWORD dwValue = 0;
        DWORD dwSize = sizeof(DWORD);
        if (RegQueryValueEx(hKey, featureReg, NULL, &dwType, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS) {
            threshold = (int)dwValue;
        }
        RegCloseKey(hKey);
    }
    return threshold;
}

static BOOL EnableFeature(int threshold)
{
    BOOL enable = FALSE;
    HKEY hKey;
    LONG keyRes = ERROR_PATH_NOT_FOUND;
#if _WIN64
    if (Is64BitAcrobatInstalled())
        keyRes = RegOpenKeyExW(HKEY_LOCAL_MACHINE, LM_IRANDOM_KEY_64BIT, 0, KEY_READ, &hKey);
    else
#endif
        keyRes = RegOpenKeyExW(HKEY_LOCAL_MACHINE, LM_IRANDOM_KEY, 0, KEY_READ, &hKey);
    if (keyRes == ERROR_SUCCESS)
    {
        DWORD dwType = 0;
        DWORD dwValue = 0;
        DWORD dwSize = sizeof(DWORD);
        if (RegQueryValueEx(hKey, L"irandom", NULL, &dwType, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS)
        {
            if (dwValue <= threshold)
                enable = TRUE;
        }
        RegCloseKey(hKey);
    }
    return enable;
}
using namespace Microsoft::WRL;

BOOL APIENTRY DllMain(HMODULE hModule,
    DWORD  ul_reason_for_call,
    LPVOID lpReserved
)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}


#define ENABLE_LOGGING 1

#if ENABLE_LOGGING
class MeasureBlockTime
{
public:
    MeasureBlockTime(std::string stringmsg) :m_stringmsg(stringmsg)
    {
        start = std::chrono::high_resolution_clock::now();
     }
    ~MeasureBlockTime()
    {
        auto elapsed = std::chrono::high_resolution_clock::now() - start;
        long long microseconds = std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count();
        OutputDebugStringA((std::string("CTIexplore Log:") + m_stringmsg + std::string("\t time taken:") + std::to_string(microseconds)).c_str());
    }
    std::chrono::high_resolution_clock::time_point start;
    std::string m_stringmsg;
};
#endif

#if ENABLE_LOGGING 
#define LOG_TIME_TAKEN(msg) MeasureBlockTime tmp = MeasureBlockTime(std::string(msg));
#else
#define LOG_TIME_TAKEN(msg)
#endif




void convertVectorToArray(const std::vector<std::wstring>& v, TCHAR** fileList) {
    size_t len = v.size();
    for (size_t i = 0; i < len; i++) {
        size_t slen = v[i].length() + 1;
        fileList[i] = new (std::nothrow) TCHAR[slen];
        if (fileList[i] && v[i].length())
        {
            _tcscpy_s(fileList[i], slen, v[i].c_str());
        }
    }
}

TCHAR** allocArrayFromVector(const std::vector<std::wstring>& v) {
    size_t len = v.size();
    if (len) {
        TCHAR** fileList = new (std::nothrow)TCHAR * [len];
        if (fileList) {
            for (size_t i = 0; i < len; i++) {
                int slen = v[i].length() + 1;
                fileList[i] = new (std::nothrow) TCHAR[slen];
                if (fileList[i]) {
                    _tcscpy_s(fileList[i], slen, v[i].c_str());
                }
            }
        }
        return fileList;
    }
    return NULL;
}

void deallocArray(TCHAR** fileList, unsigned int len) {
    for (unsigned int i = 0; i < len; i++) {
        delete[] fileList[i];
    }
    delete[] fileList;
}


class AcroExplorerCommandBaseHelper 
{
    AcroExplorerCommandBaseHelper() : 
        m_shim64_module(nullptr),
        m_timer(0),
        m_commandId(-1)
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);

        CRegKey cReg;
        if (cReg.Open(HKEY_LOCAL_MACHINE, DISTILLER_PATH_KEY, KEY_READ) == ERROR_SUCCESS)
        {
            WCHAR contextmenu_folder_path[MAX_PATH + 1] = {};
            ULONG lValue = sizeof(contextmenu_folder_path);
            cReg.QueryStringValue(NULL, contextmenu_folder_path, &lValue);
            wcscat(contextmenu_folder_path, L"\\..\\Acrobat Elements\\");
            m_shim64_iexplore_dll_path = m_contextmenu64_dll_path = m_shim64_dll_path = contextmenu_folder_path;

            m_contextmenu64_dll_path.append(CONTEXT_MENU_DLL);
            m_shim64_dll_path.append(CONTEXT_MENU_SHIM_DLL);
            m_shim64_iexplore_dll_path.append(CONTEXT_MENU_SHIM_IEXPLORE_DLL);

        }
        m_shim64_module = ::LoadLibraryEx(m_shim64_dll_path.c_str(), 0, 0);

        m_hresource = GetResourceDll(m_shim64_module);
    }
    ~AcroExplorerCommandBaseHelper()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);

       if (m_shim64_module)
            FreeLibrary(m_shim64_module);
        m_shim64_module = nullptr;
    }

    HINSTANCE LoadContextMenuDll()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);
        
        HINSTANCE hinst = ::LoadLibraryEx(m_contextmenu64_dll_path.c_str(), 0, 0);
        if (hinst)
            return hinst;

        CString szTitle, szError;
        LoadString(m_hresource, IDS_REBOOT_PROMPT_TITLE, szTitle.GetBuffer(MAX_PATH), MAX_PATH);
        LoadString(m_hresource, IDS_REBOOT_PROMPT, szError.GetBuffer(MAX_PATH), MAX_PATH);
        MessageBox(GetForegroundWindow(), szError, szTitle, MB_ICONEXCLAMATION | MB_OK);

        return NULL;
    }
public:


    static AcroExplorerCommandBaseHelper& GetInstance()
    {
        static AcroExplorerCommandBaseHelper instance;
        return instance;
   }

   
    void ConvertToPDF()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);

        typedef void (*EntryPointfuncPtr)(TCHAR**, unsigned int);
        if (HINSTANCE hContextMenuDll = LoadContextMenuDll())
        {
            EntryPointfuncPtr funcProc = (EntryPointfuncPtr)GetProcAddress(hContextMenuDll, "ConvertToPDF");
            if (funcProc) {
                unsigned int len = static_cast<unsigned int>(m_filelist.size());
                TCHAR** fileList = allocArrayFromVector(m_filelist);
                if (fileList) {
                    funcProc(fileList, len);
                    deallocArray(fileList, len);
                }
            }
            FreeLibrary(hContextMenuDll);
        }
    }

    void ConvertAndSharePDF()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__ );
        LOG_TIME_TAKEN(msg_log_timer);
        typedef void (*EntryPointfuncPtr)(TCHAR**, unsigned int, DWORD, DWORD);
        if (HINSTANCE hContextMenuDll = LoadContextMenuDll())
        {
            EntryPointfuncPtr funcProc = (EntryPointfuncPtr)GetProcAddress(hContextMenuDll, "ConvertAndSharePDF");
            if (funcProc) {
                unsigned int len = static_cast<unsigned int>(m_filelist.size());
                TCHAR** fileList = allocArrayFromVector(m_filelist);
                if (fileList) {
                    DWORD bView =0 ;
                    DWORD bPrompt =0;
                    funcProc(fileList, len, bView, bPrompt);
                    deallocArray(fileList, len);
                }
            }

            FreeLibrary(hContextMenuDll);
        }
    }

    void ChangeConversionSettings()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);
        typedef void (*EntryPointfuncPtr)();
        if (HINSTANCE hContextMenuDll = LoadContextMenuDll())
        {
            EntryPointfuncPtr funcProc = (EntryPointfuncPtr)GetProcAddress(hContextMenuDll, "ChangeConversionSettings");
            if (funcProc)
                funcProc();
            FreeLibrary(hContextMenuDll);
        }
    }

    void AddToBinder()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);
        typedef void (*EntryPointfuncPtr)(TCHAR**, unsigned int);
        if (HINSTANCE hContextMenuDll = LoadContextMenuDll()) {
            EntryPointfuncPtr funcProc = (EntryPointfuncPtr)GetProcAddress(hContextMenuDll, "AddToBinder");
            if (funcProc) {
                unsigned int len = static_cast<unsigned int>(m_filelist.size());
                TCHAR** fileList = allocArrayFromVector(m_filelist);
                if (fileList) {
                    funcProc(fileList, len);
                    deallocArray(fileList, len);
                }
            }
            FreeLibrary(hContextMenuDll);
        }
    }

    void EditPDFInAcrobat()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);
        typedef void (*EntryPointfuncPtr)(TCHAR**, unsigned int);
        if (HINSTANCE hContextMenuDll = LoadContextMenuDll())
        {
            EntryPointfuncPtr funcProc = (EntryPointfuncPtr)GetProcAddress(hContextMenuDll, "EditPDFInAcrobat");
            if (funcProc) {
                unsigned int len = static_cast<unsigned int>(m_filelist.size());
                TCHAR** fileList = allocArrayFromVector(m_filelist);
                if (fileList) {
                    funcProc(fileList, len);
                    deallocArray(fileList, len);
                }
            }
            FreeLibrary(hContextMenuDll);
        }
    }

    void SharePDF()
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__);
        LOG_TIME_TAKEN(msg_log_timer);
        typedef void(*EntryPointfuncPtr)(CString, unsigned int);
        if (HINSTANCE hContextMenuDll = LoadContextMenuDll())
        {
            EntryPointfuncPtr funcProc = (EntryPointfuncPtr)GetProcAddress(hContextMenuDll, "SharePDF");
            if (funcProc) {
                unsigned int len = static_cast<unsigned int>(m_filelist.size());
                if (len != 1)				//We support sharing of only single file as of now
                    return;
                CString fileName = m_filelist[0].c_str();
                if (fileName)
                    funcProc(fileName, len);
            }
            FreeLibrary(hContextMenuDll);
        }
    }
    
    

    virtual std::wstring GetTitle(int commandId)
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(commandId));
        LOG_TIME_TAKEN(msg_log_timer);

        UINT resId = -1;
        switch (commandId)
        {
        case CREATE_ADOBE_PDF:
            resId = IDS_CONVERT_TO_PDF;
            break;
#if 1 // share being dropped
        case CREATE_AND_EMAIL_ADOBE_PDF:
            resId = IDS_CONVERT_AND_EMAIL;
            break;
        case SHARE_PDF_FOR_REVIEW:
            //return L"More Option in Acrobat...";
            resId = IDS_SHARE_PDF_FOR_REVIEW;
            break;
#endif
        case ADD_TO_ACROBAT_BINDER:
            resId = IDS_ADD_TO_ACROBAT_BINDER;
            break;
        case EDIT_PDF_IN_ACROBAT:
            resId = IDS_EDIT_PDF_IN_ACROBAT;
            break;
        default:
            // I include default for completeness.
            // I catch more bugs by always adding this
            // to a switch that I care to count.
            return L"";
        }

        

        TCHAR szMenu[MAX_PATH] = {};
        LoadString(m_hresource, resId, szMenu, sizeof(szMenu) / sizeof(TCHAR));
        return szMenu;
    }
    
    EXPCMDSTATE GetCommandState(int commandId)
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(commandId));
        LOG_TIME_TAKEN(msg_log_timer);

        int formatType = GetOfficeFormat(m_filelist[0].c_str());

        EXPCMDSTATE retval = ECS_HIDDEN;

        switch (commandId)
        {
        case CREATE_ADOBE_PDF:
        {
            if ((IsAcrobatProInstalled() || IsFormatSupportedWithStandard(formatType)) &&
                IsThisAdobePDF(m_filelist[0].substr(m_filelist[0].length() - 4, 5 /* 1 extra for null*/).c_str()) == FALSE &&
                (GetFileAttributes(m_filelist[0].c_str()) & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                    retval = ECS_ENABLED;
            }
                break;
        }
        case ADD_TO_ACROBAT_BINDER:
        {
          if (IsAcrobatElementsOnlyInstalled() != TRUE &&
                (((GetFileAttributes(m_filelist[0].c_str()) & FILE_ATTRIBUTE_DIRECTORY)/* && (uFlags ^ (CMF_EXPLORE | CMF_VERBSONLY)))*/ ||
                    (IsThisImageFileSupported(m_filelist[0].substr(m_filelist[0].length() - 11, 12 /* 1 extra for null*/).c_str()) == TRUE ||
                        IsThisAdobePDF(m_filelist[0].substr(m_filelist[0].length() - 4, 5 /* 1 extra for null*/).c_str())) == TRUE ||
                    IsThisSupportedCombineOfficeFormat(formatType))))

            {   // Show binder menu only for the file type which can be handled by binder
                  retval = ECS_ENABLED;
            }
              break;
        }
        case EDIT_PDF_IN_ACROBAT:
        {
            if (IsAcrobatElementsOnlyInstalled() != TRUE &&
                IsThisAdobePDF(m_filelist[0].substr(m_filelist[0].length() - 4, 5 /* 1 extra for null*/).c_str()) == TRUE &&
                m_filelist.size() == 1) // Let's Edit option is only for single document
            {   // Show Edit menu only for the file type which can be handled
                    retval = ECS_ENABLED;
            }
                break;
        }
#if 1 // share being dropped
        case CREATE_AND_EMAIL_ADOBE_PDF:
        {
            if ((IsAcrobatProInstalled() || IsFormatSupportedWithStandard(formatType)) &&
                m_filelist.size() == 1 &&  // email option is only for single document
                !(GetFileAttributes(m_filelist[0].c_str()) & FILE_ATTRIBUTE_DIRECTORY)) // Check if this is folder
            {
                if (IsThisImageFileSupported(m_filelist[0].substr(m_filelist[0].length() - 11, 12 /* 1 extra for null*/).c_str()) == FALSE &&
                    IsThisAdobePDF(m_filelist[0].substr(m_filelist[0].length() - 4, 5 /* 1 extra for null*/).c_str()) == FALSE) // No Email option for Image files and Adobe PDF files
                {
                        retval = ECS_ENABLED;
                }

            }
                break;
        }
        case SHARE_PDF_FOR_REVIEW:
        {
            LPCWSTR featureReg = L"iEnableContextMenuShare";
            int threshold = GetFeatureEnableThreshold(featureReg);

            if (IsAcrobatElementsOnlyInstalled() != TRUE &&
                IsThisAdobePDF(m_filelist[0].substr(m_filelist[0].length() - 4, 5 /* 1 extra for null*/).c_str()) == TRUE &&
                m_filelist.size() == 1 && EnableFeature(threshold)) // Share option is only for single document
            {   // Show Share PDF option only for the file type which can be handled

                    retval = ECS_ENABLED;
            }
                break;
        }
#endif
        default:
            // I include default for completeness.
            // I catch more bugs by always adding this
            // to a switch that I care to count.
                break;
        }
        MessageBox(NULL,NULL,NULL,NULL);
        return retval;
    }

    void GetCommandStateDelayed()
    {
        m_result = GetCommandState(m_commandId);
        return;
    }
     
    HRESULT State(_In_opt_ IShellItemArray* selection,_In_ BOOL okToBeSlow, EXPCMDSTATE* state, int commandId)
    { 
        if (okToBeSlow)
            OutputDebugStringA("----------getstate slow call--------");
        else
            OutputDebugStringA("----------getstate fast call--------");

        std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(commandId));
        LOG_TIME_TAKEN(msg_log_timer);

        auto start = std::chrono::high_resolution_clock::now();

        m_filelist.clear();
        DWORD count = 0;
        RETURN_IF_FAILED(selection->GetCount(&count));
        
        IEnumShellItems* itemEnum = nullptr;
        RETURN_IF_FAILED(selection->EnumItems(&itemEnum));

        std::vector<IShellItem*> items(count);

        ULONG fetched = 0;
        do {
            ULONG currfetched = 0;
            RETURN_IF_FAILED(itemEnum->Next(count-fetched, &items[fetched], &currfetched));
            fetched += currfetched;

        } while (fetched < count);
        
        size_t i = 0;
        for (auto it : items)
        {
            LPWSTR ppszName = nullptr;
            if (it->GetDisplayName(SIGDN_FILESYSPATH, &ppszName) != S_OK || !ppszName)
                continue;
            
           if (GetFileAttributes(ppszName) & FILE_ATTRIBUTE_DIRECTORY) {
                m_filelist.push_back(ppszName);
            }
            else if (IsThisSupportedFileObjects(ppszName)) {
                m_filelist.push_back(ppszName);
            }

            CoTaskMemFree(ppszName);

            i++;
        }

        for (auto it : items)
            it->Release();
        items.clear();

        if (m_filelist.size() == 0)//Not for NetWork Places, My Documents
        {
            *state = ECS_HIDDEN;
            return S_OK;
        }

        // for 64bit acrobat, hide context menu for visio, project and atuocad(if autocad not installed)
        if(Is64BitAcrobatInstalled())
        {
            std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(commandId) + "\tBlock:GetOfficeFormat" );
            LOG_TIME_TAKEN(msg_log_timer);
            int formatType = GetOfficeFormat(m_filelist[0].c_str());
            if ((((formatType == VISIO_FILE)
                    || (formatType == PROJECT_FILE)
                    || ((formatType == AUTOCAD_FILE) && CheckIfAutoCADIsInstalled() == false)))
                )
            {
                *state = ECS_HIDDEN;

                std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(commandId) + "Is64BitAcrobatInstalled");
                LOG_TIME_TAKEN(msg_log_timer);

                return S_OK;
            }
        }
        
        
        if (!okToBeSlow) 
        {

            if (m_timer)
                ::KillTimer(0,m_timer); // kill old timer if there is any

            BOOL val = FALSE;
            SetUserObjectInformationW(::GetCurrentProcess(), UOI_TIMERPROC_EXCEPTION_SUPPRESSION,&val,sizeof(val));
            m_timer = SetTimer(0, 0, USER_TIMER_MINIMUM, TimerprocGetState); // set new timer
            m_commandId = commandId;
            
            *state = ECS_HIDDEN;
            return E_PENDING;
        }
        else
        {
            EXPCMDSTATE retval = ECS_ENABLED;
            if (m_timer && m_commandId== commandId && m_result != -1)
        {
                ::OutputDebugStringW(L"result is available");
                retval = m_result;
        }
        else
            {
                ::OutputDebugStringW(L"result is not available");
         
                retval = GetCommandState(commandId);
            }

            if (m_timer)
        {
                ::KillTimer(0, m_timer);
            }
            m_result = -1;
            m_timer = 0;
            m_commandId = -1;
            *state = retval;
            ::OutputDebugStringW(L"result S_OK");
            return S_OK;
        }
        return S_OK; 
    }


    std::wstring GetIcon(int commandId) {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(commandId));
        LOG_TIME_TAKEN(msg_log_timer);
      
        std::wstring icon_str(m_shim64_iexplore_dll_path);
        icon_str = icon_str + L",-";
        
        switch (commandId)
        {
        case CREATE_ADOBE_PDF:
            icon_str = icon_str + std::to_wstring(GetOSAppTheme() == LightTheme ? IDB_CREATEPDF_LIGHT_ICO : IDB_CREATEPDF_DARK_ICO);
            break;
        case ADD_TO_ACROBAT_BINDER:
            icon_str = icon_str + std::to_wstring(GetOSAppTheme() == LightTheme ? IDB_BINDER_LIGHT_ICO : IDB_BINDER_DARK_ICO);
            break;
        case EDIT_PDF_IN_ACROBAT:
            icon_str = icon_str + std::to_wstring(GetOSAppTheme() == LightTheme ? IDB_EDITPDF_LIGHT_ICO : IDB_EDITPDF_DARK_ICO);
            break;
#if 1 // share being dropped
        case CREATE_AND_EMAIL_ADOBE_PDF:
            //icon_str = icon_str + std::to_wstring(GetOSAppTheme() == LightTheme ? IDB_CREATEANDEMAIL_LIGHT_ICO : IDB_CREATEANDEMAIL_DARK_ICO);
            break;
#ifdef CHANGE_CONVERSION_SETTINGS
        case CHANGE_CONVERSION_SETTINGS:
            ChangeConversionSettings();
            break;
#endif
        case SHARE_PDF_FOR_REVIEW:
            //icon_str = L"@{Microsoft.WindowsNotepad_10.2103.6.0_x64__8wekyb3d8bbwe?ms-resource://Microsoft.WindowsNotepad/Files/Assets/NotepadAppList.scale-100.png}";
            icon_str =  icon_str + std::to_wstring(GetOSAppTheme() == LightTheme ? IDB_EDITPDF_LIGHT_ICO : IDB_EDITPDF_DARK_ICO);
            break;
#endif
        default:
            // I include default for completeness.
            // I catch more bugs by always adding this
            // to a switch that I care to count.
            return L"";
        }

        return icon_str;
    }


protected:
   
    HINSTANCE m_hresource;
    std::wstring m_shim64_dll_path;
    std::wstring m_shim64_iexplore_dll_path;
    std::wstring m_contextmenu64_dll_path;
    HMODULE m_shim64_module;
    
    std::vector<std::wstring> m_filelist;
    
    EXPCMDSTATE m_result = ECS_ENABLED;

    UINT_PTR m_timer;
    int m_commandId;
    //EXPCMDSTATE 
};


void WINAPI TimerprocGetState(
    HWND unnamedParam1,
    UINT unnamedParam2,
    UINT_PTR unnamedParam3,
    DWORD unnamedParam4
)
{
   AcroExplorerCommandBaseHelper::GetInstance().GetCommandStateDelayed();
}

class AcroExplorerCommandBase : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand, IObjectWithSite>
{
    ComPtr<IUnknown> m_site;
public:

    virtual int GetCommandId() = 0; // implemented by subclass, based on this param, icon, state and invoke action are decided

    // IObjectWithSite
    IFACEMETHODIMP SetSite(_In_ IUnknown* site) noexcept 
    {
        m_site = site; return S_OK; 
    }
    IFACEMETHODIMP GetSite(_In_ REFIID riid, _COM_Outptr_ void** site) noexcept
    { 
        return m_site.CopyTo(riid, site); 
   }

    IFACEMETHODIMP GetFlags(_Out_ EXPCMDFLAGS* flags) { *flags = ECF_DEFAULT; return S_OK; }
    IFACEMETHODIMP EnumSubCommands(_COM_Outptr_ IEnumExplorerCommand** enumCommands) { *enumCommands = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetToolTip(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* infoTip) { *infoTip = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetCanonicalName(_Out_ GUID* guidCommandName) {  return E_NOTIMPL; };
 

    IFACEMETHODIMP GetTitle(_In_opt_ IShellItemArray* items, _Outptr_result_nullonfailure_ PWSTR* name)
    {
        if (!IsWindows11OrGreater())
        {
            OutputDebugStringW((std::wstring(L"skipping GetTitle as windows <11: ")).c_str());
            return E_INVALIDARG;
        }

        std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(GetCommandId()));
        LOG_TIME_TAKEN(msg_log_timer);
        *name = nullptr;
        auto title = wil::make_cotaskmem_string_nothrow(AcroExplorerCommandBaseHelper::GetInstance().GetTitle(GetCommandId()).c_str());
        RETURN_IF_NULL_ALLOC(title);
        *name = title.release();
        return S_OK;
    }
    IFACEMETHODIMP GetIcon(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* icon_) {
        if (!IsWindows11OrGreater())
        {
            OutputDebugStringW((std::wstring(L"skipping GetIcon as windows <11: ")).c_str());
            return E_INVALIDARG;
        }

        std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(GetCommandId()));
        LOG_TIME_TAKEN(msg_log_timer);
        *icon_ = nullptr;
        auto icon = wil::make_cotaskmem_string_nothrow(AcroExplorerCommandBaseHelper::GetInstance().GetIcon(GetCommandId()).c_str());
        RETURN_IF_NULL_ALLOC(icon);
        *icon_ = icon.release();

        return S_OK;
    }
      IFACEMETHODIMP GetState(_In_opt_ IShellItemArray* selection, _In_ BOOL okToBeSlow, _Out_ EXPCMDSTATE* cmdState)
    {
          if (!IsWindows11OrGreater())
          {
              OutputDebugStringW((std::wstring(L"skipping GetState as windows <11: ")).c_str());
              *cmdState = ECS_HIDDEN;;
              return S_OK;
          }
        return  AcroExplorerCommandBaseHelper::GetInstance().State(selection, okToBeSlow, cmdState,GetCommandId());
    }
    IFACEMETHODIMP Invoke(_In_opt_ IShellItemArray* selection, _In_opt_ IBindCtx*) noexcept try
    {
        std::string msg_log_timer = (std::string("function") + __FUNCTION__ + std::to_string(GetCommandId()));
        LOG_TIME_TAKEN(msg_log_timer);

        if (IsInstallingUpdate()) {
            return S_OK;
        }

        // We have what we think is a valid command.
    // Invoke it.
        switch (GetCommandId())
        {
        case CREATE_ADOBE_PDF:
            AcroExplorerCommandBaseHelper::GetInstance().ConvertToPDF();
            break;
        case ADD_TO_ACROBAT_BINDER:
            AcroExplorerCommandBaseHelper::GetInstance().AddToBinder();
            break;
        case EDIT_PDF_IN_ACROBAT:
            AcroExplorerCommandBaseHelper::GetInstance().EditPDFInAcrobat();
            break;
#if 1 // share being dropped
        case CREATE_AND_EMAIL_ADOBE_PDF:
            AcroExplorerCommandBaseHelper::GetInstance().ConvertAndSharePDF();
            break;
        case SHARE_PDF_FOR_REVIEW:
            AcroExplorerCommandBaseHelper::GetInstance().SharePDF();
            break;
#endif
        default:
            // I include default for completeness.
            // I catch more bugs by always adding this
            // to a switch that I care to count.
            return E_INVALIDARG;
        }
        return S_OK;
    }
    CATCH_RETURN();

};

class __declspec(uuid("3282E233-C5D3-4533-9B25-44B8AAAFACFA")) AcroExplorerCommandCreateAdobePDF final : public AcroExplorerCommandBase
{
public:
    
    virtual int GetCommandId() override { return CREATE_ADOBE_PDF; }

};

class __declspec(uuid("1476525B-BBC2-4D04-B175-7E7D72F3DFF8")) AcroExplorerCommandAddtoAcrobatBinder final : public AcroExplorerCommandBase
{
public:
   virtual int GetCommandId() override { return ADD_TO_ACROBAT_BINDER; }
};

class __declspec(uuid("30DEEDF6-63EA-4042-A7D8-0A9E1B17BB99")) AcroExplorerCommandEditPDFInAcrobat final :public  AcroExplorerCommandBase
{
public:
   virtual int GetCommandId() override { return EDIT_PDF_IN_ACROBAT; }
};

#if 1 // share being dropped

class __declspec(uuid("817CF159-A4B5-41C8-8E8D-0E23A6605395")) AcroExplorerCommandCreateAndEmailAdobePDF final : public AcroExplorerCommandBase
{
public:
    virtual int GetCommandId() override { return CREATE_AND_EMAIL_ADOBE_PDF; }

};

class SubExplorerCommandHandler final : public AcroExplorerCommandBase
{
public:
    virtual int GetCommandId() override { return ADD_TO_ACROBAT_BINDER; }
};

class CheckedSubExplorerCommandHandler final : public AcroExplorerCommandBase
{
public:
    virtual int GetCommandId() override { return SHARE_PDF_FOR_REVIEW; }
};

class EnumCommands : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IEnumExplorerCommand>
{
public:
    EnumCommands()
    {
        m_commands.push_back(Make<SubExplorerCommandHandler>());
        m_commands.push_back(Make<CheckedSubExplorerCommandHandler>());
        m_current = m_commands.cbegin();
    }

    // IEnumExplorerCommand
    IFACEMETHODIMP Next(ULONG celt, __out_ecount_part(celt, *pceltFetched) IExplorerCommand** apUICommand, __out_opt ULONG* pceltFetched)
    {
        ULONG fetched{ 0 };
        wil::assign_to_opt_param(pceltFetched, 0ul);

        for (ULONG i = 0; (i < celt) && (m_current != m_commands.cend()); i++)
        {
            m_current->CopyTo(&apUICommand[0]);
            m_current++;
            fetched++;
        }

        wil::assign_to_opt_param(pceltFetched, fetched);
        return (fetched == celt) ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Skip(ULONG /*celt*/) { return E_NOTIMPL; }
    IFACEMETHODIMP Reset()
    {
        m_current = m_commands.cbegin();
        return S_OK;
    }
    IFACEMETHODIMP Clone(__deref_out IEnumExplorerCommand** ppenum) { *ppenum = nullptr; return E_NOTIMPL; }

private:
    std::vector<ComPtr<IExplorerCommand>> m_commands;
    std::vector<ComPtr<IExplorerCommand>>::const_iterator m_current;
};

class __declspec(uuid("50419A05-F966-47BA-B22B-299A95492348")) AcroExplorerCommandSharePDFForReview final :public  AcroExplorerCommandBase
{
public:
    IFACEMETHODIMP GetFlags(_Out_ EXPCMDFLAGS* flags) { *flags = ECF_HASSUBCOMMANDS; return S_OK; }
    virtual int GetCommandId() override { return SHARE_PDF_FOR_REVIEW; }
    IFACEMETHODIMP EnumSubCommands(_COM_Outptr_ IEnumExplorerCommand** enumCommands)
    {
        *enumCommands = nullptr;
        auto e = Make<EnumCommands>();
        return e->QueryInterface(IID_PPV_ARGS(enumCommands));
    }
};
#endif

CoCreatableClass(AcroExplorerCommandCreateAdobePDF)
CoCreatableClass(AcroExplorerCommandAddtoAcrobatBinder)
CoCreatableClass(AcroExplorerCommandEditPDFInAcrobat)
#if 1 // share being dropped
CoCreatableClass(AcroExplorerCommandCreateAndEmailAdobePDF)
CoCreatableClass(AcroExplorerCommandSharePDFForReview)
#endif

CoCreatableClassWrlCreatorMapInclude(AcroExplorerCommandCreateAdobePDF)
CoCreatableClassWrlCreatorMapInclude(AcroExplorerCommandAddtoAcrobatBinder)
CoCreatableClassWrlCreatorMapInclude(AcroExplorerCommandEditPDFInAcrobat)
#if 1 // share being dropped
CoCreatableClassWrlCreatorMapInclude(AcroExplorerCommandCreateAndEmailAdobePDF)
CoCreatableClassWrlCreatorMapInclude(AcroExplorerCommandSharePDFForReview)
#endif

STDAPI DllGetActivationFactory(_In_ HSTRING activatableClassId, _COM_Outptr_ IActivationFactory** factory)
{
    return Module<ModuleType::InProc>::GetModule().GetActivationFactory(activatableClassId, factory);
}

STDAPI DllCanUnloadNow()
{
    return Module<InProc>::GetModule().GetObjectCount() == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _COM_Outptr_ void** instance)
{
    return Module<InProc>::GetModule().GetClassObject(rclsid, riid, instance);
}

int GetOfficeFormat(const CString* pBuffer);
int GetOfficeFormat(const TCHAR* pBuffer)
{
    CString tmp(pBuffer);
   return GetOfficeFormat(&tmp);
}
BOOL ValidateFileName(CString szFileName, BOOL* pPrintTo, BOOL* pPrint)
{
    TCHAR szExe[2 * MAX_PATH] = { 0 };
    DWORD cchOut = sizeof(szExe);
    LPCTSTR pExtension = (LPCTSTR)szFileName;
    int index = szFileName.ReverseFind('.');
    BOOL bFoundApp = FALSE;
    if (index > 0) {
        TCHAR* pParam;
        pExtension += index;

        AssocQueryString(ASSOCF_NOUSERSETTINGS,
            ASSOCSTR_COMMAND,
            pExtension,
            NULL,
            szExe,
            &cchOut);
        _wcsupr(szExe);
        pParam = wcsstr(szExe, _T(".EXE "));
        if (!pParam)
            pParam = wcsstr(szExe, _T(".EXE\" "));
        if (pParam)
        {
            *(pParam + 4) = 0;
            bFoundApp = TRUE;
        }
    }
    if (wcslen(szExe) || FindExecutable(szFileName, NULL, szExe) > (HINSTANCE)32)
    {
        TCHAR szApplication[MAX_PATH];
        bFoundApp = TRUE;
        GetModuleFileName(0, szApplication, _countof(szApplication));
        _wcsupr(szApplication);
        if (wcsstr(szExe, szApplication))
            return bFoundApp; // User is trying to convert from the application File Open Dialog

        TCHAR* pFile, szSubKey[MAX_PATH], szPath[MAX_PATH];
        wcscpy(szSubKey, _T("LastPdfPortFolder - "));
        pFile = szSubKey + wcslen(szSubKey);
        _wcsupr(szExe);
        _tsplitpath(szExe, NULL, NULL, pFile, NULL);
        wcscat(szSubKey, _T(".EXE"));
        wcscpy(szPath, szFileName);
        pFile = _tcsrchr(szPath, '\\');
        if (pFile)
            *pFile = 0;
        DWORD dwOut = MAX_PATH;
        TCHAR szOut[MAX_PATH] = { 0 };
        AssocQueryString(ASSOCF_OPEN_BYEXENAME,
            ASSOCSTR_COMMAND,
            szExe,
            _T("printto"),
            szOut,
            &dwOut);
        //		GetLastError();
        szFileName.MakeUpper();
        if ((wcslen(szOut) && szFileName.Right(4) != _T(".CDR")) || //CorelDraw "PrintTo" doesn't work on XP
            szFileName.Right(4) == _T(".HTM") ||
            szFileName.Right(5) == _T(".HTML"))
            *pPrintTo = TRUE;
        else
        {
            *pPrintTo = FALSE;
            szOut[0] = 0;
            AssocQueryString(ASSOCF_OPEN_BYEXENAME,
                ASSOCSTR_COMMAND,
                szExe,
                _T("print"),
                szOut,
                &dwOut);

            if (wcslen(szOut))
            {
                if (_wcsnicmp(szOut, _T("rundll32.exe "), sizeof("rundll32.exe ") - 1) == 0)
                {	//Windows XP prints BMP,JPG,PNG,TIF files using Imageviewer DLL 
                    wcscpy(szSubKey, _T("LastPdfPortFolder - RUNDLL32.EXE"));
                    *pPrintTo = TRUE;
                }
                else
                    *pPrint = TRUE;
            }

        }
        if (!*pPrint && !*pPrintTo)
        {
            HKEY hKey;
            if (RegOpenKey(HKEY_CLASSES_ROOT, pExtension, &hKey) == ERROR_SUCCESS)
            {
                TCHAR szValue[MAX_PATH / 2];
                DWORD lData = sizeof(szValue);
                DWORD lType;
                RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
                if (wcslen(szValue))
                {
                    TCHAR szKeyValue[MAX_PATH];
                    wcscpy(szKeyValue, szValue);
                    wcscat(szKeyValue, _T("\\shell\\printto\\command"));
                    RegCloseKey(hKey);
                    if (RegOpenKey(HKEY_CLASSES_ROOT, szKeyValue, &hKey) == ERROR_SUCCESS)
                    {
                        *pPrintTo = TRUE;
                        RegCloseKey(hKey);
                    }
                    else
                    {
                        wcscpy(szKeyValue, szValue);
                        wcscat(szKeyValue, _T("\\shell\\print\\command"));
                        if (RegOpenKey(HKEY_CLASSES_ROOT, szKeyValue, &hKey) == ERROR_SUCCESS)
                        {
                            *pPrint = TRUE;
                            RegCloseKey(hKey);
                        }
                    }
                }
                else
                    RegCloseKey(hKey);
            }
        }
    }
    return bFoundApp;
}

BOOL IsThisImageFileSupported(LPCTSTR pExt)
{
    if (pExt == NULL)
        return FALSE;

    BOOL bIsStd = IsAcrobatStdInstalled();
    BOOL bIsPro = IsAcrobatProInstalled();

    BOOL bIsElem = FALSE;
    BOOL bIs3D = FALSE;

    //load Adist32.dll and peform ProductID/SIF check for all 4 products
    HINSTANCE hInst = LoadAdist32Dll();
    if (hInst)
    {
        typedef BOOL(*PIS_ELEMENTS_INSTALLED)(void);
        PIS_ELEMENTS_INSTALLED pIsElements;
        if (pIsElements = (PIS_ELEMENTS_INSTALLED)GetProcAddress(hInst, (LPCSTR)"IsAcrobatElementsInstalled"))
        {
            bIsElem = pIsElements();
        }

        typedef BOOL(*PIS_3D_INSTALLED)(void);
        PIS_3D_INSTALLED pIs3D;
        if (pIs3D = (PIS_3D_INSTALLED)GetProcAddress(hInst, (LPCSTR)"IsAcrobat3DInstalled"))
        {
            bIs3D = pIs3D();
        }
        FreeLibrary(hInst);
    }

    if (bIsElem)	//bail...don't show any context menu for Elements.
        return FALSE;

    //now test for recognized filetype, limited to product installed
    if (wcslen(pExt) == 11)
    {
        if (bIs3D)
        {
            if (_wcsicmp(pExt, _T(".CATPRODUCT")) == 0) // Acrobat 3D filetype
                return TRUE;
        }
        pExt++;
        pExt++;
        pExt++;
    }
    if (wcslen(pExt) == 8)
    {
        if (bIs3D)
        {
            if (_wcsicmp(pExt, _T(".CATPART")) == 0) // Acrobat 3D filetype
                return TRUE;
            if (_wcsicmp(pExt, _T(".SESSION")) == 0) // Acrobat 3D 9.0 filetype
                return TRUE;
        }
        pExt++;
    }
    if (wcslen(pExt) == 7)
    {
        if (bIs3D)
        {
            if (_wcsicmp(pExt, _T(".SLDASM")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".SLDPRT")) == 0)   // Acrobat 3D filetype
                return TRUE;
        }
        pExt++;
    }
    if (wcslen(pExt) == 6)
    {
        if (bIs3D || bIsPro || bIsStd)
        {
            if (_wcsicmp(pExt, _T(".SHTML")) == 0)
                return TRUE;
        }
        if (bIs3D)
        {
            if (_wcsicmp(pExt, _T(".MODEL")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".CADDS")) == 0 ||  // Acrobat 3D 9.0 filetype
                _wcsicmp(pExt, _T(".3DXML")) == 0)  // Acrobat 3D filetype
                return TRUE;
        }
        pExt++;
    }
    if (wcslen(pExt) == 5)
    {
        if (bIs3D || bIsPro || bIsStd)
        {
            if (
                _wcsicmp(pExt, _T(".TIFF")) == 0 ||
                _wcsicmp(pExt, _T(".HTML")) == 0 ||
                _wcsicmp(pExt, _T(".TEXT")) == 0 ||
                _wcsicmp(pExt, _T(".JPEG")) == 0)
                return TRUE;
        }
        if (bIs3D)
        {
            if (_wcsicmp(pExt, _T(".IGES")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".SDPC")) == 0 ||  // Acrobat 3D OnSpace filetype
                _wcsicmp(pExt, _T(".SDWC")) == 0 ||  // Acrobat 3D 9.0 filetype
                _wcsicmp(pExt, _T(".STEP")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".VRML")) == 0)   // Acrobat 3D filetype
                return TRUE;
        }
        pExt++;
    }
    if (wcslen(pExt) == 4)
    {
        if (bIs3D || bIsPro || bIsStd)
        {
            if (_wcsicmp(pExt, _T(".JPG")) == 0 ||
                _wcsicmp(pExt, _T(".BMP")) == 0 ||
                _wcsicmp(pExt, _T(".JPE")) == 0 ||
                _wcsicmp(pExt, _T(".J2K")) == 0 ||
                _wcsicmp(pExt, _T(".J2C")) == 0 ||
                _wcsicmp(pExt, _T(".JPF")) == 0 ||
                _wcsicmp(pExt, _T(".JPX")) == 0 ||
                _wcsicmp(pExt, _T(".JP2")) == 0 ||
                _wcsicmp(pExt, _T(".JPC")) == 0 ||
                _wcsicmp(pExt, _T(".GIF")) == 0 ||
                _wcsicmp(pExt, _T(".PNG")) == 0 ||
                _wcsicmp(pExt, _T(".PCX")) == 0 ||
                _wcsicmp(pExt, _T(".TIF")) == 0 ||
                _wcsicmp(pExt, _T(".HTM")) == 0 ||
                _wcsicmp(pExt, _T(".DIB")) == 0 ||
                _wcsicmp(pExt, _T(".RLE")) == 0 ||
                _wcsicmp(pExt, _T(".TXT")) == 0 ||
                _wcsicmp(pExt, _T(".PRN")) == 0 ||  // .prn files will be now treated as PostScript files
                _wcsicmp(pExt, _T(".XPS")) == 0 ||  // .xps files, Acrobat Pro and Standard support this file type.
                _wcsicmp(pExt, _T(".EPS")) == 0)
                return TRUE;
        }
        if (bIs3D)
        {
            if (_wcsicmp(pExt, _T("._PD")) == 0 ||  // Acrobat 3D 9.0 filetype
                _wcsicmp(pExt, _T(".3DS")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".ARC")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".ASM")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".BDL")) == 0 ||  // Acrobat 3D OnSpace filetype
                _wcsicmp(pExt, _T(".CGR")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".DAE")) == 0 ||  // Acrobat 3D 9.0 filetype
                _wcsicmp(pExt, _T(".DLV")) == 0 ||  // Acrobat 3D filetype
// All AutoCAD supported file formats will be converted through PDFMaker, so removing DXF from this list.
//				_wcsicmp(pExt, _T(".DXF")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".EXP")) == 0 ||  // Acrobat 3D 9.0 filetype
_wcsicmp(pExt, _T(".IGS")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".IAM")) == 0 ||  // Acrobat 3D filetype support for AutoCad Inventor (Apprentice)
_wcsicmp(pExt, _T(".IFC")) == 0 ||  // Acrobat 3D 9.0 filetype
_wcsicmp(pExt, _T(".IPT")) == 0 ||  // Acrobat 3D filetype support for AutoCad Inventor (Apprentice)
_wcsicmp(pExt, _T(".MF1")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".NEU")) == 0 ||  // Acrobat 3D 9.0 filetype
_wcsicmp(pExt, _T(".OBJ")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".PAR")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".PKG")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".PRC")) == 0 ||  // Acrobat 3D 9.0 filetype
_wcsicmp(pExt, _T(".PRD")) == 0 ||  // Acrobat 3D 9.0 filetype
_wcsicmp(pExt, _T(".PRT")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".SAT")) == 0 ||  // Acrobat 3D acis filetype
_wcsicmp(pExt, _T(".SAB")) == 0 ||  // Acrobat 3D 9.0 filetype
_wcsicmp(pExt, _T(".SDA")) == 0 ||  // Acrobat 3D OnSpace filetype
_wcsicmp(pExt, _T(".SDP")) == 0 ||  // Acrobat 3D OnSpace filetype
_wcsicmp(pExt, _T(".SDS")) == 0 ||  // Acrobat 3D OnSpace filetype
_wcsicmp(pExt, _T(".SDW")) == 0 ||  // Acrobat 3D OnSpace filetype
_wcsicmp(pExt, _T(".SES")) == 0 ||  // Acrobat 3D OnSpace filetype
_wcsicmp(pExt, _T(".STL")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".STP")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".U3D")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".UNV")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".WRL")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".X_B")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".X_T")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".XAS")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".XPR")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".XV3")) == 0 ||  // Acrobat 3D filetype
_wcsicmp(pExt, _T(".XV0")) == 0)   // Acrobat 3D filetype
return TRUE;
        }
        pExt++;
    }
    if (wcslen(pExt) == 3)
    {
        if (bIs3D || bIsPro || bIsStd)
        {
            if (_wcsicmp(pExt, _T(".PS")) == 0)
                return TRUE;
        }
        if (bIs3D)
        {
            if (_wcsicmp(pExt, _T(".JT")) == 0 ||  // Acrobat 3D filetype
                _wcsicmp(pExt, _T(".PD")) == 0)  // Acrobat 3D 9.0 filetype
                return TRUE;
        }
    }
    return FALSE;
}

BOOL IsThisSupportedFileObjects(TCHAR* pBuffer)
{

    CString szTemp;
    szTemp = pBuffer;
#if WIN_ENV
    Logger::LogEvent(" Inside IsThisSupportedFileObjects in Context MEnu ");
    char* destFolder;
    int dwSize = WideCharToMultiByte(CP_ACP, 0, pBuffer, -1, NULL, NULL, NULL, NULL);
    if (dwSize)
    {
        destFolder = (char*)malloc(dwSize * sizeof(char));
        if (destFolder)
        {
            destFolder[0] = '\0';
            WideCharToMultiByte(CP_ACP, 0, pBuffer, -1, destFolder, dwSize, NULL, NULL);
        }
        if (destFolder)
        {
            Logger::LogEvent("Is Supported File Type path: ");
            Logger::LogEvent(destFolder);
            free(destFolder);
        }

    }

#endif
    if (szTemp.Right(4).CompareNoCase(_T(".PRN")) == 0)
        return TRUE; // .prn files will be now treated as PostScript files
    if (szTemp.Right(3).CompareNoCase(_T(".PS")) == 0)
        return TRUE;
    if (szTemp.Right(4).CompareNoCase(_T(".EPS")) == 0)
        return TRUE;
    else if (szTemp.Right(4).CompareNoCase(_T(".PDF")) == 0)
        //(Vegas Feature #3211)
        //Allow a user to right-click on PDFs in Win desktop and add to MultiFile PDF Window.  This would include:
        //	* select a single PDF from desktop and choose Combine in PDF
        //	* select multiple PDFs from desktop and choose Combine in PDF
        //	* select one/more PDFs with one/more other files and have them all appear correctly in MultiFile PDF window.
        if (IsAcrobatElementsOnlyInstalled() == FALSE)
            return TRUE;
        else
            return FALSE;
    else if (szTemp.Right(3).CompareNoCase(_T(".AI")) == 0)
        // Illustrator should not be supported by Elements (Vegas Feature #3417 )
        return FALSE;
    else if (szTemp.Right(4).CompareNoCase(_T(".ZIP")) == 0) // We should confuse users with contextMenu for ZIP files (1089733)
        return FALSE;
    else if (szTemp.Right(4).CompareNoCase(_T(".FDF")) == 0)
        return FALSE;
    else if (szTemp.Right(4).CompareNoCase(_T(".XFDF")) == 0)
        return FALSE;
    else if (szTemp.Right(4).CompareNoCase(_T(".EXE")) == 0)
        return FALSE;
    else if (szTemp.Right(4).CompareNoCase(_T(".DLL")) == 0)
        return FALSE;
    else if (szTemp.Right(4).CompareNoCase(_T(".MDB")) == 0)
        return FALSE;

    if (IsThisImageFileSupported(szTemp.Right(11)))
        return TRUE;

    int formatType = GetOfficeFormat(&szTemp);
    if (formatType != NON_OFFICE_FILE && IsThisSupportedOfficeFormat(formatType))
        return TRUE;

    BOOL bPrintTo = FALSE;
    BOOL bPrint = FALSE;

    // Make sure we deal with only file types supporting "Print" or "PrintTo" command
    ValidateFileName(pBuffer, &bPrintTo, &bPrint);
    return (bPrint || bPrintTo);
}

int GetOfficeFormat(const CString* pBuffer)
{
    CString szTemp;
    szTemp = *pBuffer;
    int retVal = NON_OFFICE_FILE;
    // Check to see if this OFFICE files;
    int iDot = szTemp.ReverseFind('.');
    if (iDot == -1)
        return retVal;
    szTemp = szTemp.Right(szTemp.GetLength() - iDot);
    szTemp.MakeUpper();

#if WIN_ENV
    Logger::LogEvent(" Inside GetOfficeFormat in Context MEnu ");
    char* destFolder;
    int dwSize = WideCharToMultiByte(CP_ACP, 0, szTemp, -1, NULL, NULL, NULL, NULL);
    if (dwSize)
    {
        destFolder = (char*)malloc(dwSize * sizeof(char));
        if (destFolder)
        {
            destFolder[0] = '\0';
            WideCharToMultiByte(CP_ACP, 0, szTemp, -1, destFolder, dwSize, NULL, NULL);
        }
        if (destFolder)
        {
            Logger::LogEvent("GetOfficeFormat:  ");
            Logger::LogEvent(destFolder);
            free(destFolder);
        }

    }

#endif

    if (szTemp == _T(".PPT") ||
        szTemp == _T(".PPTX") || //Powerpoint 2007 file format
        szTemp == _T(".PPTM")) //Powerpoint 2007 Macro file format
        // || szTemp == _T(".PPS") )	// PowerPoint slides not supported by PDFMaker
    {
        HKEY hKey;
        if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
        {
            TCHAR szValue[MAX_PATH];
            DWORD lData = sizeof(szValue);
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
            RegCloseKey(hKey);
            if (wcsncmp(szValue, _T("PowerPoint."), sizeof("PowerPoint.") - 1) == 0)
                retVal = POWER_POINT_FILE;
        }
    }
    else if (szTemp == _T(".DOC") || 	// Word or Wordpad Documents
        szTemp == _T(".DOCX") || 	// Word 2007 file format
        szTemp == _T(".DOCM") || 	// Word 2007 Macro file format
        szTemp == _T(".RTF")) //Rich text format
    {
        HKEY hKey;
#if WIN_ENV
        Logger::LogEvent(" Opening Registry for Word document ");
#endif
        if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
        {
            TCHAR szValue[MAX_PATH];
            DWORD lData = sizeof(szValue);
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
#if WIN_ENV
            Logger::LogEvent("szValue for Opening Registry for Word document ");
            char* destFolder = Logger::GetCharEquivOfWideChar(szValue);
            if (destFolder)
            {
                Logger::LogEvent("Registry Value");
                Logger::LogEvent(destFolder);
                free(destFolder);
            }
#endif
            RegCloseKey(hKey);
            if (wcsncmp(szValue, _T("Word.Document"), sizeof("Word.Document") - 1) == 0)
                retVal = WORD_FILE;
            else if (wcsncmp(szValue, _T("Word.RTF"), sizeof("Word.RTF") - 1) == 0)
                retVal = WORD_FILE;
        }
    }
    else if (szTemp == _T(".XLS") || // Excel Spread sheets
        szTemp == _T(".XLSX") || 	// Excel 2007 file format
        szTemp == _T(".XLSM") || 	// Excel 2007 Macro file format
        szTemp == _T(".XLSB") || 	// Excel 2007 Binary file format
        szTemp == _T(".XLB") ||
        szTemp == _T(".XLW")) 		// Excel Worksheet
    {
        HKEY hKey;
#if WIN_ENV
        Logger::LogEvent("Reading Registry for Excel DOc");
#endif
        if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
        {
            TCHAR szValue[MAX_PATH];
            DWORD lData = sizeof(szValue);
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
#if WIN_ENV
            Logger::LogEvent("szValue for Opening Registry for Excel document ");
            char* destFolder = Logger::GetCharEquivOfWideChar(szValue);
            if (destFolder)
            {
                Logger::LogEvent("Registry Value");
                Logger::LogEvent(destFolder);
                free(destFolder);
            }
#endif
            RegCloseKey(hKey);
            if (wcsncmp(szValue, _T("Excel."), sizeof("Excel.") - 1) == 0)
                retVal = EXCEL_FILE;
        }
    }
    else if (szTemp == _T(".VSD") || // Visio Drawing
             szTemp == _T(".VSDX") ||// Visio Drawing
             szTemp == _T(".VSS") || // Vision Stencil
             szTemp == _T(".VSSX"))	 // Vision Stencil
    {
        HKEY hKey;
        if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
        {
            TCHAR szValue[MAX_PATH];
            DWORD lData = sizeof(szValue);
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
            RegCloseKey(hKey);
            if (wcsncmp(szValue, _T("Visio."), sizeof("Visio.") - 1) == 0)
                retVal = VISIO_FILE;
        }
    }
    else if (szTemp == _T(".DWG") || // AutoCAD Drawing
        szTemp == _T(".DWF") || // AutoCAD Viewer Files
        szTemp == _T(".DXF") || // AutoCAD Files
        szTemp == _T(".DST") || // AutoCAD Sheet Set Files
        szTemp == _T(".DWT"))  // AutoCAD Template files
    {
        if (IsAcrobatElementsOnlyInstalled())
        {
            HKEY hKey;
            if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
            {
                TCHAR szValue[MAX_PATH];
                DWORD lData = sizeof(szValue);
                DWORD lType;
                RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
                RegCloseKey(hKey);
                if (wcsncmp(szValue, _T("AutoCAD."), sizeof("AutoCAD.") - 1) == 0 ||
                    wcsncmp(szValue, _T("AutoCADTemplate."), sizeof("AutoCADTemplate.") - 1) == 0)
                    retVal = AUTOCAD_FILE;
            }
        }
        else
            retVal = AUTOCAD_FILE;
    }
    else if (szTemp == _T(".MPP")) // MSProject file
    {
        HKEY hKey;
        if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
        {
            TCHAR szValue[MAX_PATH];
            DWORD lData = sizeof(szValue);
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
            RegCloseKey(hKey);
            if (wcsncmp(szValue, _T("MSProject."), sizeof("MSProject.") - 1) == 0)
                retVal = PROJECT_FILE;
        }
    }
    else if (szTemp == _T(".PUB")) // MS PUblisher file
    {
        HKEY hKey;
        if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
        {
            TCHAR szValue[MAX_PATH];
            DWORD lData = sizeof(szValue);
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
            RegCloseKey(hKey);
            if (wcsncmp(szValue, _T("Publisher.Document"), sizeof("Publisher.Document") - 1) == 0)
                retVal = PUBLISHER_FILE;
        }
    }
    else if (szTemp == _T(".MDB")) // MS Access file
    {
        HKEY hKey;
        if (RegOpenKey(HKEY_CLASSES_ROOT, szTemp, &hKey) == ERROR_SUCCESS)
        {
            TCHAR szValue[MAX_PATH];
            DWORD lData = sizeof(szValue);
            DWORD lType;
            RegQueryValueEx(hKey, NULL, NULL, &lType, (LPBYTE)&szValue, &lData);
            RegCloseKey(hKey);
            if (wcsncmp(szValue, _T("Access.Application."), sizeof("Access.Application.") - 1) == 0)
                retVal = ACCESS_FILE;
        }
    }

    return retVal;
}

BOOL IsThisAdobePDF(LPCTSTR pExt)
{
    if (pExt != NULL && _wcsicmp(pExt, _T(".PDF")) == 0 &&
        IsAcrobatElementsOnlyInstalled() == FALSE)
        return TRUE;
    return FALSE;
}

BOOL GetAcrobatLanguage(TCHAR* pszLanguage, TCHAR* pszLocale)
{
    HKEY hkey;
    LONG err;
    DWORD type;
    TCHAR buf[256];
    DWORD cb = sizeof(buf);

    _tcscpy(pszLanguage, L".EXE"); //intialize to a known value

    if (RegOpenKey(HKEY_CURRENT_USER,
        ACROBAT_LANGUAGE_KEY, &hkey) != ERROR_SUCCESS &&
        RegOpenKey(HKEY_LOCAL_MACHINE,
            ACROBAT_LANGUAGE_KEY, &hkey) != ERROR_SUCCESS)
    {  // Acrobat is not installed now check for Reader
        if (RegOpenKey(HKEY_CURRENT_USER,
            ACROBAT_RDR_LANG_KEY, &hkey) != ERROR_SUCCESS &&
            RegOpenKey(HKEY_LOCAL_MACHINE,
                ACROBAT_RDR_LANG_KEY, &hkey) != ERROR_SUCCESS)
        { // Reader is not installed now check elements langauge preference
            if (RegOpenKey(HKEY_CURRENT_USER,
                ELEMENTS_LANGUAGE_KEY, &hkey) != ERROR_SUCCESS &&
                RegOpenKey(HKEY_LOCAL_MACHINE,
                    ELEMENTS_LANGUAGE_KEY, &hkey) != ERROR_SUCCESS)
                return FALSE; // Unable to find the lanugage preference
            {	// Acrobat not found let us read the Elements Language.
                err = RegQueryValueEx(hkey, L"UI", NULL, &type, (LPBYTE)buf, &cb);
                RegCloseKey(hkey);
                if (err == ERROR_SUCCESS && type == REG_SZ && cb <= 4)
                {
                    _tcscpy(pszLanguage, buf);
                    return TRUE;
                }
                else
                    return FALSE;
            }
        }
    }
    // Acrobat found let us check the Language preference
    err = RegQueryValueEx(hkey, NULL, NULL, &type, (LPBYTE)&buf[0], &cb);
    RegCloseKey(hkey);
    if (err == ERROR_SUCCESS && type == REG_SZ && cb >= 4)
    {
        _tcscpy(pszLanguage, &buf[wcslen(buf) - 4]);
        if (pszLocale && cb >= 13) /* locale\fr_FR\ */
            _tcsncpy(pszLocale, &buf[0], 13);
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

static const TCHAR* GetLocale(TCHAR* sLang)
{
    struct SLangToLocaleMap
    {
        const TCHAR* psLang;
        const TCHAR* psLocale;
    };
    static const SLangToLocaleMap m_langToLocaleMap[] =
    {
        // The Following must be sorted on m_sLang.
        {L".ARA", L"ar_ae"},
        {L".BGR", L"bg_bg"},
        {L".CHS", L"zh_cn"},
        {L".CHT", L"zh_tw"},
        {L".CZE", L"cs_cz"},
        {L".DAN", L"da_dk"},
        {L".DEU", L"de_de"},
        {L".ENU", L"en_us"},
        {L".ESP", L"es_es"},
        {L".ETI", L"et_ee"},
        {L".FRA", L"fr_fr"},
        {L".GRE", L"el_gr"},
        {L".HEB", L"he_il"},
        {L".HRV", L"hr_hr"},
        {L".HUN", L"hu_hu"},
        {L".ITA", L"it_it"},
        {L".JPN", L"ja_jp"},
        {L".KOR", L"ko_kr"},
        {L".LTH", L"lt_lt"},
        {L".LVI", L"lv_lv"},
        {L".MEA", L"en_AE"},
        {L".MEH", L"en_IL"},
        {L".NAF", L"fr_MA"},
        {L".NLD", L"nl_nl"},
        {L".NON", L"no_no"},
        {L".NOR", L"nb_no"},
        {L".POL", L"pl_pl"},
        {L".PTB", L"pt_br"},
        {L".RUM", L"ro_ro"},
        {L".RUS", L"ru_ru"},
        {L".SKY", L"sk_sk"},
        {L".SLV", L"sl_si"},
        {L".SUO", L"fi_fi"},
        {L".SVE", L"sv_se"},
        {L".TUR", L"tr_tr"},
        {L".UKR", L"uk_ua"}
    };
    long uBound = (sizeof(m_langToLocaleMap) / sizeof(SLangToLocaleMap)) - 1, lBound = 0, nMid;
    int nLocation;
    while (lBound <= uBound)
    {
        nMid = (lBound + uBound) / 2;
        nLocation = _tcscmp(sLang, (m_langToLocaleMap[nMid].psLang));
        if (nLocation == 0)
        {
            return m_langToLocaleMap[nMid].psLocale;
        }
        else if (nLocation < 0)
        {
            uBound = nMid - 1;
        }
        else
        {
            lBound = nMid + 1;
        }
    }
    return(L"en_us");
}

HINSTANCE GetResourceDll(HINSTANCE hAdist)
{
    HINSTANCE resourcedllhandle;
    TCHAR* pExt, * pName, resourceDllName[MAX_PATH], fName[_MAX_FNAME];
    TCHAR szLocale[MAX_PATH] = { 0 };
    if (hAdist)
        GetModuleFileName(hAdist, resourceDllName, sizeof(resourceDllName));
    else {
        CRegKey cReg;
        if (cReg.Open(HKEY_LOCAL_MACHINE, DISTILLER_PATH_KEY, KEY_READ) == ERROR_SUCCESS)
        {
            ULONG lValue = sizeof(resourceDllName);
            cReg.QueryStringValue(NULL, resourceDllName, &lValue);
            wcscat(resourceDllName, L"\\..\\Acrobat Elements\\");
            wcscat(resourceDllName, CONTEXT_MENU_DLL);
            hAdist = GetModuleHandle(resourceDllName);
            if (!hAdist)
                hAdist = LoadLibrary(resourceDllName);
        }
    }

    pExt = _tcsrchr(resourceDllName, '.');

    if (pExt && (_wcsicmp(pExt, L".EXE") == 0 ||
        _wcsicmp(pExt, L".DLL") == 0))
    {
        TCHAR szLanguage[10];
        GetAcrobatLanguage(szLanguage, szLocale);
        if (_tcsicmp(szLanguage, L".EXE") == 0 ||
            _tcsicmp(szLanguage, L".DLL") == 0 ||
            _tcsicmp(szLanguage, L".ENU") == 0)
            return hAdist; //return for English

        _tsplitpath(resourceDllName, NULL, NULL, fName, NULL);
        pName = _tcsrchr(resourceDllName, '\\');
        if (_tcslen(szLocale) > 0)
        {
            _tcscpy(pName, LOCALE_PATH_1);
            wcscat(pName, szLocale);
            wcscat(pName, LOCALE_PATH_2);
        }
        else
        {
            _tcscpy(pName, LOCALE_PATH_3);
            wcscat(pName, GetLocale(szLanguage));
            wcscat(pName, LOCALE_PATH_4);
        }
        wcscat(pName, fName);
        wcscat(pName, szLanguage);
        /*_tcscpy(pExt,szLanguage);*/
    }
    else
        return hAdist;

    // load the language resource
    resourcedllhandle = LoadLibrary(resourceDllName);
    if (!resourcedllhandle)
    {   // if we don't find the Language resource use English
        if ((GetKeyState(VK_RCONTROL) & 0x8000) && (GetKeyState(VK_RSHIFT) & 0x8000))
        {
            DWORD err = GetLastError();
            TCHAR buf[256];
            swprintf(buf,sizeof(buf), L"LoadLibrary(%s) Returned 0 :GetLastError : %d ", resourceDllName, err);
            MessageBox(GetForegroundWindow(), buf, L"SideCar file Error", MB_ICONEXCLAMATION | MB_OK);
        }
        resourcedllhandle = hAdist;
    }
    return resourcedllhandle;
}

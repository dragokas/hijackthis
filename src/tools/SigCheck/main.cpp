#include <iostream>
#include <stdio.h>
#include <windows.h>
#include <Softpub.h>
#include <wincrypt.h>
#include <wintrust.h>
#include <mscat.h>
#include <versionhelpers.h>
#include <vector>
#include <shlwapi.h>

#pragma comment (lib, "wintrust")
#pragma comment (lib, "Crypt32")
#pragma comment (lib, "Shlwapi.lib")

struct SignResult
{
    std::wstring HashFinalCert;
    std::wstring SubjectName;
};

PVOID m_RedirectorOldValue = NULL;

void PrintLastError(const WCHAR* msg)
{
    LPWSTR str = NULL;
    if (FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL, 
        GetLastError(),
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // or MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US) to use English
        (LPWSTR)&str, 
        0,
        NULL))
    {
        while (wchar_t* pFound = wcsstr(str, L"\r\n"))
        {
            *pFound = L' ';
            *(pFound+1) = L' ';
        }
        wprintf_s(L"%s: %s\n", msg, str);
        LocalFree(str);
    }
}

bool ToggleFileSystemRedirector(bool bEnable, LPCWSTR path = NULL)
{
    static bool IsNotRedirected;
    static bool IsInit;
    static BOOL IsWow;
    static WCHAR WinSysDir[MAX_PATH] = L"\0";
    static size_t cchWinSysDir;

    if (!IsInit)
    {
        IsInit = true;
        IsWow64Process(GetCurrentProcess(), &IsWow);

        if (ExpandEnvironmentStrings(L"%SystemRoot%", WinSysDir, MAX_PATH))
        {
            wcscat_s(WinSysDir, MAX_PATH, L"\\system32");
            cchWinSysDir = wcsnlen_s(WinSysDir, MAX_PATH);
        }
    }
    if (!IsWow)
    {
        return false;
    }
    if (path != NULL)
    {
        if (_wcsnicmp(path, WinSysDir, cchWinSysDir) != 0)
        {
            return false;
        }
    }
    if (bEnable)
    {
        if (IsNotRedirected)
        {
            IsNotRedirected = false;
            return Wow64RevertWow64FsRedirection(m_RedirectorOldValue);
        }
    }
    else {
        if (!IsNotRedirected)
        {
            IsNotRedirected = true;
            return Wow64DisableWow64FsRedirection(&m_RedirectorOldValue);
        }
    }
    return false;
}

std::wstring ExtractStringFromCertificate(PCCERT_CONTEXT pCertificate, DWORD dwType, DWORD dwFlags = 0)
{
    DWORD size = CertGetNameString(pCertificate, dwType, dwFlags, NULL, NULL, NULL);
    if (size)
    {
        std::vector<wchar_t> buff(size);
        CertGetNameString(pCertificate, dwType, dwFlags, NULL, (LPWSTR)&buff[0], size);
        std::wstring result(buff.begin(), buff.end());
        return result;
    }
    return L"";
}

void GetSignerInfo(HANDLE hWVTStateData, SignResult &SignResult)
{
    CRYPT_PROVIDER_DATA *pProvData = WTHelperProvDataFromStateData(hWVTStateData);
    if (pProvData != NULL)
    {
        int idxSigner = 0;
        CRYPT_PROVIDER_SGNR *pCPSigner = WTHelperGetProvSignerFromChain(pProvData, idxSigner, FALSE, 0);
        if (pCPSigner != NULL)
        {
            PCCERT_CONTEXT pCertificate = CertDuplicateCertificateContext(pCPSigner->pasCertChain->pCert);
            if (pCertificate != NULL)
            {
                DWORD size = 0;
                CertGetCertificateContextProperty(pCertificate, CERT_HASH_PROP_ID, NULL, &size);
                if (size != 0)
                {
                    std::vector<uint8_t> buff(size);
                    if (CertGetCertificateContextProperty(pCertificate, CERT_HASH_PROP_ID, &buff.front(), &size))
                    {
                        WCHAR pszMemberTag[100] = { 0 };
                        for (DWORD i = 0; i < size; ++i)
                        {
                            wsprintfW(&pszMemberTag[i * 2], L"%02X", buff[i]);
                        }
                        SignResult.HashFinalCert = pszMemberTag;
                    }
                }

                SignResult.SubjectName = ExtractStringFromCertificate(pCertificate, CERT_NAME_SIMPLE_DISPLAY_TYPE, 0);

                CertFreeCertificateContext(pCertificate);
            }
        }
    }
}



bool VerifySignature(LPCWSTR lpFileName, SignResult &signResult)
{
    BOOL bRet = FALSE, bIsVerified = FALSE;
    WINTRUST_DATA wd = { 0 };
    WINTRUST_FILE_INFO wfi = { 0 };
    WINTRUST_CATALOG_INFO wci = { 0 };
    CATALOG_INFO catalogInfo = { 0 };
    WCHAR pszMemberTag[260] = { 0 };
    HCATINFO hCatInfoContext = NULL;
    LONG iResult = 0;
    DRIVER_VER_INFO verInfo = { 0 };
    WINTRUST_SIGNATURE_SETTINGS signSettings = { 0 };
    HCATADMIN hCatAdmin = NULL;

    // redirector OFF
    ToggleFileSystemRedirector(false, lpFileName);

    HANDLE hFile = CreateFileW(lpFileName, 
        FILE_READ_ATTRIBUTES | FILE_READ_DATA | STANDARD_RIGHTS_READ, 
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    
    // redirector ON
    ToggleFileSystemRedirector(true, NULL); 

    if (INVALID_HANDLE_VALUE == hFile)
    {
        PrintLastError(L"Error in CreateFileW");
        return FALSE;
    }

    GUID DriverGuid = DRIVER_ACTION_VERIFY;
    GUID VerifyGuid = WINTRUST_ACTION_GENERIC_VERIFY_V2;
    if (IsWindows8OrGreater())
    {
        CryptCATAdminAcquireContext2(&hCatAdmin, &DriverGuid, BCRYPT_SHA256_ALGORITHM, NULL, 0);
        if (hCatAdmin != NULL)
        {
            wprintf_s(L"Used CryptCATAdminAcquireContext2\n");
        }
    }
    if (hCatAdmin == NULL)
    {
        if (!CryptCATAdminAcquireContext(&hCatAdmin, &DriverGuid, 0))
        {
            PrintLastError(L"Error in CryptCATAdminAcquireContext");
            return FALSE;
        }
    }

    DWORD dwHashSize = 0;
    std::vector<uint8_t> fileHash;
    if (IsWindows8OrGreater())
    {
        bRet = CryptCATAdminCalcHashFromFileHandle2(hCatAdmin, hFile, &dwHashSize, NULL, 0);
        
        if (GetLastError() == ERROR_INSUFFICIENT_BUFFER)
        {
            fileHash.resize(dwHashSize);
            bRet = CryptCATAdminCalcHashFromFileHandle2(hCatAdmin, hFile, &dwHashSize, &fileHash.front(), 0);
            if (bRet)
            {
                wprintf_s(L"Used CryptCATAdminCalcHashFromFileHandle2\n");
            }
        }
    }
    if (!bRet || dwHashSize == 0)
    {
        bRet = CryptCATAdminCalcHashFromFileHandle(hFile, &dwHashSize, NULL, 0);

        if (GetLastError() == ERROR_INSUFFICIENT_BUFFER)
        {
            fileHash.resize(dwHashSize);
            bRet = CryptCATAdminCalcHashFromFileHandle(hFile, &dwHashSize, &fileHash.front(), 0);
        }
        if (!bRet)
        {
            PrintLastError(L"Error in CryptCATAdminCalcHashFromFileHandle");
            goto Finalize;
        }
    }

    for (DWORD i = 0; i < dwHashSize; ++i)
    {
        wsprintfW(&pszMemberTag[i * 2], L"%02X", fileHash[i]);
    }

    hCatInfoContext = CryptCATAdminEnumCatalogFromHash(hCatAdmin, &fileHash.front(), dwHashSize, 0, NULL);

    if (hCatInfoContext != NULL)
    {
        catalogInfo.cbStruct = sizeof(CATALOG_INFO);
        if (!CryptCATCatalogInfoFromContext(hCatInfoContext, &catalogInfo, 0))
        {
            CryptCATAdminReleaseCatalogContext(hCatAdmin, hCatInfoContext, 0);
            hCatInfoContext = NULL;
        }
    }

    if (IsWindows8OrGreater())
    {
        wd.cbStruct = sizeof(WINTRUST_DATA);
    }
    else {
        wd.cbStruct = sizeof(WINTRUST_DATA) - sizeof(void*);
    }
    wd.dwUIChoice = WTD_UI_NONE;
    wd.dwStateAction = WTD_STATEACTION_VERIFY;
    wd.fdwRevocationChecks = WTD_REVOKE_NONE;
    wd.dwProvFlags = IsWindowsVistaOrGreater() ? WTD_CACHE_ONLY_URL_RETRIEVAL : WTD_REVOCATION_CHECK_NONE;
    wd.dwProvFlags |= WTD_SAFER_FLAG;
    wd.hWVTStateData = NULL;
    wd.pwszURLReference = NULL;

    if (hCatInfoContext != NULL)
    {
        verInfo.cbStruct = sizeof(DRIVER_VER_INFO);

        wd.pPolicyCallbackData = &verInfo;
        wd.dwUnionChoice = WTD_CHOICE_CATALOG;
        wd.pCatalog = &wci;
        wd.dwUIContext = WTD_UICONTEXT_EXECUTE;
        
        wci.cbStruct = sizeof(WINTRUST_CATALOG_INFO);
        wci.pcwszCatalogFilePath = catalogInfo.wszCatalogFile;
        wci.pcwszMemberTag = pszMemberTag;
        wci.pcwszMemberFilePath = lpFileName;
        wci.hMemberFile = hFile;
        wci.pbCalculatedFileHash = &fileHash.front();
        wci.cbCalculatedFileHash = dwHashSize;
        wci.hCatAdmin = hCatAdmin;

        wprintf_s(L"Verified by Catalogue.\n");
    } 
    else {
        wd.dwUnionChoice = WTD_CHOICE_FILE;
        wd.pFile = &wfi;
        
        wfi.cbStruct = sizeof(WINTRUST_FILE_INFO);
        wfi.pcwszFilePath = NULL;
        wfi.hFile = hFile;
        wfi.pgKnownSubject = NULL;

        if (IsWindows8OrGreater())
        {
            signSettings.cbStruct = sizeof(signSettings);
            signSettings.pCryptoPolicy = NULL;
            signSettings.dwFlags = WSS_GET_SECONDARY_SIG_COUNT;
            wd.pSignatureSettings = &signSettings;
        }

        wprintf_s(L"Verified Internal.\n");

        // To check specific signature index (in double-signature file):
        //signSettings.dwFlags = WSS_VERIFY_SPECIFIC;
        //signSettings.dwIndex = ... 0, 1
    }

    iResult = WinVerifyTrust((HWND)INVALID_HANDLE_VALUE, &VerifyGuid, &wd);
    wprintf_s(L"WinVerifyTrust returned: 0x%X (err: 0x%X)\n", iResult, GetLastError());
    bIsVerified = TRUE;

    if (iResult == 0 || iResult == CERT_E_UNTRUSTEDROOT || iResult == CERT_E_EXPIRED)
    {
        GetSignerInfo(wd.hWVTStateData, signResult);
    }

Finalize:

    if (hCatAdmin && hCatInfoContext)
        CryptCATAdminReleaseCatalogContext(hCatAdmin, hCatInfoContext, 0);

    if (bIsVerified)
    {
        wd.dwStateAction = WTD_STATEACTION_CLOSE;
        WinVerifyTrust((HWND)INVALID_HANDLE_VALUE, &VerifyGuid, &wd);
    }

    if (hCatAdmin)
        CryptCATAdminReleaseContext(hCatAdmin, 0);

    CloseHandle(hFile);

    return iResult == 0 || iResult == CERT_E_UNTRUSTEDROOT || iResult == CERT_E_EXPIRED;
}

bool IsDragoCert(std::wstring certHash)
{
    return (0 == certHash.compare(L"05F1F2D5BA84CDD6866B37AB342969515E3D912E"));
}

bool FileExists(LPCWSTR file)
{
    DWORD attr = GetFileAttributesW(file);
    return attr != INVALID_FILE_ATTRIBUTES && ((attr & FILE_ATTRIBUTE_DIRECTORY) == 0);
}

int wmain(int argc, wchar_t* argv[], wchar_t* envp[])
{
    SignResult signResult;
    //WCHAR file[MAX_PATH] = L"";
    
    if (argc <= 1)
    {
        wchar_t* name = _wcsdup(argv[0]);
        PathStripPathW(name);
        wprintf_s(L"Using:\n");
        wprintf_s(L"\"%s\" <file>\n", name);
        return 1;
    }

    if (!FileExists(argv[1]))
    {
        wprintf_s(L"File not found: %s\n", argv[1]);
        return 1;
    }

    if (VerifySignature(argv[1], signResult))
    {
        wprintf_s(L"Cert Hash: %s\n", signResult.HashFinalCert.c_str());
        wprintf_s(L"Publisher: %s\n", signResult.SubjectName.c_str());
    }
    else wprintf_s(L"Signature: Failed!\n");

    return 0;
}


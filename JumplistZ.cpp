/*
JumplistZ
https://github.com/zackz/JumplistZ
*/

#define UNICODE
#define _UNICODE

#include <stdio.h>
#include <tchar.h>
#include <windows.h>
#include <objbase.h>
#include <Shlobj.h>
#include <shobjidl.h>
#include <Knownfolders.h>
#include <propkey.h>
#include <Propvarutil.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")

const TCHAR NAME[]                = _T("JumplistZ");
const TCHAR VERSION[]             = _T("0.5.1");
const TCHAR CFGKEY_DEBUG_BITS[]   = _T("DEBUG_BITS");
const TCHAR CFGKEY_GROUP_NAME[]   = _T("GROUP_DISPLAY_NAME");
const TCHAR SECTION_PROPERTIES[]  = _T("PROPERTIES");
const TCHAR SECTION_PREFIX[]      = _T("GROUP");
const TCHAR ITEM_PREFIX[]         = _T("ITEM");
const TCHAR ITEM_SUFFIX_NAME[]    = _T("_NAME");
const TCHAR ITEM_SUFFIX_CMD[]     = _T("_CMD");
const int   CFG_MAX_COUNT         = 100;
const int   CFG_VALUE_LEN         = 1024;

TCHAR g_szAppName[MAX_PATH];
DWORD g_dwDebugBits = 0;

void dbg(LPCTSTR szFormat, ...)
{
	if (g_dwDebugBits == 0)
		return;
	TCHAR buf[1024];
	va_list args;
	va_start(args, szFormat);
	_vsntprintf(buf, ARRAYSIZE(buf), szFormat, args);
	va_end(args);
	buf[ARRAYSIZE(buf) - 1] = 0;

	if (g_dwDebugBits & 1)
		_tprintf(_T("%s\n"), buf);
	if (g_dwDebugBits & 2)
		OutputDebugString(buf);
}

void err(LPCTSTR szFormat, ...)
{
	TCHAR buf[1024];
	va_list args;
	va_start(args, szFormat);
	_vsntprintf(buf, ARRAYSIZE(buf), szFormat, args);
	va_end(args);
	buf[ARRAYSIZE(buf) - 1] = 0;
	MessageBox(NULL, buf, g_szAppName, MB_OK | MB_ICONWARNING);
}

void GetCFGItem(DWORD nSection, DWORD nItem,
	LPCTSTR szSuffix, TCHAR * szINI, LPTSTR bufValue)
{
	TCHAR szSection[100], szKey[100];
	_stprintf(szSection, _T("%s%d"), SECTION_PREFIX, nSection);
	_stprintf(szKey, _T("%s%d%s"), ITEM_PREFIX, nItem, szSuffix);
	if (NULL == GetPrivateProfileString(szSection, szKey,
		NULL, bufValue, CFG_VALUE_LEN, szINI))
	{
		bufValue[0] = 0;
	}
}

void SplitFileAndParameters(LPCTSTR szCMD, LPTSTR bufFile, LPTSTR bufParam)
{
	const TCHAR * pFile = szCMD;
	while (_tcschr(_T(" \t"), *pFile))
		pFile++;
	const TCHAR * pParam = pFile;
	if (*pParam == _T('"'))
	{
		pFile++;
		pParam++;
		while (*pParam && *pParam != _T('"'))
			pParam++;
	}
	else
	{
		while (*pParam && NULL == _tcschr(_T(" \t"), *pParam))
			pParam++;
	}
	if (*pParam)
	{
		_tcsncpy(bufFile, pFile, pParam - pFile);
		bufFile[pParam - pFile] = 0;
		_tcscpy(bufParam, pParam + 1);
	}
	else
	{
		_tcscpy(bufFile, pFile);
		bufParam[0] = 0;
	}
}

HRESULT GetDefaultAssocPath(LPCTSTR szExtra, LPTSTR bufPath)
{
	bufPath[0] = 0;
	IEnumAssocHandlers * peah;
	HRESULT hr = SHAssocEnumHandlers(szExtra, ASSOC_FILTER_RECOMMENDED, &peah);
	if (SUCCEEDED(hr))
	{
		IAssocHandler * pah;
		hr = peah->Next(1, &pah, NULL);
		if (SUCCEEDED(hr))
		{
			LPWSTR name;
			hr = pah->GetName(&name);
			if (SUCCEEDED(hr))
			{
				_tcscpy(bufPath, name);
				CoTaskMemFree(name);
			}
			pah->Release();
		}
		peah->Release();
	}
	return hr;
}

HRESULT SetTitle(IShellLink * psl, LPCTSTR szTitle)
{
	IPropertyStore * pps;
	HRESULT hr = psl->QueryInterface(IID_PPV_ARGS(&pps));
	if (SUCCEEDED(hr))
	{
		PROPVARIANT pv;
		hr = InitPropVariantFromString(szTitle, &pv);
		if (SUCCEEDED(hr))
		{
			hr = pps->SetValue(PKEY_Title, pv);
			if (SUCCEEDED(hr))
				hr = pps->Commit();
			PropVariantClear(&pv);
		}
		pps->Release();
	}
	return hr;
}

IShellLink * GetShellLink(LPCTSTR szName, LPCTSTR szCMD)
{
	TCHAR bufFile[CFG_VALUE_LEN];
	TCHAR bufParam[CFG_VALUE_LEN];
	SplitFileAndParameters(szCMD, bufFile, bufParam);
	dbg(_T("-Name: <%s>"), szName);
	dbg(_T("-CMD:  <%s>, <%s>"), bufFile, bufParam);

	HRESULT hr = 0;
	while (TRUE)
	{
		IShellLink * psl;
		hr = CoCreateInstance(
			CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psl));
		if (FAILED(hr)) break;

		hr = psl->SetPath(bufFile);
		if (FAILED(hr)) break;

		LPCTSTR prefixes[] = {_T("http://"), _T("https://")};
		int i = 0;
		for (; i < ARRAYSIZE(prefixes); i++)
			if (0 == _tcsncicmp(bufFile, prefixes[i], sizeof(prefixes[i])))
				break;
		if (i < ARRAYSIZE(prefixes))
		{
			TCHAR bufPath[MAX_PATH];
			hr = GetDefaultAssocPath(_T(".html"), bufPath);
			if (SUCCEEDED(hr))
			{
				hr = psl->SetIconLocation(bufPath, 0);
				if (FAILED(hr)) break;
			}
		}
		else
		{
			hr = psl->SetIconLocation(bufFile, 0);
			if (FAILED(hr)) break;
		}

		hr = psl->SetArguments(bufParam);
		if (FAILED(hr)) break;

		hr = SetTitle(psl, szName);
		if (FAILED(hr)) break;

		hr = psl->SetDescription(szCMD);
		if (FAILED(hr)) break;

		TCHAR bufPath[MAX_PATH];
		if (SHGetSpecialFolderPath(0, bufPath, CSIDL_PROFILE, FALSE))
		{
			hr = psl->SetWorkingDirectory(bufPath);
			if (FAILED(hr)) break;
		}
		return psl;
	}
	dbg(_T("Error call GetShellItem_link, hr: 0x%08x"), hr);
	return NULL;
}

IShellLink * GetShellLink(DWORD nSection, DWORD nItem, TCHAR * szINI)
{
	TCHAR bufCMD[CFG_VALUE_LEN];
	GetCFGItem(nSection, nItem, ITEM_SUFFIX_CMD, szINI, bufCMD);
	if (_tcslen(bufCMD) == 0)
		return NULL;
	TCHAR bufName[CFG_VALUE_LEN];
	GetCFGItem(nSection, nItem, ITEM_SUFFIX_NAME, szINI, bufName);
	dbg(_T("GetShellLink, section: %d, item: %d"), nSection, nItem);
	return GetShellLink(bufName, bufCMD);
}

int AddGroup(ICustomDestinationList * pcdl, DWORD nSection, TCHAR * szINI)
{
	TCHAR szSection[100], bufGroupName[100];
	_stprintf(szSection, _T("%s%d"), SECTION_PREFIX, nSection);
	if (NULL == GetPrivateProfileString(szSection, CFGKEY_GROUP_NAME,
		NULL, bufGroupName, ARRAYSIZE(bufGroupName), szINI))
	{
		bufGroupName[0] = 0;
	}
	if (_tcslen(bufGroupName) == 0)
		return 0;
	dbg(_T("AddGroup %s, Name: <%s>"), szSection, bufGroupName);

	IObjectCollection * poc;
	HRESULT hr = CoCreateInstance(CLSID_EnumerableObjectCollection,
		NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&poc));
	if (FAILED(hr))
		return 0;

	for (int k = 1; k < CFG_MAX_COUNT; k++)
	{
		IShellLink * psi = GetShellLink(nSection, k, szINI);
		if (psi)
		{
			dbg(_T("Shell object: 0x%08x"), psi);
			poc->AddObject(psi);
			psi->Release();
		}
	}

	// Append category
	UINT nObjects = 0;
	IObjectArray * poa;
	hr = poc->QueryInterface(IID_PPV_ARGS(&poa));
	if (SUCCEEDED(hr))
	{
		hr = pcdl->AppendCategory(bufGroupName, poa);
		if (FAILED(hr))
			dbg(_T("Error call AppendCategory, %s, hr: 0x%08x"), bufGroupName, hr);
		hr = poa->GetCount(&nObjects);
		poa->Release();
	}
	poc->Release();
	return nObjects;
}

void AddTasks(ICustomDestinationList * pcdl, TCHAR * szINI)
{
	IObjectCollection * poc;
	HRESULT hr = CoCreateInstance(CLSID_EnumerableObjectCollection,
		NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&poc));
	if (SUCCEEDED(hr))
	{
		TCHAR bufPath[MAX_PATH];
		GetDefaultAssocPath(_T(".ini"), bufPath);
		TCHAR bufCMD[MAX_PATH];
		_stprintf(bufCMD, _T("%s %s"), bufPath, szINI);

		IShellLink * psl;
		psl = GetShellLink(_T("Edit configuration"), bufCMD);
		if (psl)
			poc->AddObject(psl);
		psl = GetShellLink(
			_T("About JumplistZ"), _T("https://github.com/zackz/JumplistZ"));
		if (psl)
			poc->AddObject(psl);
		pcdl->AddUserTasks(poc);
	}
}

int BuildJumplist(TCHAR * szINI)
{
	dbg(_T("- BuildJumplist"));

	ICustomDestinationList * pcdl;
	HRESULT hr = CoCreateInstance(
		CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pcdl));
	if (FAILED(hr))
	{
		err(_T("Error call CoCreateInstance, CLSID_DestinationList, hr = 0x%08x"), hr);
		return 0;
	}

	UINT uMaxSlots;
	IObjectArray * poaRemoved;  // Not care, should remove it in INI file.
	hr = pcdl->BeginList(&uMaxSlots, IID_PPV_ARGS(&poaRemoved));
	if (FAILED(hr))
	{
		err(_T("Error call BeginList, hr = 0x%08x"), hr);
		pcdl->Release();
		return 0;
	}

	// Category
	int nObjects = 0;
	for (int i = 1; i < CFG_MAX_COUNT; i++)
		nObjects += AddGroup(pcdl, i, szINI);

	// Tasks
	AddTasks(pcdl, szINI);

	pcdl->CommitList();
	poaRemoved->Release();
	pcdl->Release();
	return nObjects;
}

int CALLBACK WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
	LPSTR lpCmdLine, int nCmdShow)
{
	TCHAR bufFile[CFG_VALUE_LEN];
	TCHAR bufParam[CFG_VALUE_LEN];
	SplitFileAndParameters(GetCommandLine(), bufFile, bufParam);

	// Get ini full path
	TCHAR szINI[MAX_PATH];
	if (_tcsrchr(bufFile, _T('\\')) != 0)
	{
		_tcscpy(szINI, bufFile);
	}
	else
	{
		_tgetcwd(szINI, ARRAYSIZE(szINI));
		if (szINI[_tcslen(szINI) - 1] != _T('\\'))
			_tcscat(szINI, _T("\\"));
		_tcscat(szINI, bufFile);
		_tcscpy(bufFile, szINI);  // Full path
	}
	*_tcsrchr(szINI, '.') = 0;  // Remove ".exe"
	_tcscat(szINI, _T(".ini"));

	// Title of MessageBox, "JumplistZ X.X.X"
	_stprintf(g_szAppName, _T("%s %s"), NAME, VERSION);

	// Not exists INI?
	WIN32_FIND_DATA wfd;
	if (INVALID_HANDLE_VALUE == FindFirstFile(szINI, &wfd))
	{
		if (IDYES != MessageBox(NULL,
			_T("Can't find ini file. Do you want to create a sample one?"),
			g_szAppName, MB_YESNO | MB_ICONQUESTION))
		{
			return 0;
		}
		FILE * f = _tfopen(szINI, _T("wb"));
		if (!f)
		{
			err(_T("Can't create sample file!\n%s"), szINI);
			return 0;
		}
		HRSRC hs = FindResource(NULL, _T("SMAPLEINI"), _T("BIN"));
		if (!hs)
		{
			err(_T("Can't find sample file in resource!"));
			fclose(f);
			return 0;
		}
		LPVOID pData = LockResource(LoadResource(NULL, hs));
		fwrite(pData, 1, SizeofResource(NULL, hs), f);
		fclose(f);
	}

	// Read configuration
	g_dwDebugBits = GetPrivateProfileInt(
		SECTION_PROPERTIES, CFGKEY_DEBUG_BITS, 0, szINI);

	// Initialize jumplist
	if (SUCCEEDED(CoInitialize(NULL)))
	{
		int nObjects = BuildJumplist(szINI);
		CoUninitialize();

		// Result message. App path is included in message because defferent path
		// leads to different jumplist even thought the executable file is same.
		TCHAR buf[100 + MAX_PATH];
		_stprintf(buf, _T("Updated jumplist. Total %d items.\nPath: \"%s\""),
			nObjects, bufFile);
		MessageBox(NULL, buf, g_szAppName, MB_OK | MB_ICONINFORMATION);
	}
	return 0;
}



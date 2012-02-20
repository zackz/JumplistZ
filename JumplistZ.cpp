
#define UNICODE
#define _UNICODE

#include <stdio.h>
#include <tchar.h>
#include <windows.h>
#include <shobjidl.h>
#include <objbase.h>
#include <Shlobj.h>
#include <Knownfolders.h>
#include <Propvarutil.h>
#include <propkey.h>

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")

REFKNOWNFOLDERID TEMP_FOLDERID = FOLDERID_Templates;
const TCHAR TEMP_FILE_PREFIX[] = _T("ZZZzzzzzzz");

const int CFG_MAX_COUNT = 100;
const int CFG_VALUE_LEN = 1024;
const TCHAR CFGKEY_DEBUG_BITS[] = _T("DEBUG_BITS");
const TCHAR CFGKEY_GROUP_NAME[] = _T("GROUP_DISPLAY_NAME");
const TCHAR SECTION_PROPERTIES[] = _T("PROPERTIES");
const TCHAR SECTION_PREFIX[] = _T("GROUP");
const TCHAR ITEM_PREFIX[] = _T("ITEM");
const TCHAR ITEM_SUFFIX_NAME[] = _T("_NAME");
const TCHAR ITEM_SUFFIX_CMD[] = _T("_CMD");

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
	MessageBox(NULL, buf, g_szAppName, MB_OK);
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

BOOL FindStrInList(TCHAR * str, TCHAR * plist)
{
	// The list includes multiple sequence strings, and the last one is empty.
	for (TCHAR * pstr = plist; *pstr; pstr += _tcslen(pstr) + 1)
	{
		if (_tcsstr(pstr, str))
			return TRUE;
	}
	return FALSE;
}

void GetTempFileName(TCHAR * bufRet, DWORD nSection, DWORD nItem)
{
	_stprintf(bufRet, _T("%s_%s_%02d%02d.JumplistZ"),
		TEMP_FILE_PREFIX, g_szAppName, nSection, nItem);
}

int GetTempNamePrefixLen()
{
	// Length of "ZZZzzzzzzz_JumplistZ_"
	return _tcslen(TEMP_FILE_PREFIX) + _tcslen(g_szAppName) + 2;
}

IShellItem * GetShellItem(DWORD nSection, DWORD nItem, TCHAR * szINI)
{
	TCHAR bufCMD[CFG_VALUE_LEN];
	GetCFGItem(nSection, nItem, ITEM_SUFFIX_CMD, szINI, bufCMD);
	if (_tcslen(bufCMD) == 0)
		return NULL;

	TCHAR bufName[MAX_PATH];
	GetTempFileName(bufName, nSection, nItem);

	FILE * f = _tfopen(bufName, _T("wb"));
	fwrite(bufCMD, sizeof(bufCMD[0]), _tcslen(bufCMD) + 1, f);
	fclose(f);

	IShellItem * psi;
	HRESULT hr = SHCreateItemInKnownFolder(
		TEMP_FOLDERID, KF_FLAG_DEFAULT, bufName, IID_PPV_ARGS(&psi));
	if (SUCCEEDED(hr))
		return psi;
	return NULL;
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

IShellLink * GetShellLink(DWORD nSection, DWORD nItem, TCHAR * szINI)
{
	TCHAR bufCMD[CFG_VALUE_LEN];
	GetCFGItem(nSection, nItem, ITEM_SUFFIX_CMD, szINI, bufCMD);
	if (_tcslen(bufCMD) == 0)
		return NULL;
	TCHAR bufName[CFG_VALUE_LEN];
	GetCFGItem(nSection, nItem, ITEM_SUFFIX_NAME, szINI, bufName);
	dbg(_T("GetShellLink, section: %d, item: %d"), nSection, nItem);

	TCHAR bufFile[CFG_VALUE_LEN];
	TCHAR bufParam[CFG_VALUE_LEN];
	SplitFileAndParameters(bufCMD, bufFile, bufParam);
	dbg(_T("-Name: <%s>"), bufName);
	dbg(_T("-CMD:  <%s>, <%s>"), bufFile, bufParam);

	HRESULT hr = 0;
	while (TRUE)
	{
		IShellLink * psl;
		hr = CoCreateInstance(
			CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psl));
		if(!SUCCEEDED(hr)) break;

		hr = psl->SetPath(bufFile);
		if(!SUCCEEDED(hr)) break;

		hr = psl->SetIconLocation(bufFile, 0);
		if(!SUCCEEDED(hr)) break;

		hr = psl->SetArguments(bufParam);
		if(!SUCCEEDED(hr)) break;

		hr = SetTitle(psl, bufName);
		if(!SUCCEEDED(hr)) break;

		hr = psl->SetDescription(bufCMD);
		if(!SUCCEEDED(hr)) break;

		TCHAR bufPath[MAX_PATH];
		if (SHGetSpecialFolderPath(0, bufPath, CSIDL_PROFILE, FALSE))
		{
			hr = psl->SetWorkingDirectory(bufPath);
			if(!SUCCEEDED(hr)) break;
		}
		return psl;
	}
	dbg(_T("Error call GetShellItem_link, hr: 0x%08x"), hr);
	return NULL;
}

void AddGroup(ICustomDestinationList * pcdl, DWORD nSection, TCHAR * szINI)
{
	TCHAR szSection[100], bufGroupName[100];
	_stprintf(szSection, _T("%s%d"), SECTION_PREFIX, nSection);
	if (NULL == GetPrivateProfileString(szSection, CFGKEY_GROUP_NAME,
		NULL, bufGroupName, ARRAYSIZE(bufGroupName), szINI))
	{
		bufGroupName[0] = 0;
	}
	if (_tcslen(bufGroupName) == 0)
		return;
	dbg(_T("AddGroup %s, Name: <%s>"), szSection, bufGroupName);

	IObjectCollection * poc;
	HRESULT hr = CoCreateInstance(CLSID_EnumerableObjectCollection,
		NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&poc));
	if (!SUCCEEDED(hr))
		return;

	for (int k = 1; k < CFG_MAX_COUNT; k++)
	{
		IShellItem * psi = GetShellItem(nSection, k, szINI);
		//~ IShellLink * psi = GetShellLink(nSection, k, szINI);
		if (psi)
		{
			dbg(_T("Shell object: 0x%08x"), psi);
			poc->AddObject(psi);
			psi->Release();
		}
	}

	// Append category
	IObjectArray * poa;
	hr = poc->QueryInterface(IID_PPV_ARGS(&poa));
	if (SUCCEEDED(hr))
	{
		hr = pcdl->AppendCategory(bufGroupName, poa);
		if (!SUCCEEDED(hr))
			dbg(_T("Error call AppendCategory, %s, hr: 0x%08x"), bufGroupName, hr);
		poa->Release();
	}
	poc->Release();
}

void BuildJumplist(TCHAR * szINI)
{
	dbg(_T("- BuildJumplist"));

	ICustomDestinationList * pcdl;
	HRESULT hr = CoCreateInstance(
		CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pcdl));
	if (!SUCCEEDED(hr))
	{
		err(_T("Error call CoCreateInstance, CLSID_DestinationList, hr = 0x%08x"), hr);
		return;
	}

	UINT uMaxSlots;
	IObjectArray * poaRemoved;  // Not care, should remove it in INI file.
	hr = pcdl->BeginList(&uMaxSlots, IID_PPV_ARGS(&poaRemoved));
	if (!SUCCEEDED(hr))
	{
		err(_T("Error call BeginList, hr = 0x%08x"), hr);
		return;
	}

	PWSTR temppath;
	SHGetKnownFolderPath(TEMP_FOLDERID, KF_FLAG_DEFAULT, NULL, &temppath);
	_tchdir(temppath);
	CoTaskMemFree(temppath);

	for (int i = 1; i < CFG_MAX_COUNT; i++)
		AddGroup(pcdl, i, szINI);

	pcdl->CommitList();
	poaRemoved->Release();
	pcdl->Release();
}

LPCTSTR GetTimeStr(const FILETIME * lpFileTime)
{
	static TCHAR buf[100];
	SYSTEMTIME st;
	if (FileTimeToSystemTime(lpFileTime, &st))
	{
		_stprintf(buf, _T("%04d-%02d-%02d %02d:%02d:%02d.%03d"),
			st.wYear, st.wMonth, st.wDay,
			st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
		return buf;
	}
	return _T("");
}

BOOL FindLatestFile(LPCTSTR lpFileName, LPTSTR bufLatest)
{
	WIN32_FIND_DATA wfdLatest;
	memset(&wfdLatest, 0, sizeof(wfdLatest));
	WIN32_FIND_DATA wfd;
	HANDLE h = FindFirstFile(lpFileName, &wfd);
	while (h != INVALID_HANDLE_VALUE)
	{
		dbg(_T("%s, %s"), wfd.cFileName, GetTimeStr(&wfd.ftLastWriteTime));
		if (CompareFileTime(&wfd.ftLastWriteTime, &wfdLatest.ftLastWriteTime) == 1)
		{
			wfdLatest = wfd;
		}
		if (!FindNextFile(h, &wfd))
		{
			FindClose(h);
			break;
		}
	}
	dbg(_T("Latest one:"));
	dbg(_T("%s, %s"), wfdLatest.cFileName, GetTimeStr(&wfdLatest.ftLastWriteTime));
	_tcscpy(bufLatest, wfdLatest.cFileName);
	return TRUE;
}

void ChangeDisplayNames(TCHAR * szINI)
{
	dbg(_T("- ChangeDisplayNames"));

	// Find latest "*.customDestinations-ms". And assume it was just edited one.
	TCHAR szPath[MAX_PATH];
	PWSTR pathRecent;
	SHGetKnownFolderPath(FOLDERID_Recent, KF_FLAG_DEFAULT, NULL, &pathRecent);
	_tcscpy(szPath, pathRecent);
	CoTaskMemFree(pathRecent);
	if (szPath[_tcslen(szPath) - 1] != '\\')
		_tcscat(szPath, _T("\\"));
	_tcscat(szPath, _T("CustomDestinations\\*.customDestinations-ms"));

	TCHAR bufLatest[MAX_PATH];
	FindLatestFile(szPath, bufLatest);
	_tcscpy(_tcsrchr(szPath, _T('\\')) + 1, bufLatest);

	// Read all file data
	char buf[100 * 1024];
	memset(buf, 0, sizeof(buf));
	FILE * f = _tfopen(szPath, _T("r+b"));
	size_t nLen = fread(buf, 1, sizeof(buf), f);
	dbg(_T("Size: %d"), nLen);

	// Trick, binary-->string
	TCHAR bufPrefix[100];
	_tcscpy(bufPrefix, TEMP_FILE_PREFIX);
	for (char * p = (char *)bufPrefix; ((int)p % 2) || (*p || *(p + 1)); p++)
		*p = *p ? *p : 1;
	for (char * p = buf; p - buf < nLen; p++)
		*p = *p ? *p : 1;

	TCHAR * p = (TCHAR *)buf;
	while (TRUE)
	{
		// Find index of original display name
		TCHAR * p0 = _tcsstr(p, bufPrefix);
		TCHAR * p1 = _tcsstr((TCHAR *)((char *)p + 1), bufPrefix);
		p = min(p0, p1);
		if (!p)
			p = max(p0, p1);
		if (!p)
			break;

		// Get section & item index from display name
		// Such as "ZZZzzzzzzz_JumplistZ_0105.JumplistZ"
		char * pch = (char *)p + GetTempNamePrefixLen() * 2;
		int nSection = (*pch - '0') * 10 + (*(pch + 2) - '0');
		pch += 4;
		int nItem = (*pch - '0') * 10 + (*(pch + 2) - '0');

		// Get new display name
		TCHAR bufName[CFG_VALUE_LEN];;
		GetCFGItem(nSection, nItem, ITEM_SUFFIX_NAME, szINI, bufName);
		dbg(_T("Section: %2d, Item: %2d, New display name: <%s>"), nSection, nItem, bufName);

		// Write name to file
		if (0 == fseek(f, (char *)p - buf, SEEK_SET))
			fwrite(bufName, 1, (_tcslen(bufName) + 1) * sizeof(bufName[0]), f);

		p++;
	}
	fclose(f);
}

int wmain(int argc, wchar_t * argv[])
{
	// Get ini full path
	TCHAR * pch = _tcsrchr(argv[0], _T('\\'));
	_tcscpy(g_szAppName, pch ? pch + 1 : argv[0]);
	*_tcsrchr(g_szAppName, '.') = 0;  // Remove ".exe"

	TCHAR szINI[MAX_PATH];
	if (pch)
	{
		_tcsncpy_s(szINI, argv[0], pch - argv[0] + 1);
	}
	else
	{
		_tgetcwd(szINI, ARRAYSIZE(szINI));
		if (szINI[_tcslen(szINI) - 1] != _T('\\'))
			_tcscat(szINI, _T("\\"));
	}
	_tcscat(szINI, g_szAppName);
	_tcscat(szINI, _T(".ini"));

	// Read configuration
	g_dwDebugBits = GetPrivateProfileInt(SECTION_PROPERTIES, CFGKEY_DEBUG_BITS, 0, szINI);

	// Dump all argv
	dbg(_T("argc:    %d"), argc);
	for (int i = 0; i < argc; i++)
		dbg(_T("argv[%d]: %s"), i, argv[i]);

	if (argc == 1)
	{
		// Initialize jumplist
		if (SUCCEEDED(CoInitialize(NULL)))
		{
			BuildJumplist(szINI);
			ChangeDisplayNames(szINI);
			CoUninitialize();
		}
	}
	else
	{
		// Run...

		// Something wrong? Can't read argv[1], errno = 22
		// So read CMD from ini
		TCHAR * p = _tcsstr(argv[1], TEMP_FILE_PREFIX);
		if (p)
		{
			char * pch = (char *)p + GetTempNamePrefixLen() * 2;
			int nSection = (*pch - '0') * 10 + (*(pch + 2) - '0');
			pch += 4;
			int nItem = (*pch - '0') * 10 + (*(pch + 2) - '0');

			// Get CMD from ini
			TCHAR bufCMD[CFG_VALUE_LEN];
			GetCFGItem(nSection, nItem, ITEM_SUFFIX_CMD, szINI, bufCMD);
			dbg(_T("CMD: <%s>"), bufCMD);
			_tsystem(bufCMD);
			_tsystem(_T("pause"));
		}

		// After calling jumplist, display name had restored to file name.
		ChangeDisplayNames(szINI);
	}
}



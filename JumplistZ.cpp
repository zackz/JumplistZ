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
#pragma comment(lib, "Gdi32.lib")

const TCHAR NAME[]                = _T("JumplistZ");
const TCHAR VERSION[]             = _T("0.5.2");
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

void AddBackslash(LPTSTR bufPath)
{
	if (bufPath && *bufPath && bufPath[_tcslen(bufPath) - 1] != _T('\\'))
		_tcscat(bufPath, _T("\\"));
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

// http://msdn.microsoft.com/en-us/library/ms997538.aspx
// http://en.wikipedia.org/wiki/ICO_%28file_format%29
#pragma pack(push)
#pragma pack(1)
typedef struct
{
	BYTE        bWidth;          // Width, in pixels, of the image
	BYTE        bHeight;         // Height, in pixels, of the image
	BYTE        bColorCount;     // Number of colors in image (0 if >=8bpp)
	BYTE        bReserved;       // Reserved ( must be 0)
	WORD        wPlanes;         // Color Planes
	WORD        wBitCount;       // Bits per pixel
	DWORD       dwBytesInRes;    // How many bytes in this resource?
	DWORD       dwImageOffset;   // Where in the file is this image?
} ICONDIRENTRY, *LPICONDIRENTRY;
typedef struct
{
	WORD           idReserved;   // Reserved (must be 0)
	WORD           idType;       // Resource Type (1 for icons)
	WORD           idCount;      // How many images?
	ICONDIRENTRY   idEntries[1]; // An entry for each image (idCount of 'em)
} ICONDIR, *LPICONDIR;
#pragma pack(pop)

BOOL IconGetData(HICON hIcon, ICONDIRENTRY * pEntry, LPVOID * pRet)
{
	ICONINFO ii;
	memset(&ii, 0, sizeof(ii));
	HDC  hdc  = NULL;
	BOOL bRet = FALSE;
	*pRet = NULL;
	while (TRUE)
	{
		if (!GetIconInfo(hIcon, &ii))
		{
			dbg(_T("Error: GetIconInfo, %d"), hIcon);
			break;
		}
		hdc = CreateCompatibleDC(NULL);
		if (hdc == NULL)
		{
			dbg(_T("Error: CreateCompatibleDC"));
			break;
		}

		// GetDIBits might write data over bi.bmiHeader.biSize.
		char bufBI[1024];
		memset(&bufBI, 0, sizeof(bufBI));
		BITMAPINFO &bi = (BITMAPINFO &)bufBI;
		bi.bmiHeader.biSize = sizeof(bi);

		// GetDIBits - 1
		int n = GetDIBits(hdc, ii.hbmColor, 0, 0, NULL, &bi, DIB_RGB_COLORS);
		if (n == 0)
		{
			dbg(_T("Error: GetDIBits - 1"));
			break;
		}
		int nLenHeader = bi.bmiHeader.biSize;
		int nLenColor  = bi.bmiHeader.biSizeImage;
		int nLenMask   = (bi.bmiHeader.biWidth + 31) / 32 * 4 * bi.bmiHeader.biHeight;
		int nLenData   = nLenHeader + nLenMask + nLenColor;

		// Result buffer
		*pRet = malloc(nLenData);
		// Bitmap header
		bi.bmiHeader.biCompression = BI_RGB;
		memcpy(*pRet, &bi, nLenHeader);
		((BITMAPINFO *)*pRet)->bmiHeader.biHeight *= 2;

		// GetDIBits - 2
		n = GetDIBits(hdc, ii.hbmColor, 0, bi.bmiHeader.biHeight,
			(BYTE *)*pRet + nLenHeader, &bi, DIB_RGB_COLORS);
		if (n != bi.bmiHeader.biHeight)
		{
			dbg(_T("Error: GetDIBits - 2, %d"), n);
			break;
		}

		// GetDIBits - 3
		bi.bmiHeader.biBitCount = 1;
		n = GetDIBits(hdc, ii.hbmMask, 0, bi.bmiHeader.biHeight,
			(BYTE *)*pRet + nLenHeader + nLenColor, &bi, DIB_RGB_COLORS);
		if (n != bi.bmiHeader.biHeight)
		{
			dbg(_T("Error: GetDIBits - 3, %d"), n);
			break;
		}

		// Icon entry
		pEntry->bWidth        = bi.bmiHeader.biWidth;
		pEntry->bHeight       = bi.bmiHeader.biHeight;
		pEntry->bColorCount   = 0;
		pEntry->bReserved     = 0;
		pEntry->wPlanes       = 0;
		pEntry->wBitCount     = 0;
		pEntry->dwBytesInRes  = nLenData;
		pEntry->dwImageOffset = 0; // <-- Caller sets it

		// Done
		bRet = TRUE;
		break;
	}
	if (!bRet)
	{
		free(*pRet);
		*pRet = NULL;
	}
	DeleteDC(hdc);
	DeleteObject(ii.hbmMask);
	DeleteObject(ii.hbmColor);
	return bRet;
}

BOOL IconWriteToFile(HICON hIcon, LPCTSTR szFile)
{
	ICONDIR id;
	LPVOID  pData;
	if (!IconGetData(hIcon, id.idEntries, &pData))
		return FALSE;
	FILE * f = _tfopen(szFile, _T("wb"));
	if (!f)
	{
		dbg(_T("Error: IconWriteToFile, _tfopen, %s"), szFile);
		return FALSE;
	}
	id.idReserved = 0;
	id.idType     = 1;
	id.idCount    = 1;
	id.idEntries[0].dwImageOffset = sizeof(id);
	fwrite(&id, 1, sizeof(id), f);
	fwrite(pData, 1, id.idEntries[0].dwBytesInRes, f);
	fclose(f);
	free(pData);
	return TRUE;
}

void IconTestOnefile(LPCTSTR szFile)
{
	static int nIndex = 0;
	dbg(_T("%d, %s"), ++nIndex, szFile);

	UINT flags[] = {
		SHGFI_ICON,
		SHGFI_LARGEICON   | SHGFI_ICON,
		SHGFI_LINKOVERLAY | SHGFI_ICON,
		SHGFI_OPENICON    | SHGFI_ICON,
		SHGFI_SELECTED    | SHGFI_ICON,
		SHGFI_SMALLICON   | SHGFI_ICON
	};
	for (int i = 0; i < ARRAYSIZE(flags); i++)
	{
		TCHAR bufDest[MAX_PATH];
		_stprintf(bufDest, _T("dumps\\Icon_%d_%d.ico"), nIndex, i);
		dbg(_T("%s"), bufDest);

		SHFILEINFO sfi;
		if (0 != SHGetFileInfo(szFile, 0, &sfi, sizeof(sfi), flags[i]))
		{
			if (!IconWriteToFile(sfi.hIcon, bufDest))
				dbg(_T("Error: IconWriteToFile"));
			DestroyIcon(sfi.hIcon);
		}
		else
			dbg(_T("Error: SHGetFileInfo"));
	}
}

void IconTest()
{
	TCHAR bufPath[MAX_PATH];
	SHGetSpecialFolderPath(0, bufPath, CSIDL_SYSTEM, FALSE);
	AddBackslash(bufPath);
	TCHAR bufFind[MAX_PATH];
	_tcscpy(bufFind, bufPath);
	_tcscat(bufFind, _T("*"));
	WIN32_FIND_DATA wfd;
	HANDLE h = FindFirstFile(bufFind, &wfd);
	while (h != INVALID_HANDLE_VALUE)
	{
		TCHAR bufFile[MAX_PATH];
		_tcscpy(bufFile, bufPath);
		_tcscat(bufFile, wfd.cFileName);
		IconTestOnefile(bufFile);
		if (!FindNextFile(h, &wfd))
		{
			FindClose(h);
			h = INVALID_HANDLE_VALUE;
		}
	}
}

HRESULT ShellLinkSetIcon(IShellLink * psl, LPCTSTR szFile)
{
	// No icon? Not a big deal.
	HRESULT hrRet = S_OK;

	// Temporary path
	TCHAR bufTemp[MAX_PATH];
	if (0 == GetTempPath(MAX_PATH, bufTemp))
	{
		dbg(_T("Error: GetTempPath"));
		return hrRet;
	}
	AddBackslash(bufTemp);

	// Hash
	int hash = 0;
	for (TCHAR *p = GetCommandLine(); *p; p++)
		hash = hash * 31 + *p;

	// Icon file path
	static int nIndex = 1;
	TCHAR bufIcoFile[MAX_PATH];
	_stprintf(bufIcoFile, _T("%sJumplistZ_%08X_%d.ico"), bufTemp, hash, nIndex++);
	dbg(bufIcoFile);

	// Get icon
	SHFILEINFO sfi;
	memset(&sfi, 0, sizeof(sfi));
	if (0 == SHGetFileInfo(szFile, 0, &sfi, sizeof(sfi),
		SHGFI_SMALLICON | SHGFI_ICON))
	{
		dbg(_T("Error: SetIcon, SHGetFileInfo"));
		return hrRet;
	}

	// Write icon
	if (IconWriteToFile(sfi.hIcon, bufIcoFile))
		psl->SetIconLocation(bufIcoFile, 0);
	DestroyIcon(sfi.hIcon);
	return hrRet;
}

HRESULT ShellLinkSetTitle(IShellLink * psl, LPCTSTR szTitle)
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

		hr = ShellLinkSetIcon(psl, bufFile);
		if (FAILED(hr)) break;

		hr = psl->SetArguments(bufParam);
		if (FAILED(hr)) break;

		hr = ShellLinkSetTitle(psl, szName);
		if (FAILED(hr)) break;

		hr = psl->SetDescription(szCMD);
		if (FAILED(hr)) break;

		TCHAR bufPath[MAX_PATH];
		_tcscpy(bufPath, bufFile);
		TCHAR * pch = _tcsrchr(bufPath, _T('\\'));
		if (pch)
			*pch = 0;
		else
			SHGetSpecialFolderPath(0, bufPath, CSIDL_PROFILE, FALSE);
		hr = psl->SetWorkingDirectory(bufPath);
		if (FAILED(hr)) break;
		return psl;
	}
	dbg(_T("Error: GetShellLink, hr: 0x%08x"), hr);
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
			dbg(_T("Error: AppendCategory, %s, hr: 0x%08x"), bufGroupName, hr);
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
		//~ psl = GetShellLink(
			//~ _T("About JumplistZ"), _T("https://github.com/zackz/JumplistZ"));
		//~ if (psl)
			//~ poc->AddObject(psl);
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
		err(_T("Error: CoCreateInstance, CLSID_DestinationList, hr = 0x%08x"), hr);
		return 0;
	}

	UINT uMaxSlots;
	IObjectArray * poaRemoved;  // Not care, should remove it in INI file.
	hr = pcdl->BeginList(&uMaxSlots, IID_PPV_ARGS(&poaRemoved));
	if (FAILED(hr))
	{
		err(_T("Error: BeginList, hr = 0x%08x"), hr);
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
		AddBackslash(szINI);
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



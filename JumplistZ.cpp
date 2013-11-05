/*
JumplistZ
https://github.com/zackz/JumplistZ
*/

#define UNICODE
#define _UNICODE

#include <stdio.h>
#include <tchar.h>
#include <windows.h>
#include <Shlobj.h>
#include <propkey.h>
#include <Propvarutil.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")

const TCHAR NAME[]                = _T("JumplistZ");
const TCHAR VERSION[]             = _T("0.6.2");
const TCHAR CFGKEY_DEBUG_BITS[]   = _T("DEBUG_BITS");
const TCHAR CFGKEY_EDITOR[]       = _T("EDITOR");
const TCHAR CFGKEY_GROUP_NAME[]   = _T("GROUP_DISPLAY_NAME");
const TCHAR SECTION_PROPERTIES[]  = _T("PROPERTIES");
const TCHAR SECTION_PREFIX[]      = _T("GROUP");
const TCHAR ITEM_PREFIX[]         = _T("ITEM");
const TCHAR ITEM_SUFFIX_NAME[]    = _T("_NAME");
const TCHAR ITEM_SUFFIX_CMD[]     = _T("_CMD");
const int   CFG_MAX_COUNT         = 100;
const int   CFG_VALUE_LEN         = 1024;

TCHAR g_szAppName[MAX_PATH];
TCHAR g_szAppPath[MAX_PATH];
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
	/*
	IN:
		szCMD    = <   "file name" parameter1 parameter2 ...   >
		           <    file_name  parameter1 parameter2 ...   >
	OUT:
		bufFile  = <file name>
		bufParam = < parameter1 parameter2 ...   >
	*/
	const TCHAR * pFile = szCMD;
	while (*pFile && _tcschr(_T(" \t"), *pFile))
		pFile++;
	const TCHAR * pParam = pFile;
	if (*pFile == _T('"'))
	{
		pFile++;
		pParam = _tcschr(pFile, _T('"'));
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

BOOL SilentCMD(LPCTSTR szCMD, LPBYTE bufOut=NULL, DWORD * pdwLen=NULL)
{
	/*
	Run shell commands without console window, and retrieve output (stdout &
	stderr). Not fully implemented yet, apparently missing stdin, and using
	same pipe in stdout and stderr.
	Return TRUE after successful calling CreateProcess. "bufOut" is straight
	output of commands which is ansi string in most case. And "bufOut" always
	inluded an additional null terminator which wasn't counted in "pwdLen".
	*/
	BOOL   bRet      = FALSE;
	HANDLE hOutRead  = 0;
	HANDLE hOutWrite = 0;
	do
	{
		SECURITY_ATTRIBUTES sa;
		sa.nLength              = sizeof(SECURITY_ATTRIBUTES);
		sa.lpSecurityDescriptor = NULL;
		sa.bInheritHandle       = TRUE;
		bRet = CreatePipe(&hOutRead, &hOutWrite, &sa, 0);
		if (!bRet)
		{
			dbg(_T("Error: CreatePipe"));
			break;
		}

		STARTUPINFO si;
		memset(&si, 0, sizeof(si));
		si.dwFlags     = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
		si.hStdOutput  = hOutWrite;
		si.hStdError   = hOutWrite;
		si.wShowWindow = SW_HIDE;

		LPCTSTR pComSpec = _tgetenv(_T("ComSpec"));
		if (!pComSpec)
		{
			dbg(_T("Error: Can't get path of cmd.exe"));
			break;
		}

		TCHAR bufCMD[MAX_PATH] = _T("/c ");
		_tcscat(bufCMD, szCMD);
		PROCESS_INFORMATION pi;
		bRet = CreateProcess(pComSpec, bufCMD,
			NULL, NULL, TRUE, 0, NULL, NULL, &si, &pi);
		if (!bRet)
		{
			dbg(_T("Error: CreateProcess, %d"), GetLastError());
			break;
		}
		CloseHandle(hOutWrite);
		hOutWrite = 0;

		LPBYTE pBuf = bufOut;
		DWORD dwBufLen = bufOut ? *pdwLen - 1: 0;
		BYTE bufDummy[1024 * 32];
		DWORD dwRead;
		DWORD dw;
		BOOL b;
		while (1)
		{
			if (!PeekNamedPipe(hOutRead, NULL, 0, NULL, &dw, NULL))
				break;
			if (dw == 0)
			{
				b = GetExitCodeProcess(pi.hProcess, &dw);
				if (b && dw != STILL_ACTIVE)
					break;
				Sleep(200);
				continue;
			}

			dbg(_T("PeekNamedPipe, %d"), dw);
			if (dwBufLen != 0)
			{
				if (!ReadFile(hOutRead, pBuf, dwBufLen, &dwRead, NULL))
					break;
				pBuf += dwRead;
				dwBufLen -= dwRead;
			}
			else
			{
				if (!ReadFile(hOutRead, bufDummy, ARRAYSIZE(bufDummy), &dwRead, NULL))
					break;
			}
		}
		if (bufOut)
		{
			*pBuf = 0;
			*pdwLen = pBuf - bufOut;
		}
		CloseHandle(hOutRead);
		hOutRead = 0;
		bRet = TRUE;
	}
	while (0);
	CloseHandle(hOutRead);
	CloseHandle(hOutWrite);
	return bRet;
}

BOOL IsFile(LPCTSTR szFile)
{
	WIN32_FIND_DATA wfd;
	if (INVALID_HANDLE_VALUE == FindFirstFile(szFile, &wfd))
		return FALSE;
	return !(FILE_ATTRIBUTE_DIRECTORY & wfd.dwFileAttributes);
}

HRESULT ShellLinkSetIcon(IShellLink * psl, LPCTSTR szFile);

HRESULT ShellLinkSetOpenerIcon(IShellLink * psl, LPCTSTR szFile)
{
	HRESULT hr = E_FAIL;
	TCHAR buf[MAX_PATH];
	DWORD dw = ARRAYSIZE(buf);
	hr = AssocQueryString(0, ASSOCSTR_EXECUTABLE, szFile, NULL, buf, &dw);
	if (SUCCEEDED(hr))
	{
		dbg(_T("Opener: %s"), buf);
		if (0 == _tcscmp(buf, _T("%1")))
			return E_FAIL;  // *.exe, *.bat, ...
		else
			return ShellLinkSetIcon(psl, buf);
	}
	return hr;
}

HRESULT ShellLinkSetLinksIcon(IShellLink * psl, LPCTSTR szFile)
{
	HRESULT hr = E_FAIL;
	LPCTSTR prefixes[] = {_T("http://"), _T("https://")};
	int i = 0;
	for (; i < ARRAYSIZE(prefixes); i++)
		if (0 == _tcsncicmp(szFile, prefixes[i], sizeof(prefixes[i])))
			break;
	if (i == ARRAYSIZE(prefixes))
		return hr;

	static TCHAR bufTempHTML[MAX_PATH] = {0};
	if (*bufTempHTML == 0)
	{
		if (0 == GetTempPath(MAX_PATH, bufTempHTML))
		{
			dbg(_T("Error: GetTempPath"));
		}
		else
		{
			PathAppend(bufTempHTML, _T("JumplistZs_dummy.html"));
			FILE * f = _tfopen(bufTempHTML, _T("w"));
			if (f)
			{
				LPCSTR txt = "JumplistZ's dummy HTML file.\n";
				fwrite(txt, 1, strlen(txt), f);
				fclose(f);
			}
			else
				*bufTempHTML = 0;
		}
	}
	if (*bufTempHTML != 0)
	{
		return ShellLinkSetOpenerIcon(psl, bufTempHTML);
	}
	return hr;
}

HRESULT ShellLinkSetCMDsIcon(IShellLink * psl, LPCTSTR szFile)
{
	TCHAR bufNew[MAX_PATH];
	_tcscpy(bufNew, szFile);
	if (!PathFindOnPath(bufNew, NULL))
	{
		_tcscat(bufNew, _T(".exe"));
		if (!PathFindOnPath(bufNew, NULL))
		{
			return E_FAIL;
		}
	}
	dbg(_T("ShellLinkSetCMDsIcon, found %s"), bufNew);
	return ShellLinkSetIcon(psl, bufNew);
}

HRESULT ShellLinkSetIcon(IShellLink * psl, LPCTSTR szFile)
{
	HRESULT hrRet = S_OK;  // No icon? Not a big deal.
	HRESULT hr;

	dbg(_T("ShellLinkSetIcon, <%s>"), szFile);

	// Try to find opener's icon, such as *.py *.html
	if (IsFile(szFile))
	{
		LPCTSTR use_original_icon[] = {
			//~ _T(".ini"),
			//~ _T(".log"),
			NULL
		};
		LPCTSTR * p = use_original_icon;
		for (; *p; p++)
		{
			if (0 == _tcsicmp(szFile + _tcslen(szFile) - _tcslen(*p), *p))
				break;
		}
		if (!*p)
		{
			hr = ShellLinkSetOpenerIcon(psl, szFile);
			if (SUCCEEDED(hr))
				return hrRet;
		}
	}

	// Got default icon?
	TCHAR buf[MAX_PATH];
	DWORD dw = ARRAYSIZE(buf);
	hr = AssocQueryString(0, ASSOCSTR_DEFAULTICON,
		szFile, NULL, buf, &dw);
	if (SUCCEEDED(hr))
	{
		dbg(_T("Default icon: %s"), buf);
		if (0 == _tcsicmp(buf, _T("%1")))
		{
			dbg(_T("Icon: %s"), szFile);
			psl->SetIconLocation(szFile, 0);
		}
		else
		{
			TCHAR * pch = _tcschr(buf, _T(','));
			if (pch)
			{
				dw = _tstoi(pch + 1);
				*pch = 0;
				dbg(_T("Icon: %s, %d"), buf, dw);
				psl->SetIconLocation(buf, dw);
			}
			else
			{
				dbg(_T("Icon: %s"), buf);
				psl->SetIconLocation(buf, 0);
			}
		}
		return hrRet;
	}

	// Is url?
	hr = ShellLinkSetLinksIcon(psl, szFile);
	if (SUCCEEDED(hr))
		return hrRet;

	// Others, try expand commandline and append ".exe"
	hr = ShellLinkSetCMDsIcon(psl, szFile);
	if (SUCCEEDED(hr))
		return hrRet;

	dbg(_T("Error: AssocQueryString, %s"), szFile);
	psl->SetIconLocation(szFile, 0);  // %ComSpec%
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

BOOL GetProgramPathFromStartParameters(LPCTSTR szParam, LPTSTR bufOut)
{
	// Split parameters
	BYTE  buf[CFG_VALUE_LEN * 2];
	DWORD dwLen = sizeof(buf);
	TCHAR bufCMD[CFG_VALUE_LEN * 2];
	LPCTSTR fmt = _T("@echo off & for %%i in (%s) do (echo %%i)");
	_stprintf(bufCMD, fmt, szParam);
	if (!SilentCMD(bufCMD, buf, &dwLen))
		return FALSE;

	/*
	START ["title"] [/D path] [/I] [/MIN] [/MAX] [/SEPARATE | /SHARED]
	      [/LOW | /NORMAL | /HIGH | /REALTIME | /ABOVENORMAL | /BELOWNORMAL]
	      [/NODE <NUMA node>] [/AFFINITY <hex affinity mask>] [/WAIT] [/B]
	      [command/program] [parameters]
	*/
	BOOL bNODE = FALSE;
	BOOL bAFFINITY = FALSE;
	BOOL bD = FALSE;
	char bufPath[MAX_PATH] = {0};
	char bufProgram[MAX_PATH] = {0};
	char * pStart = (char *)buf;
	char * pEnd;
	while (pEnd = strchr(pStart, '\n'))
	{
		*pEnd = 0;
		if (pEnd - 1 >= 0 && *(pEnd - 1) == '\r')
			*(pEnd - 1) = 0;
		dbg(_T("Parameter:  <%S>"), pStart);

		if (pStart == (char *)buf && *pStart == '"')
		{
			// "title"
		}
		else if (bD)
		{
			strcpy(bufPath, pStart);
			bD = FALSE;
		}
		else if (bNODE || bAFFINITY)
		{
			bNODE = bAFFINITY = FALSE;
		}
		else if (0 == stricmp(pStart, "/D"))
		{
			bD = TRUE;
		}
		else if (0 == stricmp(pStart, "/NODE"))
		{
			bNODE = TRUE;
		}
		else if (0 == stricmp(pStart, "/AFFINITY"))
		{
			bAFFINITY = TRUE;
		}
		else if (*pStart == '/')
		{
			// Options without parameter
		}
		else
		{
			// command/program
			strcpy(bufProgram, pStart);
			break;
		}
		pStart = pEnd + 1;
	}
	PathUnquoteSpacesA(bufPath);
	PathUnquoteSpacesA(bufProgram);
	const char * bufOtherDirs[] = {bufPath, 0};
	PathFindOnPathA(bufProgram, bufOtherDirs);
	MultiByteToWideChar(CP_ACP, 0, bufProgram, strlen(bufProgram) + 1, bufOut, MAX_PATH);
	dbg(_T("bufPath,    <%S>"), bufPath);
	dbg(_T("bufProgram, <%S>"), bufProgram);
	dbg(_T("Result,     <%s>"), bufOut);
	return TRUE;
}

IShellLink * GetShellLink(LPCTSTR szName, LPCTSTR szCMD)
{
	TCHAR bufFile[CFG_VALUE_LEN];
	TCHAR bufParam[CFG_VALUE_LEN];
	TCHAR bufActualPath[MAX_PATH] = {0};
	SplitFileAndParameters(szCMD, bufFile, bufParam);
	dbg(_T("-Name: <%s>"), szName);
	dbg(_T("-CMD:  <%s>, <%s>"), bufFile, bufParam);
	if (0 == _tcsicmp(_T("start"), bufFile))
	{
		_tcscpy(bufFile, g_szAppPath);
		GetProgramPathFromStartParameters(bufParam, bufActualPath);
		dbg(_T("-CMD:  <%s>, <%s>"), bufFile, bufActualPath);
	}

	HRESULT hr = 0;
	do
	{
		IShellLink * psl;
		hr = CoCreateInstance(
			CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psl));
		if (FAILED(hr)) break;

		hr = psl->SetPath(bufFile);
		if (FAILED(hr)) break;

		hr = ShellLinkSetIcon(psl, *bufActualPath ? bufActualPath: bufFile);
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
	while (0);

	dbg(_T("Error: GetShellLink, hr: 0x%08x"), hr);
	return NULL;
}

IShellLink * GetShellLinkFromINI(DWORD nSection, DWORD nItem, TCHAR * szINI)
{
	TCHAR bufCMD[CFG_VALUE_LEN];
	GetCFGItem(nSection, nItem, ITEM_SUFFIX_CMD, szINI, bufCMD);
	if (_tcslen(bufCMD) == 0)
		return NULL;
	TCHAR bufName[CFG_VALUE_LEN];
	GetCFGItem(nSection, nItem, ITEM_SUFFIX_NAME, szINI, bufName);
	dbg(_T("> GetShellLink, section: %d, item: %d"), nSection, nItem);
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
		IShellLink * psi = GetShellLinkFromINI(nSection, k, szINI);
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
		IShellLink * psl;

		TCHAR bufValue[MAX_PATH * 3] = _T("\"");
		GetPrivateProfileString(SECTION_PROPERTIES, CFGKEY_EDITOR,
			NULL, bufValue + 1, MAX_PATH, szINI);
		if (_tcslen(bufValue) != 1)
		{
			_tcscat(bufValue, _T("\" \""));
		}
		_tcscat(bufValue, szINI);
		_tcscat(bufValue, _T("\""));
		psl = GetShellLink(_T("Edit configuration"), bufValue);
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
	// Title of MessageBox, "JumplistZ X.X.X"
	_stprintf(g_szAppName, _T("%s %s"), NAME, VERSION);

	// Check os version
	OSVERSIONINFOEX osvi;
	ZeroMemory(&osvi, sizeof(osvi));
	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
	if (GetVersionEx((OSVERSIONINFO *) &osvi))
	{
		if (osvi.dwMajorVersion * 10 + osvi.dwMinorVersion < 61)
		{
			err(_T("Not supported. Jumplist is a new feature in Windows 7\n\
https://github.com/zackz/JumplistZ"));
			return 0;
		}
	}

	// Get ini full path
	TCHAR bufFile[CFG_VALUE_LEN];
	TCHAR bufParam[CFG_VALUE_LEN];
	SplitFileAndParameters(GetCommandLine(), bufFile, bufParam);
	_tfullpath(g_szAppPath, bufFile, MAX_PATH);
	TCHAR szINI[MAX_PATH];
	_tcscpy(szINI, g_szAppPath);
	*_tcsrchr(szINI, '.') = 0;  // Remove ".exe"
	_tcscat(szINI, _T(".ini"));

	// Has parameters?
	const TCHAR * pParam = bufParam;
	while (*pParam && _tcschr(_T(" \t"), *pParam))
		pParam++;
	if (*pParam)
	{
		TCHAR cmd[CFG_VALUE_LEN] = _T("start ");
		_tcsncat(cmd, pParam, CFG_VALUE_LEN);
		return SilentCMD(cmd) ? 0: -1;
	}

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

		// Result message. App path is included in message because different path
		// leads to different jumplist even though the executable file is same.
		TCHAR buf[100 + MAX_PATH];
		_stprintf(buf, _T("Updated jumplist. Total %d items.\nPath: \"%s\""),
			nObjects, g_szAppPath);
		MessageBox(NULL, buf, g_szAppName, MB_OK | MB_ICONINFORMATION);
	}
	return 0;
}

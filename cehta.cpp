// HTML Application host for Windows CE
// Copyright (c) datadiode
// SPDX-License-Identifier: GPL-2.0-or-later
#include <windows.h>
#include <wininet.h>
#include <oleauto.h>
#include <tchar.h>

#ifndef _WIN32_WCE
#define GetProcAddressA GetProcAddress
#endif

#define SafeInvoke(p) (p) == NULL ? E_POINTER : (p)

// Some flags copied from mingw-w64's mshtmhst.h:
#define HTMLDLG_MODAL 0x20
#define HTMLDLG_VERIFY 0x100

template<typename f>
struct DllImport {
	FARPROC p;
	f operator*() const { return reinterpret_cast<f>(p); }
};

static struct urlmon {
	HMODULE h;
	DllImport<HRESULT(WINAPI*)(IMoniker *, LPCWSTR, IMoniker **)> CreateURLMoniker;
} const urlmon = {
	LoadLibraryW(L"URLMON.DLL"),
	GetProcAddressA(urlmon.h, "CreateURLMoniker"),
};

static struct mshtml {
	HMODULE h;
	DllImport<HRESULT(WINAPI*)(HWND, IMoniker *, DWORD, VARIANT *, LPWSTR, VARIANT *)> ShowHTMLDialogEx;
} const mshtml = {
	LoadLibraryW(L"MSHTML.DLL"),
	GetProcAddressA(mshtml.h, "ShowHTMLDialogEx"),
};

static HRESULT CreateLoaderScript(LPWSTR path, LPWSTR script, DWORD limit)
{
	LPCWSTR f = L"javascript:"
				L"xml=new ActiveXObject('MSXML.DOMDocument');"
				L"xml.async='false';xml.load(\"file://?\");"
				L"document.writeln(\"<title>?</title>\");"
				L"document.writeln(xml.text);"
				L"document.close()";
	LPWSTR d = script;
	LPWSTR const e = script + limit;
	LPWSTR s = NULL;
	for (WCHAR c; (d < e) && ((c = *d = *f++) != '\0'); )
	{
		if (c != '?')
		{
			++d;
			continue;
		}
		s = path;
		for (WCHAR c; (d < e) && (c = *s) != '\0' && (c != '?'); *d++ = c)
		{
			++s;
			switch (c)
			{
			case '\b': c = 'b'; break;
			case '\f': c = 'f'; break;
			case '\n': c = 'n'; break;
			case '\r': c = 'r'; break;
			case '\t': c = 't'; break;
			case '\v': c = 'v'; break;
			case '\"': c = '"'; break;
			case '\\': break;
			default: continue;
			}
			*d++ = '\\';
			if (d >= e)
				break;
		}
	}
	if (d >= e)
		return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
	if (s != NULL)
	{
		WCHAR c = *s;
		*s = '\0';
		DWORD attr = GetFileAttributesW(path);
		*s = c;
		if (attr & FILE_ATTRIBUTE_DIRECTORY) // path is invalid or points to a folder
			return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
	}
	return S_OK;
}

int WINAPI _tWinMain(HINSTANCE, HINSTANCE, LPWSTR url, int)
{
	HRESULT hr;
	if (SUCCEEDED(hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)))
	{
		WCHAR script[INTERNET_MAX_URL_LENGTH];
		if (SUCCEEDED(hr = CreateLoaderScript(url, script, _countof(script))))
		{
			IMoniker *pMoniker = NULL;
			if (SUCCEEDED(hr = SafeInvoke(*urlmon.CreateURLMoniker)(NULL, script, &pMoniker)))
			{
				VARIANT in, out;
				VariantInit(&in);
				VariantInit(&out);
				V_VT(&in) = VT_BSTR;
				V_BSTR(&in) = SysAllocString(url);
				WCHAR options[] = L"dialogWidth=80;dialogHeight=50;resizable:yes;"; // intentionally non-const
				hr = SafeInvoke(*mshtml.ShowHTMLDialogEx)(NULL, pMoniker, HTMLDLG_MODAL | HTMLDLG_VERIFY, &in, options, &out);
				if (SUCCEEDED(hr = VariantChangeType(&out, &out, 0, VT_I4)))
					hr = V_I4(&out);
				VariantClear(&in);
				VariantClear(&out);
				pMoniker->Release();
			}
		}
		CoUninitialize();
	}
	if (FAILED(hr))
	{
		WCHAR msg[1024];
		int len = wsprintfW(msg, L"CEhta failed with 0x%08lX\n", hr);
		FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, msg + len, _countof(msg) - len, NULL);
		MessageBoxW(NULL, msg, NULL, MB_ICONSTOP);
	}
	return hr;
}

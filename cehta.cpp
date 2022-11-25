// HTML Application host for Windows CE
// Copyright (c) datadiode
// SPDX-License-Identifier: GPL-2.0-or-later
#include <windows.h>
#include <wininet.h>
#include <oleauto.h>
#include <wchar.h>
#include <tchar.h>
#include <mshtml.h>
#include <mshtmhst.h>

#ifndef _WIN32_WCE
#define GetProcAddressA GetProcAddress
#endif

#define SafeInvoke(p) (p) == NULL ? E_POINTER : (p)

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

static HRESULT CoGetLastError()
{
	DWORD error = GetLastError();
	return HRESULT_FROM_WIN32(error);
}

class Loader
	: public IElementBehaviorFactory
	, public IDispatch
{
private:
	LPCWSTR const cmdline;
	BSTR content;	// pointer to beginning of file content
	LPWSTR payload;	// pointer to apparent beginning of html payload
	BSTR path;
public:
	Loader(LPCWSTR cl) : cmdline(cl), content(NULL), payload(NULL)
	{
		UINT len = static_cast<UINT>(wcslen(cl));
		if (LPCWSTR end = *cl == L'"' ? wmemchr(++cl, L'"', len) : wcspbrk(cl, L" \t\r\n"))
			len = static_cast<UINT>(end - cl);
		path = SysAllocStringLen(cl, len);
	}
	~Loader()
	{
		SysFreeString(path);
	}
	BSTR QueryProcessingInstruction(LPCWSTR name)
	{
		// Look for processing instructions only up to where the html payload
		// begins, to keep the effort low and stay safe from bogus matches
		LPWSTR p = content;
		while (p < payload)
		{
			*payload = L'\0';
			p = wcsstr(p, name);
			p = p ? p + wcslen(name) : payload;
			*payload = L'<';
			if (size_t n = wcsspn(p, L" \t\r\n"))
				if (LPWSTR q = wcsstr(p += n, L"?>"))
					return SysAllocStringLen(p, static_cast<UINT>(q - p));
		}
		return NULL;
	}
	HRESULT Open()
	{
		HRESULT hr = S_OK;
		HANDLE const file = CreateFile(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
		if (file != INVALID_HANDLE_VALUE)
		{
			DWORD cb = GetFileSize(file, NULL);
			if (cb != INVALID_FILE_SIZE)
			{
				if (PSTR pstrUtf8 = reinterpret_cast<PSTR>(SysAllocStringByteLen(NULL, cb)))
				{
					if (ReadFile(file, pstrUtf8, cb, &cb, NULL))
					{
						if (*reinterpret_cast<BSTR>(pstrUtf8) == 0xFEFF)
						{
							// file is in UCS-2 already, so no need to convert
							content = reinterpret_cast<BSTR>(pstrUtf8);
							*content = L'\n'; // hides BOM from MSHTML
							pstrUtf8 = NULL;
						}
						else if (DWORD const cch = MultiByteToWideChar(CP_UTF8, 0, pstrUtf8, cb, NULL, 0))
						{
							if (BSTR const bstrWide = SysAllocStringLen(NULL, cch))
							{
								MultiByteToWideChar(CP_UTF8, 0, pstrUtf8, cb, bstrWide, cch);
								content = bstrWide;
							}
						}
					}
					else
					{
						hr = CoGetLastError(); // from ReadFile()
					}
					SysFreeString(reinterpret_cast<BSTR>(pstrUtf8));
				}
			}
			else
			{
				hr = CoGetLastError(); // from GetFileSize()
			}
			CloseHandle(file);
			if (content)
			{
				LPWSTR p = content;
				do
				{
					payload = p = wcschr(p, '<');
				} while (p && wcschr(L"?!", *++p));
			}
			else
			{
				hr = E_OUTOFMEMORY;
			}
		}
		else
		{
			hr = CoGetLastError(); // from CreateFile()
		}
		return hr;
	}
	// IUnknown
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject)
	{
		*ppvObject = NULL;
		if (riid == IID_IUnknown || riid == IID_IElementBehaviorFactory)
			*ppvObject = static_cast<IElementBehaviorFactory *>(this);
		else if (riid == IID_IDispatch)
			*ppvObject = static_cast<IDispatch *>(this);
		else return E_NOINTERFACE;
		return S_OK;
	}
	STDMETHOD_(ULONG, AddRef)()
	{
		return 1;
	}
	STDMETHOD_(ULONG, Release)()
	{
		return 1;
	}
	// IElementBehaviorFactory
	STDMETHOD(FindBehavior)(BSTR, BSTR, IElementBehaviorSite *pSite, IElementBehavior **ppBehavior)
	{
		// Instead of doing what its name seems to imply, this method rather rewrites the document

		IHTMLElement *pElement = NULL;
		SafeInvoke(pSite)->GetElement(&pElement);
		IDispatch *pDispatch = NULL;
		SafeInvoke(pElement)->get_document(&pDispatch);
		IHTMLDocument2 *pDocument;
		SafeInvoke(pDispatch)->QueryInterface(&pDocument);
		IOleWindow *pOleWindow = NULL;
		SafeInvoke(pDispatch)->QueryInterface(&pOleWindow);

		HWND hwnd = NULL;
		if (SUCCEEDED(SafeInvoke(pOleWindow)->GetWindow(&hwnd)))
		{
			while (HWND hwndParent = GetParent(hwnd))
				hwnd = hwndParent;
			SetWindowText(hwnd, path);
			SetForegroundWindow(hwnd);
		}

		if (SAFEARRAY *const psa = SafeArrayCreateVector(VT_VARIANT, 0, 1))
		{
			VARIANT *const pvar = static_cast<VARIANT *>(psa->pvData);
			V_VT(pvar) = VT_BSTR;
			V_BSTR(pvar) = content;
			SafeInvoke(pDocument)->write(psa);
			SafeArrayDestroy(psa);
		}

		SafeInvoke(pDocument)->close();

		SafeInvoke(pOleWindow)->Release();
		SafeInvoke(pDocument)->Release();
		SafeInvoke(pDispatch)->Release();
		SafeInvoke(pElement)->Release();

		SysFreeString(content);
		payload = content = NULL;

		*ppBehavior = NULL;
		return S_OK;
	}
	// IDispatch
	STDMETHOD(GetTypeInfoCount)(UINT *) { return E_NOTIMPL; }
	STDMETHOD(GetTypeInfo)(UINT, LCID, ITypeInfo **) { return E_NOTIMPL; }
	STDMETHOD(GetIDsOfNames)(REFIID, LPOLESTR *rgNames, UINT cNames, LCID, DISPID *rgDispId)
	{
		// Offer hta:application like access to the command line by exposing
		// "commandLine" as an alias name for this object's default property
		if ((rgNames == NULL) || (rgDispId == NULL) || (cNames != 1))
			return E_INVALIDARG;
		if (rgNames[0] == NULL || wcsicmp(rgNames[0], L"commandLine") != 0)
			return DISP_E_UNKNOWNNAME;
		rgDispId[0] = 0;
		return S_OK;
	}
	STDMETHOD(Invoke)(DISPID dispid, REFIID, LCID, WORD wFlags, DISPPARAMS *, VARIANT *pVarResult, EXCEPINFO *, UINT *)
	{
		// Expose the URL as default property to allow scripts to extract parameters
		if ((dispid != 0) || ((wFlags & DISPATCH_PROPERTYGET) == 0))
			return DISP_E_MEMBERNOTFOUND;
		V_VT(pVarResult) = VT_BSTR;
		V_BSTR(pVarResult) = SysAllocString(cmdline);
		return S_OK;
	}
};

int WINAPI _tWinMain(HINSTANCE, HINSTANCE, LPWSTR cmdline, int)
{
	HRESULT hr;
	if (SUCCEEDED(hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)))
	{
		Loader loader(cmdline);
		if (SUCCEEDED(hr = loader.Open()))
		{
			WCHAR const script[] = L"javascript:document.writeln();document.close();document.body.addBehavior(null,dialogArguments)";
			IMoniker *pMoniker = NULL;
			if (SUCCEEDED(hr = SafeInvoke(*urlmon.CreateURLMoniker)(NULL, script, &pMoniker)))
			{
				VARIANT in, out;
				VariantInit(&in);
				VariantInit(&out);
				V_VT(&in) = VT_DISPATCH;
				V_DISPATCH(&in) = &loader;
				BSTR options = loader.QueryProcessingInstruction(L"<?cehta-options");
				if (options == NULL)
					options = SysAllocString(L"dialogWidth=80;dialogHeight=50;resizable=yes");
				hr = SafeInvoke(*mshtml.ShowHTMLDialogEx)(NULL, pMoniker, HTMLDLG_MODAL | HTMLDLG_VERIFY, &in, options, &out);
				if (SUCCEEDED(hr = VariantChangeType(&out, &out, 0, VT_I4)))
					hr = V_I4(&out);
				VariantClear(&in);
				VariantClear(&out);
				SysFreeString(options);
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

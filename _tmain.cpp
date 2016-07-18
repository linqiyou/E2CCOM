// _tmain.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"

typedef HRESULT (__stdcall* PROC_DllGetClassObject)(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv);

static IClassFactory* __stdcall CoGetClassObject(HMODULE hModule, CLSID clsid) {
	PROC_DllGetClassObject pfnGCO = (PROC_DllGetClassObject)GetProcAddress(hModule, "DllGetClassObject");
	if(pfnGCO == NULL) {
		return NULL;
	}
	LPVOID pICF(NULL);
	if(S_OK != pfnGCO(clsid, __uuidof(IClassFactory), &pICF)){
		return NULL;
	}
	return (IClassFactory*)pICF;
}

// static crate com
static LPVOID __stdcall CoCreateInstance(LPCSTR pszLdrLibName, CLSID clsid, IID iid) {
	HMODULE hModule = GetModuleHandleA(pszLdrLibName);
	if(hModule == NULL) {
		hModule = LoadLibraryA(pszLdrLibName);
	}
	if(hModule == NULL) {
		return NULL;
	}
	IClassFactory* pICF(CoGetClassObject(hModule, clsid));
	if(pICF == NULL) {
		return NULL;
	}
	LPVOID pIDispatch(NULL);
	if(S_OK != pICF->CreateInstance(NULL, iid, &pIDispatch)) {
		return NULL;
	}
	return pIDispatch;
}

static LPVOID __stdcall CoCreateInstance(LPCSTR pszLdrLibName, LPCTSTR clsid, LPCTSTR iid) {
	CLSID CLSID_UNKNOWN;
	if(S_OK != CLSIDFromString(clsid, &CLSID_UNKNOWN)) {
		return NULL;
	}
	IID IID_UNKNOWN;
	if(S_OK != IIDFromString(iid, &IID_UNKNOWN)) {
		return NULL;
	}
	return CoCreateInstance(pszLdrLibName, CLSID_UNKNOWN, IID_UNKNOWN);
}

int _tmain(int argc, _TCHAR* argv[])
{
	if(::CoInitialize(NULL) != S_OK) {
		throw;
	}
	CComPtr<IDispatch> pdisp = (IDispatch*)CoCreateInstance("G:\\dm.dll", L"{26037A0E-7CBD-4FFF-9C63-56F2D0770214}",
		L"{00020400-0000-0000-C000-000000000046}");
	if(pdisp == NULL) {
		throw;
	}
	VARIANT params[3]; // 
	params[1].vt = VT_BSTR; // class 
	params[1].bstrVal = L"Notepad";
	params[0].vt = VT_BSTR; // title
	params[0].bstrVal = NULL;
	if(S_OK != pdisp.InvokeN(L"FindWindow", params, 2, &params[2])) {
		throw;
	}
	printf_s("hWnd\t%d\r\n", (params[2]).intVal);
	return getchar();
}


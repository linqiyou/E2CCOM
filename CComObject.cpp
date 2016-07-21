#include "stdafx.h"
#include "CComObject.h"

using namespace System;

bool CComObject::CreateObject(LPCTSTR lpszTypeName)
{
	CLSID clsid = { 0x00000000, 0x0000, 0x0000, { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 } };
	if(FAILED(CLSIDFromString(lpszTypeName, &clsid))) {
		return false;
	}
	LPVOID ppv = NULL;
	if(FAILED(CoCreateInstance(clsid, NULL, CLSCTX_ALL,  __uuidof(IDispatch), &ppv))) {
		return false;
	}
	if(ppv == NULL) {
		return false;
	}
	this->m_ptr = (IDispatch*)ppv; return true;
}

VARIANT CComObject::Invoke(LPCTSTR lpszName, VARIANT* pvarParams)
{
	VARIANT value = { VT_NULL };
	if(lpszName == NULL || pvarParams == NULL) {
		return value;	
	}
	VARIANT* pVarBuffer = pvarParams;
	int nVarSize = 0;
	do {
		VARIANT value = *++pVarBuffer;
		if(value.wReserved1 != value.wReserved2) { 
			break;
		}
		nVarSize++;
	} while(true);
	pVarBuffer = new VARIANT[nVarSize];
	for(int j = 0, i = nVarSize - 1; i >= 0; i--, j++) {
		pVarBuffer[j] = pvarParams[i];
	}
	if(FAILED(this->m_ptr.InvokeN(lpszName, pVarBuffer, nVarSize, &value))) {
		value.vt = VT_NULL;
	}
	delete[] pVarBuffer; return value;
}

// 方法
_variant_t CComObject::Invoke(LPCTSTR lpszName, _variant_t* pvarParams)
{
	VARIANT* pargs = static_cast<VARIANT*>(pvarParams);
	return this->Invoke(lpszName, pargs);
}

 // 附加对象
bool CComObject::AttachObject(IDispatch* pdisp)
{
	if(pdisp == NULL) {
		return false;
	}
	this->m_ptr = pdisp; return true;
}

// 取属性
VARIANT CComObject::GetProperty(LPCTSTR lpszName)
{
	VARIANT value = { VT_NULL };
	if(lpszName == NULL) {
		return value;
	}
	if(FAILED(this->m_ptr.GetPropertyByName(lpszName, &value))) {
		value.vt = NULL;
	}
	return value;
}

// 设置属性
bool CComObject::SetProperty(LPCTSTR lpszName, VARIANT* args)
{
	if(lpszName == NULL) {
		return false;
	}
	return SUCCEEDED(this->m_ptr.PutPropertyByName(lpszName, args));
}

IDispatch* CComObject::operator*()
{
	return this->m_ptr;
}
#pragma once

namespace System
{
	class CComObject
	{
	private:
		CComPtr<IDispatch> m_ptr;

	public:
		bool CreateObject(LPCTSTR lpszTypeName);

	public:
		bool AttachObject(IDispatch* pdisp);

	public:
		VARIANT Invoke(LPCTSTR lpszName, VARIANT* pvarParams);
		_variant_t Invoke(LPCTSTR lpszName, _variant_t* pvarParams);
		VARIANT GetProperty(LPCTSTR lpszName);
		bool SetProperty(LPCTSTR lpszName, VARIANT* args);
	
	public:
		IDispatch* operator*();
	};
}


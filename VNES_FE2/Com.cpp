//
// COM利用サポートクラス
//
#include <objbase.h>
#include "Com.h"

INT	COM::m_nRefCount = 0;

LRESULT	COM::AddRef()
{
	if( !m_nRefCount ) {
		HRESULT	hr;
		if( (hr = ::CoInitialize( NULL )) != S_OK ) {
			return	hr;
		}
	}
	m_nRefCount++;
	return	0L;
}

void	COM::Release()
{
	if( !m_nRefCount ) {
		return;
	}
	if( !(--m_nRefCount) ) {
		::CoUninitialize();
	}
}


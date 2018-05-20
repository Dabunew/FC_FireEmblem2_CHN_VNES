//
// Debug output
//
#ifndef	__DEBUGOUT_INCLUDED__
#define	__DEBUGOUT_INCLUDED__

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdio.h>
#include <stdarg.h>

#include <string>
using namespace std;

typedef	void (__cdecl *DEBUGWINDOWOUTPUT)(LPCTSTR);

class CDebugOut {
public:
	CDebugOut();

	void __cdecl Out( LPSTR fmt, ... );
	void Out( const string& str );

protected:
	HWND	hWndDebugOutput;

private:
};

extern	CDebugOut	Dbg;

#if	defined(_DEBUG) || defined(_DEBUGOUT)
#define	DEBUGOUT	Dbg.Out
#else
#define	DEBUGOUT
#endif	// !_DEBUG

#endif	// !__DEBUGOUT_INCLUDED__


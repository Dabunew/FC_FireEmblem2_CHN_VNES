//
// DirectInput class
//
#ifndef	__DIRECTINPUT_INCLUDED__
#define	__DIRECTINPUT_INCLUDED__

#define	WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <mmsystem.h>

#define DIRECTINPUT_VERSION	0x0700
#include <dinput.h>

#include "typedef.h"
#include "macro.h"

class	CDirectInput
{
public:
	// パブリックメンバ変数
	typedef	struct	tagDIKEYTBL {
		WORD	key;
		LPCSTR	name;
	} DIKEYTBL, *LPDIKEYTBL;

	INT	m_nJoystickNum;
	enum { DIJOYSTICK_MAX = 8 };
	enum { DI_UP = 0, DI_DOWN = 1, DI_LEFT = 2, DI_RIGHT = 3, DI_BUTTON = 16 };

	BYTE	m_Sw[256+64*DIJOYSTICK_MAX];

	static	DIKEYTBL	DIKeyTable[];
	static	LPSTR		DIKeyDirTable[];

	// パブリックメンバ関数
	CDirectInput();
	virtual ~CDirectInput();

	BOOL	InitialDInput( HWND hWnd, HINSTANCE hInst );
	void	ReleaseDInput( void );

	void	Acquire();
	void	Unacquire();
	void	Poll();

	BOOL	AddJoystickDevice( GUID deviceguid );

	LPCSTR	SearchKeyName( INT key );

protected:
	// プロテクトメンバ変数
	LPDIRECTINPUT7		m_lpDI;
	LPDIRECTINPUTDEVICE	m_lpKeyboard;
	LPDIRECTINPUTDEVICE7	m_lpJoystick[DIJOYSTICK_MAX];

	// プロテクトメンバ関数
	static	BOOL CALLBACK DIEnumDevicesCallback( LPDIDEVICEINSTANCE lpddi, LPVOID pvRef );

	BOOL	SetupJoystick( HWND hWnd, LPDIRECTINPUTDEVICE7 lpJoy );
};

extern	CDirectInput	DirectInput;

#endif // !__DIRECTINPUT_INCLUDED__

//
// DirectInput class
//
#include "DirectInput.h"
#include "COM.h"

CDirectInput	DirectInput;

CDirectInput::DIKEYTBL	CDirectInput::DIKeyTable[] = {
	DIK_ESCAPE,	"ESC",		DIK_1,		"1",
	DIK_2,		"2",		DIK_3,		"3",
	DIK_4,		"4",		DIK_5,		"5",
	DIK_6,		"6",		DIK_7,		"7",
	DIK_8,		"8",		DIK_9,		"9",
	DIK_0,		"0",		DIK_MINUS,	"-",
	DIK_EQUALS,	"=",		DIK_BACK,	"BackSpace",
	DIK_TAB,	"TAB",		DIK_Q,		"Q",
	DIK_W,		"W",		DIK_E,		"E",
	DIK_R,		"R",		DIK_T,		"T",
	DIK_Y,		"Y",		DIK_U,		"U",
	DIK_I,		"I",		DIK_O,		"O",
	DIK_P,		"P",		DIK_LBRACKET,	"[",
	DIK_RBRACKET,	"]",		DIK_RETURN,	"Enter",
	DIK_LCONTROL,	"L Ctrl",	DIK_A,		"A",
	DIK_S,		"S",		DIK_D,		"D",
	DIK_F,		"F",		DIK_G,		"G",
	DIK_H,		"H",		DIK_J,		"J",
	DIK_K,		"K",		DIK_L,		"L",
	DIK_SEMICOLON,	";",		DIK_APOSTROPHE,	"'",
	DIK_GRAVE,	"`",		DIK_LSHIFT,	"L Shift",
	DIK_BACKSLASH,	"\\",		DIK_Z,		"Z",
	DIK_X,		"X",		DIK_C,		"C",
	DIK_V,		"V",		DIK_B,		"B",
	DIK_N,		"N",		DIK_M,		"M",
	DIK_COMMA,	",",		DIK_PERIOD,	".",
	DIK_SLASH,	"/",		DIK_RSHIFT,	"R Shift",
	DIK_MULTIPLY,	"*",		DIK_LMENU,	"L Alt",
	DIK_SPACE,	"Space",
	DIK_F1,		"F1",		DIK_F2,		"F2",
	DIK_F3,		"F3",		DIK_F4,		"F4",
	DIK_F5,		"F5",		DIK_F6,		"F6",
	DIK_F7,		"F7",		DIK_F8,		"F8",
	DIK_F9,		"F9",		DIK_F10,	"F10",

	DIK_NUMPAD7,	"Num 7",	DIK_NUMPAD8,	"Num 8",
	DIK_NUMPAD9,	"Num 9",	DIK_SUBTRACT,	"Num -",
	DIK_NUMPAD4,	"Num 4",	DIK_NUMPAD5,	"Num 5",
	DIK_NUMPAD6,	"Num 6",	DIK_ADD,	"Num +",
	DIK_NUMPAD1,	"Num 1",	DIK_NUMPAD2,	"Num 2",
	DIK_NUMPAD3,	"Num 3",	DIK_NUMPAD0,	"Num 0",
	DIK_DECIMAL,	"Num .",	DIK_F11,	"F11",
	DIK_F12,	"F12",		DIK_F13,	"F13",
	DIK_F14,	"F14",		DIK_F15,	"F15",
	DIK_CONVERT,	"変換",
	DIK_NOCONVERT,	"無変換",	DIK_YEN,	"\\",
	DIK_NUMPADEQUALS,"Num =",	DIK_CIRCUMFLEX,	"^",
	DIK_AT,		"@",		DIK_COLON,	":",
	DIK_UNDERLINE,	"_",
	DIK_STOP,	"Stop",		DIK_NUMPADENTER,"Num Enter",
	DIK_RCONTROL,	"R Ctrl",	DIK_NUMPADCOMMA,"Num ,",
	DIK_DIVIDE,	"Num /",	DIK_SYSRQ,	"SysRq",
	DIK_RMENU,	"R Alt",	DIK_PAUSE,	"Pause",
	DIK_HOME,	"Home",		DIK_UP,		"Up",
	DIK_PRIOR,	"Page Up",	DIK_LEFT,	"Left",
	DIK_RIGHT,	"Right",	DIK_END,	"End",
	DIK_DOWN,	"Down",		DIK_NEXT,	"Page Down",
	DIK_INSERT,	"Insert",	DIK_DELETE,	"Delete",
	DIK_LWIN,	"L Windows",	DIK_LWIN,	"R Windows",
	DIK_APPS,	"AppMenu",

#if	0
// トグル系キーなので使えない
	DIK_CAPITAL,	"Caps Lock",
	DIK_NUMLOCK,	"NumLock",
	DIK_SCROLL,	"ScrollLock",
	DIK_KANA,	"カナ",	
	DIK_KANJI,	"漢字",
#endif
	0x00,		NULL
};

LPSTR	CDirectInput::DIKeyDirTable[] = {
	"Up", "Down", "Left", "Right"
};

//////////////////////////////////////////////////////////////////////
// 構築/消滅
//////////////////////////////////////////////////////////////////////

CDirectInput::CDirectInput()
{
	m_lpDI          = NULL;
	m_lpKeyboard    = NULL;

	m_nJoystickNum  = 0;

	ZeroMemory( m_lpJoystick, sizeof(m_lpJoystick) );
	ZeroMemory( m_Sw, sizeof(m_Sw) );
}

CDirectInput::~CDirectInput()
{
	ReleaseDInput();
}

//////////////////////////////////////////////////////////////////////
// メンバ関数
//////////////////////////////////////////////////////////////////////
// デバイスオブジェクト列挙コールバック
BOOL CALLBACK CDirectInput::DIEnumDevicesCallback( LPDIDEVICEINSTANCE lpddi, LPVOID pvRef )
{
	CDirectInput* pCDi = (CDirectInput*)pvRef;


	if( pCDi->AddJoystickDevice( lpddi->guidInstance ) )
		return	DIENUM_CONTINUE;

	return	DIENUM_STOP;
}

// ジョイスティックデバイスの設定
BOOL CDirectInput::SetupJoystick(HWND hWnd, LPDIRECTINPUTDEVICE7 lpJoy)
{
	if( lpJoy->SetDataFormat( &c_dfDIJoystick ) != DI_OK ) {
		return	FALSE;
	}

	if( lpJoy->SetCooperativeLevel( hWnd, DISCL_NONEXCLUSIVE|DISCL_BACKGROUND) != DI_OK ) {
		return	FALSE;
	}

	// レンジの設定
	DIPROPRANGE	diprg; 
	diprg.diph.dwSize       = sizeof(DIPROPRANGE);
	diprg.diph.dwHeaderSize = sizeof(DIPROPHEADER);
	diprg.diph.dwHow        = DIPH_BYOFFSET;
	diprg.diph.dwObj        = DIJOFS_X;
	diprg.lMin              = -10000;
	diprg.lMax              = +10000;
	lpJoy->SetProperty( DIPROP_RANGE, &diprg.diph );
	diprg.diph.dwObj        = DIJOFS_Y;
	lpJoy->SetProperty( DIPROP_RANGE, &diprg.diph );

	// デッドゾーンの設定
	DIPROPDWORD dipdw;
	dipdw.diph.dwSize       = sizeof(DIPROPDWORD);
	dipdw.diph.dwHeaderSize = sizeof(dipdw.diph);
	dipdw.diph.dwHow        = DIPH_BYOFFSET;
	dipdw.dwData            = 2000;
	dipdw.diph.dwObj         = DIJOFS_X;
	lpJoy->SetProperty(DIPROP_DEADZONE, &dipdw.diph);
	dipdw.diph.dwObj         = DIJOFS_Y;
	lpJoy->SetProperty(DIPROP_DEADZONE, &dipdw.diph);

	// 軸モードの設定
	dipdw.diph.dwSize       = sizeof(DIPROPDWORD);
	dipdw.diph.dwHeaderSize = sizeof(dipdw.diph);
	dipdw.diph.dwHow        = DIPH_DEVICE;
	dipdw.dwData            = DIPROPAXISMODE_ABS;
	dipdw.diph.dwObj        = 0;
	lpJoy->SetProperty( DIPROP_AXISMODE, &dipdw.diph );

	if( lpJoy->Acquire() != DI_OK ) {
	}

	return	TRUE;
}

// ジョイスティックデバイスオブジェクトの作成
BOOL CDirectInput::AddJoystickDevice(GUID deviceguid)
{
	if( m_nJoystickNum > DIJOYSTICK_MAX-1 )
		return	FALSE;

	if( m_lpDI->CreateDeviceEx( deviceguid, IID_IDirectInputDevice7,
		(LPVOID*)&m_lpJoystick[m_nJoystickNum], NULL ) != DI_OK ) {
		return	FALSE;
	}

	m_nJoystickNum++;

	return	TRUE;
}

#define	COMUSE	TRUE

// DirectInputオブジェクト／デバイスオブジェクトの構築
BOOL CDirectInput::InitialDInput(HWND hWnd, HINSTANCE hInst)
{
	try {
		// CDirectInputオブジェクトの作成
#if	!COMUSE
		if( DirectInputCreateEx( hInst, DIRECTINPUT_VERSION, IID_IDirectInput7, (LPVOID*)&m_lpDI, NULL ) != DI_OK ) {
			m_lpDI = NULL;
			throw "CDirectInput:DirectInputCreateEx failed.";
		}
#else
		// COM的利用
		COM::AddRef();
		if( ::CoCreateInstance( CLSID_DirectInput, NULL, CLSCTX_INPROC_SERVER, IID_IDirectInput7, (VOID**)&m_lpDI ) != S_OK ) {
			m_lpDI = NULL;
			throw	"CDirectInput:CoCreateInstance failed.";
		}
		if( m_lpDI->Initialize( hInst, DIRECTINPUT_VERSION ) != DI_OK )
			throw	"CDirectInput:IDirectInput7->Initialize failed.";
#endif

		if( m_lpDI->CreateDevice( GUID_SysKeyboard, &m_lpKeyboard, NULL ) != DI_OK )
			throw	"CDirectInput:CreateDevice failed.";

		if( m_lpKeyboard ) {
			if( m_lpKeyboard->SetDataFormat( &c_dfDIKeyboard ) != DI_OK )
				throw	"CDirectInput:SetDataFormat failed.";
			if( m_lpKeyboard->SetCooperativeLevel( hWnd, DISCL_NONEXCLUSIVE|DISCL_BACKGROUND) != DI_OK )
				throw	"CDirectInput:SetCooperativeLevel failed.";
			if( m_lpKeyboard->Acquire() != DI_OK ) {
			}
		}

		m_nJoystickNum = 0;
		if( m_lpDI->EnumDevices( DIDEVTYPE_JOYSTICK, (LPDIENUMDEVICESCALLBACK)DIEnumDevicesCallback,
					(LPVOID)this, DIEDFL_ATTACHEDONLY ) != DI_OK ) {
		}

		if( !m_nJoystickNum ) {
		} else {
			for( INT i = 0; i < m_nJoystickNum; i++ ) {
				if( !SetupJoystick( hWnd, m_lpJoystick[i] ) )
					throw 	"CDirectInput:SetProperty failed.";
			}
		}

	} catch( char *str ) {
		ReleaseDInput();

		MessageBox( hWnd, str, "ERROR", MB_ICONERROR|MB_OK );

		return	FALSE;
	}

	return	TRUE;
}

void CDirectInput::ReleaseDInput()
{
	for( INT i = 0; i < m_nJoystickNum; i++ ) {
		if( m_lpJoystick[i] ) {
//			m_lpJoystick[i]->Unacquire();
			RELEASE( m_lpJoystick[i] );
		}
	}

	if( m_lpKeyboard ) {
//		m_lpKeyboard->Unacquire();
		RELEASE( m_lpKeyboard );
	}

	if( m_lpDI ) {
		RELEASE( m_lpDI );
#if	COMUSE
		COM::Release();
#endif
	}
}

// 入力フォーカスを取得
void CDirectInput::Acquire()
{
	if( !m_lpDI )
		return;
	if( m_lpKeyboard )
		m_lpKeyboard->Acquire();
	for( INT i = 0; i < m_nJoystickNum; i++ )
		m_lpJoystick[i]->Acquire();
}

// 入力フォーカスを開放
void CDirectInput::Unacquire()
{
	if( !m_lpDI )
		return;
	if( m_lpKeyboard )
		m_lpKeyboard->Unacquire();
	for( INT i = 0; i < m_nJoystickNum; i++ )
		m_lpJoystick[i]->Unacquire();
}

// データポーリング
void CDirectInput::Poll()
{
DIJOYSTATE	js;

	ZeroMemory( m_Sw, sizeof(m_Sw) );

	if( !m_lpDI ) {
		return;
	}

	if( m_lpKeyboard ) {
		if( m_lpKeyboard->GetDeviceState( 256, &m_Sw ) == DIERR_INPUTLOST ) {
			m_lpKeyboard->Acquire();
			m_lpKeyboard->GetDeviceState( 256, &m_Sw );
		}
	}

	INT	idx;
	for( INT i = 0; i < m_nJoystickNum; i++ ) {
		idx = 256+i*64;

		if( m_lpJoystick[i]->Poll() == DIERR_INPUTLOST ) {
			m_lpJoystick[i]->Acquire();
			m_lpJoystick[i]->Poll();
		}
		if( m_lpJoystick[i]->GetDeviceState( sizeof(DIJOYSTATE), &js ) != DI_OK )
			ZeroMemory( &js, sizeof(DIJOYSTATE) );

		if( js.lX > 8000 )
			m_Sw[idx + DI_RIGHT] = 0x80;
		if( js.lX < -8000 )
			m_Sw[idx + DI_LEFT] = 0x80;
		if( js.lY > 8000 )
			m_Sw[idx + DI_DOWN] = 0x80;
		if( js.lY < -8000 )
			m_Sw[idx + DI_UP] = 0x80;

		for( INT j = 0; j < 32; j++ ) {
			m_Sw[idx + DI_BUTTON + j] = js.rgbButtons[j];
		}
	}
}

LPCSTR	CDirectInput::SearchKeyName( INT key )
{
LPDIKEYTBL kt = DIKeyTable;
static	CHAR	KeyStr[32];

	if( key == 0x00 )
		return	NULL;

	if( key < 0x100 ) {
		while( kt->name != NULL ) {
			if( kt->key == key )
				return	kt->name;
			kt++;
		}
	} else {
		INT	no  = (key-256)>>6;
		INT	idx = key & 0x3F;
		if( idx <= DI_RIGHT ) {
			::wsprintf( KeyStr, "J:%d %s", no, DIKeyDirTable[idx] );
			return	KeyStr;
		} else if( idx >= DI_BUTTON && idx < DI_BUTTON+32 ) {
			::wsprintf( KeyStr, "J:%d B:%02d", no, idx-DI_BUTTON );
			return	KeyStr;
		}
	}

	return	NULL;
}


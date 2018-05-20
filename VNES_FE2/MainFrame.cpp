//
// 儊僀儞僂僀儞僪僂僋儔僗
//
//
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <commdlg.h>
#include <imm.h>

#include <string>
using namespace std;

#include "VirtuaNESres.h"
#include "App.h"
#include "Pathlib.h"
#include "Plugin.h"

#include "Wnd.h"
#include "WndHook.h"
#include "MainFrame.h"
#include "Recent.h"
#include "Config.h"

#include "AboutDlg.h"
#include "EmulatorDlg.h"
#include "GraphicsDlg.h"
#include "SoundDlg.h"
#include "ControllerDlg.h"
#include "ShortcutDlg.h"
//#include "FolderDlg.h"
//#include "LanguageDlg.h"
#include "MovieDlg.h"
#include "GameOptionDlg.h"
#include "MovieInfoDlg.h"

#include "DirectDraw.h"
#include "DirectSound.h"
#include "DirectInput.h"

#include "nes.h"
#include "rom.h"
#include "pad.h"

// 儊僢僙乕僕儅僢僾
WND_MESSAGE_BEGIN(CMainFrame)
// 儊僢僙乕僕
WND_ON_MESSAGE( WM_CREATE,	OnCreate )
WND_ON_MESSAGE( WM_CLOSE,	OnClose )
WND_ON_MESSAGE( WM_DESTROY,	OnDestroy )
WND_ON_MESSAGE( WM_GETMINMAXINFO, OnGetMinMaxInfo )
WND_ON_MESSAGE( WM_ACTIVATE,	OnActivate )
WND_ON_MESSAGE( WM_ACTIVATEAPP,	OnActivateApp )
WND_ON_MESSAGE( WM_VNS_SHORTCUTENABLE, OnShortCutEnable )
WND_ON_MESSAGE( WM_ENABLE,	OnEnable )
WND_ON_MESSAGE( WM_ENTERMENULOOP, OnEnterMenuLoop )
WND_ON_MESSAGE( WM_EXITMENULOOP, OnExitMenuLoop )
WND_ON_MESSAGE( WM_SETFOCUS,	OnSetFocus )
WND_ON_MESSAGE( WM_KILLFOCUS,	OnKillFocus )
WND_ON_MESSAGE( WM_INITMENU,	OnInitMenu )
WND_ON_MESSAGE( WM_INITMENUPOPUP, OnInitMenuPopup )
WND_ON_MESSAGE( WM_PAINT,	OnPaint )
WND_ON_MESSAGE( WM_DISPLAYCHANGE, OnDisplayChange )
WND_ON_MESSAGE( WM_PALETTECHANGED, OnPaletteChanged )
WND_ON_MESSAGE( WM_QUERYNEWPALETTE, OnQueryNewPalette )
WND_ON_MESSAGE( WM_MENUCHAR,	OnMenuChar )
WND_ON_MESSAGE( WM_KEYDOWN,	OnKeyDown )
WND_ON_MESSAGE( WM_SIZE,	OnSize )
WND_ON_MESSAGE( WM_SYSCOMMAND,	OnSysCommand )
WND_ON_MESSAGE( WM_EXITSIZEMOVE, OnExitSizeMove )

WND_ON_MESSAGE( WM_SETCURSOR,	OnSetCursor )
WND_ON_MESSAGE( WM_TIMER,	OnTimer )

WND_ON_MESSAGE( WM_COPYDATA, OnCopyData )

WND_ON_MESSAGE( WM_VNS_ERRORMSG, OnErrorMessage )

//WND_ON_MESSAGE( WM_DROPFILES,	OnDropFiles )
WND_ON_MESSAGE( WM_VNS_COMMANDLINE, OnCommandLine )
//WND_ON_MESSAGE( WM_VNS_LAUNCHERCMD, OnLauncherCommand )

//WND_ON_MESSAGE( WM_NETPLAY, OnNetPlay )
//WND_ON_MESSAGE( WM_NETPLAY_CLOSE, OnNetPlayClose )
//WND_ON_MESSAGE( WM_NETPLAY_ERROR, OnNetPlayError )

// 僐儅儞僪
WND_COMMAND_BEGIN()

WND_ON_COMMAND( ID_SELFONT, OnSelFont )

//WND_ON_COMMAND( ID_EXIT, OnExit )
WND_ON_COMMAND( ID_HTMLHELP, OnHelp )
WND_ON_COMMAND( ID_ABOUT, OnAbout )
//WND_ON_COMMAND( ID_OPEN, OnFileOpen )
//WND_ON_COMMAND( ID_CLOSE, OnFileClose )
//WND_ON_COMMAND( ID_ROMINFO, OnRomInfo )
//WND_ON_COMMAND( ID_WAVERECORD, OnWaveRecord )
//WND_ON_COMMAND( ID_LAUNCHER, OnLauncher )
//WND_ON_COMMAND_RANGE( ID_MRU_PATH0, ID_MRU_PATH9, OnRecentOpenPath )
//WND_ON_COMMAND_RANGE( ID_MRU_FILE0, ID_MRU_FILE9, OnRecentOpen )

//WND_ON_COMMAND( ID_NETPLAY_CONNECT, OnNetPlayConnect )
//WND_ON_COMMAND( ID_NETPLAY_DISCONNECT, OnNetPlayDisconnect )

WND_ON_COMMAND( ID_CFG_EMULATOR, OnEmulatorCfg )
WND_ON_COMMAND( ID_CFG_GRAPHICS, OnGraphicsCfg )
WND_ON_COMMAND( ID_CFG_SOUND, OnSoundCfg )
WND_ON_COMMAND( ID_CFG_CONTROLLER, OnControllerCfg )
WND_ON_COMMAND( ID_CFG_SHORTCUT, OnShortcutCfg )
//WND_ON_COMMAND( ID_CFG_FOLDER, OnFolderCfg )
//WND_ON_COMMAND( ID_CFG_LANGUAGE, OnLanguageCfg )
WND_ON_COMMAND( ID_CFG_MOVIE, OnMovieCfg )
WND_ON_COMMAND( ID_CFG_GAMEOPTION, OnGameOptionCfg )

WND_ON_COMMAND( ID_SEARCH, OnSearch )
WND_ON_COMMAND( ID_CHEAT, OnCheat )
WND_ON_COMMAND_RANGE( ID_CHEAT_ENABLE, ID_CHEAT_DISABLE, OnCheatCommand )
WND_ON_COMMAND( ID_GENIE, OnGenie )

WND_ON_COMMAND( ID_FULLSCREEN, OnFullScreen )
WND_ON_COMMAND( ID_ONTOP, OnOnTop )
WND_ON_COMMAND_RANGE( ID_ZOOMx1, ID_ZOOMx4, OnZoom )
WND_ON_COMMAND( ID_HWRESET, OnEmuCommand )
WND_ON_COMMAND( ID_SWRESET, OnEmuCommand )
WND_ON_COMMAND( ID_PAUSE, OnEmuCommand )
WND_ON_COMMAND( ID_THROTTLE, OnEmuCommand )
WND_ON_COMMAND( ID_STATE_UP, OnEmuCommand )
WND_ON_COMMAND( ID_STATE_DOWN, OnEmuCommand )
WND_ON_COMMAND( ID_FPSDISP, OnEmuCommand )
WND_ON_COMMAND( ID_TVASPECT, OnEmuCommand )
WND_ON_COMMAND( ID_TVFRAME, OnEmuCommand )
WND_ON_COMMAND( ID_SCANLINE, OnEmuCommand )
WND_ON_COMMAND( ID_ALLLINE, OnEmuCommand )
WND_ON_COMMAND( ID_ALLSPRITE, OnEmuCommand )
WND_ON_COMMAND( ID_SYNCDRAW, OnEmuCommand )
WND_ON_COMMAND( ID_FITSCREEN, OnEmuCommand )
WND_ON_COMMAND( ID_SNAPSHOT, OnEmuCommand )

WND_ON_COMMAND( ID_LEFTCLIP, OnEmuCommand )
WND_ON_COMMAND( ID_ONEFRAME, OnEmuCommand )

WND_ON_COMMAND( ID_STATE_LOAD, OnStateCommand )
WND_ON_COMMAND( ID_STATE_SAVE, OnStateCommand )

WND_ON_COMMAND( ID_MOVIE_PLAY, OnMovieCommand )
WND_ON_COMMAND( ID_MOVIE_REC, OnMovieCommand )
WND_ON_COMMAND( ID_MOVIE_REC_APPEND, OnMovieCommand )
WND_ON_COMMAND( ID_MOVIE_STOP, OnMovieCommand )
WND_ON_COMMAND( ID_MOVIE_INFO, OnMovieInfo )

WND_ON_COMMAND( ID_FRAMESKIP_AUTO, OnEmuCommand )
WND_ON_COMMAND( ID_FRAMESKIP_UP, OnEmuCommand )
WND_ON_COMMAND( ID_FRAMESKIP_DOWN, OnEmuCommand )

// Viewers
//WND_ON_COMMAND( ID_VIEW_PATTERN, OnViewCommand )
//WND_ON_COMMAND( ID_VIEW_NAMETABLE, OnViewCommand )
//WND_ON_COMMAND( ID_VIEW_PALETTE, OnViewCommand )
//WND_ON_COMMAND( ID_VIEW_MEMORY, OnViewCommand )

//
WND_ON_COMMAND_RANGE( ID_STATE_SLOT0, ID_STATE_SLOT9, OnEmuCommand )
//WND_ON_COMMAND_RANGE( ID_DISK_EJECT, ID_DISK_1B, OnEmuCommand )
//WND_ON_COMMAND_RANGE( ID_EXCTR_NONE, ID_EXCTR_CRAZYCLIMBER, OnEmuCommand )
WND_ON_COMMAND_RANGE( ID_MUTE_0, ID_MUTE_F, OnEmuCommand )
WND_COMMAND_END()
WND_MESSAGE_END()

BOOL	CMainFrame::Create( HWND hWndParent )
{
	// NES僆僽僕僃僋僩
	Nes = NULL;

	// 儊儞僶曄悢
	m_hMenu = NULL;

	WNDCLASSEX	wcl;
	ZeroMemory( &wcl, sizeof(wcl) );
	wcl.cbSize		= sizeof(wcl);
	wcl.lpszClassName	= "VirtuaNESwndclass";
	wcl.lpfnWndProc		= g_WndProc;
//	wcl.style		= CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
	wcl.style		= CS_DBLCLKS;
	wcl.cbClsExtra		= wcl.cbWndExtra = 0;
	wcl.hInstance		= CApp::GetInstance();
	wcl.lpszMenuName	= NULL;
	wcl.hIcon		= 
	wcl.hIconSm		= ::LoadIcon( CApp::GetInstance(), MAKEINTRESOURCE(IDI_ICON) );
	wcl.hCursor		= ::LoadCursor( NULL, IDC_ARROW );
	wcl.hbrBackground	= 0;

	if( !RegisterClassEx( &wcl ) ) {
//		DEBUGOUT( "RegisterClassEx faild.\n" );
		return	FALSE;
	}

	HWND hWnd = CreateWindowEx(
			WS_EX_CLIENTEDGE,
			VIRTUANES_WNDCLASS,
			"圣火徽章外传-VirtuaNES",
			WS_OVERLAPPEDWINDOW,
			CW_USEDEFAULT,
			CW_USEDEFAULT,
			CW_USEDEFAULT,
			CW_USEDEFAULT,
			hWndParent,
			NULL,
			CApp::GetInstance(),
			(LPVOID)this		// This 傪杽傔崬傓堊
		);
	if( !hWnd ) {
//		DEBUGOUT( "CreateWindow faild.\n" );
		return	FALSE;
	}

	// 儊僀儞僂僀儞僪僂偲偟偰搊榐
	CApp::SetHWnd( hWnd );

	// 僗僥乕僩僗儘僢僩
	m_nStateSlot = 0;

	return	TRUE;
}

void	CMainFrame::Destroy()
{
	if( m_hWnd && IsWindow(m_hWnd) ) {
		::DestroyWindow( m_hWnd );
		m_hWnd = NULL;
	}
}

BOOL	CMainFrame::PreTranslateMessage( MSG* pMsg )
{
	if( pMsg->message >= WM_KEYFIRST && pMsg->message <= WM_KEYLAST ) {
//		return	TRUE;
	}
	return	FALSE;
}

WNDMSG	CMainFrame::OnCreate( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnCreate\n" );

	DirectDraw.SetSystemMemory( Config.graphics.bSystemMemory );
	DirectDraw.SetUseHEL      ( Config.graphics.bUseHEL );

	if( !DirectDraw.InitialDDraw( m_hWnd ) ) {
		hResult = -1L;
		return	TRUE;
	}
	if( !DirectSound.InitialDSound( m_hWnd ) ) {
		hResult = -1L;
		return	TRUE;
	}
	if( !DirectInput.InitialDInput( m_hWnd, CApp::GetInstance() ) ) {
		hResult = -1L;
		return	TRUE;
	}
	// DirectDraw堤呒偺愝掕
	DirectDraw.SetFlipMode     ( Config.graphics.bSyncDraw );
	DirectDraw.SetAspectMode   ( Config.graphics.bAspect );
	DirectDraw.SetAllLineMode  ( Config.graphics.bAllLine );
	DirectDraw.SetMaxZoom      ( Config.graphics.bFitZoom );
	DirectDraw.SetTVFrameMode  ( Config.graphics.bTVFrame );
	DirectDraw.SetScanlineMode ( Config.graphics.bScanline );
	DirectDraw.SetScanlineColor( Config.graphics.nScanlineColor );

	DirectDraw.SetWindowVSyncMode( Config.graphics.bWindowVSync );

	DirectDraw.SetDoubleSize  ( Config.graphics.bDoubleSize );

	if(strcmp(Config.general.stLogFont.lfFaceName, "没选") != 0)
		memcpy((void *)&DirectDraw.m_logFont, (void *)&Config.general.stLogFont, sizeof(LOGFONT));


	// DirectDraw僒乕僼僃僗偺峔抸
	if( !DirectDraw.InitialSurface( FALSE ) ) {
		hResult = -1L;
		return	TRUE;
	}

	// 僷儗僢僩僼傽僀儖偺儘乕僪
	if( Config.graphics.bPaletteFile ) {
		DirectDraw.SetPaletteFile( Config.graphics.szPaletteFile );
	}

	// DirectSound堤呒偺愝掕
	DirectSound.SetSamplingRate( Config.sound.nRate, Config.sound.nBits );
	DirectSound.SetBufferSize( Config.sound.nBufferSize );

	// DirectSound僶僢僼傽偺峔抸
	if( !DirectSound.InitialBuffer() ) {
		hResult = -1L;
		return	TRUE;
	}

	// 僷儗僢僩僥乕僽儖偺寁嶼
	DirectDraw.CalcPaletteTable();

	// 僱僢僩僾儗僀

	// 儊僯儏乕偺儘乕僪
	m_hMenu = CApp::LoadMenu( IDR_MENU );
	CApp::SetMenu( m_hMenu );
	::SetMenu( m_hWnd, m_hMenu );
	m_bMenu = TRUE;		// 儊僯儏乕昞帵拞僼儔僌

	// 儊僯儏乕偵傾僋僙儔儗乕僞僉乕柤偺捛壛
	OnRebuildMenu();

	// 僂僀儞僪僂埵抲偺愝掕
	RECT	rc = Config.general.rcWindowPos;
	if( !((rc.right-rc.left) <= 0 || (rc.bottom-rc.top) <= 0) ) {
		if( (m_bZoomed = Config.general.bWindowZoom) ) {
			CApp::SetCmdShow( SW_SHOWMAXIMIZED );
		}
		// 曐懚偝傟偰偄偨僂僀儞僪僂僒僀僘偵偡傞
		m_WindowRect = rc;
		::SetWindowPos( m_hWnd, HWND_NOTOPMOST, rc.left, rc.top,
				rc.right-rc.left, rc.bottom-rc.top, 0 );
	} else {
		m_bZoomed = FALSE;
		// 僨僼僅儖僩僒僀僘傪挷惍偡傞
		OnSetWindowSize();
	}

	// 僂僀儞僪僂僗僞僀儖偺曐懚
	m_StyleBackup   = ::GetWindowLong( m_hWnd, GWL_STYLE );
	m_ExStyleBackup = ::GetWindowLong( m_hWnd, GWL_EXSTYLE );

	// IME巊梡嬛巭:D
	::ImmAssociateContext( m_hWnd, NULL );

	// Drag&Drop偺嫋壜
	//::DragAcceptFiles( m_hWnd, TRUE );

	// 僼儔僌椶
	m_bActivate = TRUE;
	m_bActivateApp = TRUE;
	m_bForcus = TRUE;
	m_bCursor = TRUE;
	m_bKeyEnable = TRUE;
	m_bKeyChecking = FALSE;
	m_LastMovedTime = 0;

	m_uTimerID = ::SetTimer( m_hWnd, 0x0001, 1000, NULL );
	m_uKeyTimerID = ::SetTimer( m_hWnd, 0x0100, 30, NULL );	// 30ms

	::ZeroMemory( m_KeyBuf, sizeof(m_KeyBuf) );
	::ZeroMemory( m_KeyCnt, sizeof(m_KeyCnt) );

	//Used Memory Rom
	OnEmulationStart( NULL, FALSE );
	Emu.Event( CEmuThread::EV_THROTTLE );

	//是否放在最顶
	m_OnTop = Config.general.bOnTop;
	if(m_OnTop == 1)
		SetWindowPos(m_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE);
	else
	{
		CheckMenuItem(GetMenu(m_hWnd), 60000, MF_BYCOMMAND|MF_UNCHECKED);
	}	

#if	0
	// For Command Lines
	if( ::strlen( CApp::GetCmdLine() ) > 0 ) {
		LPSTR	pCmd = CApp::GetCmdLine();
		if( pCmd[0] == '"' ) {	// Shell execute!!
			::ZeroMemory( m_szCommandLine, sizeof(m_szCommandLine) );
			::memcpy( m_szCommandLine, pCmd+1, ::strlen(pCmd)-2 );
			m_szCommandLine[::strlen(m_szCommandLine)] = '\0';
		} else {
			::strcpy( m_szCommandLine, pCmd );
		}
		::PostMessage( m_hWnd, WM_VNS_COMMANDLINE, 0, 0L );
	}
#endif
	return	FALSE;
}

WNDMSG	CMainFrame::OnClose( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnClose\n" );

	// 僱僢僩僾儗僀

	// 僄儈儏儗乕僔儑儞廔椆
	Emu.Stop();
	DELETEPTR( Nes );

	// 僞僀儅乕偺廔椆
	::KillTimer( m_hWnd, m_uTimerID );
	::KillTimer( m_hWnd, m_uKeyTimerID );
	m_uTimerID = 0;
	m_uKeyTimerID = 0;

	//保存字体效果
	memcpy((void *)&Config.general.stLogFont, (void *)&DirectDraw.m_logFont, sizeof(LOGFONT));

	// 僂僀儞僪僂攋婞
	::DestroyWindow( m_hWnd );

	// 僒乕僠僟僀傾儘僌廔椆
	m_SearchDlg.Destroy();

	// 儔儞僠儍乕廔椆
	//m_LauncherDlg.Destroy();

	// 僷僞乕儞價儏傾廔椆
	//m_PatternView.Destroy();
	//m_NameTableView.Destroy();
	//m_PaletteView.Destroy();
	//m_MemoryView.Destroy();

	// 僂僀儞僪僂埵抲偺曐懚
	Config.general.bWindowZoom = m_bZoomed;
	Config.general.rcWindowPos = m_WindowRect;

	Config.general.bOnTop = m_OnTop;

	::DragAcceptFiles( m_hWnd, FALSE );

	return	FALSE;
}

WNDMSG	CMainFrame::OnDestroy( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnDestroy\n" );

	::PostQuitMessage( 0 );
	return	FALSE;
}

WNDMSG	CMainFrame::OnGetMinMaxInfo( WNDMSGPARAM )
{
	MINMAXINFO* lpMMI = (MINMAXINFO*)lParam;

	// 嵟彫僒僀僘偩偗
	lpMMI->ptMinTrackSize.x = 128;
	lpMMI->ptMinTrackSize.y = 128;
//	lpMMI->ptMaxTrackSize.x = 65535;
//	lpMMI->ptMaxTrackSize.y = 65535;

	return	TRUE;
}

WNDMSG	CMainFrame::OnActivate( WNDMSGPARAM )
{
	if( LOWORD(wParam) == WA_INACTIVE ) {
//		DEBUGOUT( "CMainFrame::OnActivate:Inactive\n" );
		m_bActivate = FALSE;
//		m_bActivate = TRUE;
	} else {
//		DEBUGOUT( "CMainFrame::OnActivate:Active&ClickActive\n" );
		m_bActivate = TRUE;
	}
	return	FALSE;
}

WNDMSG	CMainFrame::OnActivateApp( WNDMSGPARAM )
{
	if( (BOOL)wParam ) {
//		DEBUGOUT( "CMainFrame::OnActivateApp:Active\n" );
		m_bActivateApp = TRUE;
		if( !Config.emulator.bBackground ) {
			Emu.Resume();
		}
		DirectInput.Acquire();
	} else {
//		DEBUGOUT( "CMainFrame::OnActivateApp:Inactive\n" );
		if( m_bActivateApp ) {
			m_bActivateApp = FALSE;
			if( !Config.emulator.bBackground ) {
				Emu.Pause();
			}
		}
		DirectInput.Unacquire();
	}
	return	FALSE;
}

WNDMSG	CMainFrame::OnShortCutEnable( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnShortCutmode = %s\n", (BOOL)wParam?"TRUE":"FALSE" );

	m_bKeyEnable = (BOOL)wParam;

	return	FALSE;
}

WNDMSG	CMainFrame::OnEnable( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnEnable = %s\n", (BOOL)wParam?"TRUE":"FALSE" );

	if( (BOOL)wParam )
		Emu.Resume();
	else
		Emu.Pause();

	m_bKeyEnable = (BOOL)wParam;

	return	FALSE;
}

WNDMSG	CMainFrame::OnEnterMenuLoop( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnEnterMenuLoop\n" );
	m_bKeyEnable = FALSE;
	Emu.Pause();
	return	FALSE;
}

WNDMSG	CMainFrame::OnExitMenuLoop( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnExitMenuLoop\n" );
	m_bKeyEnable = TRUE;
	Emu.Resume();
	return	FALSE;
}

WNDMSG	CMainFrame::OnSetFocus( WNDMSGPARAM )
{
	m_bForcus = TRUE;
	if( Config.general.bScreenMode ) {
		if( Emu.IsRunning() ) {
			if( !m_bMenu ) {
				Emu.EventParam( CEmuThread::EV_FULLSCREEN_GDI, FALSE );
			}
		}
	}
	return	FALSE;
}

WNDMSG	CMainFrame::OnKillFocus( WNDMSGPARAM )
{
	m_bForcus = FALSE;
	return	FALSE;
}

WNDMSG	CMainFrame::OnInitMenu( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnInitMenu\n" );
	CRecent::UpdateMenu( CApp::GetMenu() );
	DrawMenuBar( m_hWnd );
	return	FALSE;
}

WNDMSG	CMainFrame::OnInitMenuPopup( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnInitMenuPopup\n" );
	// System menu
	if( (BOOL)HIWORD(lParam) )
		return	FALSE;

//	DEBUGOUT( "SubMenu=%08X uPos=%d bSys=%d\n", wParam, (UINT)LOWORD(lParam), (UINT)HIWORD(lParam) );

	HMENU	hMenu = (HMENU)wParam;

	INT	MenuCount = ::GetMenuItemCount( hMenu );
	for( INT i = 0; i < MenuCount; i++ ) {
		OnUpdateMenu( hMenu, ::GetMenuItemID( hMenu, i ) );
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnPaint( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnPaint\n" );
	HDC	hDC;
	PAINTSTRUCT	ps;
	hDC = ::BeginPaint( m_hWnd, &ps );
	if( !Emu.IsRunning() ) {
		RECT	rc;
		::GetClientRect( m_hWnd, &rc );
		::SetBkColor( hDC, 0x00000000 );
		::ExtTextOut( hDC, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL );
	} else if( Emu.IsPausing() ) {
		DirectDraw.OnScreenDraw();
	}
	::EndPaint( m_hWnd, &ps );

	return	TRUE;
}

WNDMSG	CMainFrame::OnDisplayChange( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnDisplayChange\n" );

	// 僼儖僗僋儕乕儞側傜偽帺暘偱僠僃儞僕拞傕偟偔偼ALT+TAB愗傝懼偊拞
	if( !Config.general.bScreenMode ) {
		// 婋尟側偺偱僗儗僢僪傪億乕僘
		Emu.Pause();
		DirectDraw.OnChangeDisplayMode();

		// 嵞昤夋
		::InvalidateRect( m_hWnd, NULL, TRUE );

		// 僗儗僢僪傪儗僕儏乕儉
		Emu.Resume();
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnPaletteChanged( WNDMSGPARAM )
{
	if( (HWND)wParam == m_hWnd )
		return	FALSE;

	DirectDraw.RealizePalette();
	return	FALSE;
}

WNDMSG	CMainFrame::OnQueryNewPalette( WNDMSGPARAM )
{
	DirectDraw.RealizePalette();
	return	TRUE;
}

WNDMSG	CMainFrame::OnMenuChar( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnMenuChar\n" );
	// 僉儞僉儞偆傞偝偄偐傜暵偠傞乧偲塕傪偮偔
	hResult = MAKELONG(0,MNC_CLOSE);
	return	TRUE;
}

WNDMSG	CMainFrame::OnKeyDown( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnKeyDown\n" );

	if( Config.general.bScreenMode ) {
		if( wParam == VK_ESCAPE && (lParam & (1<<30)) == 0 ) {
			OnChangeMenu( !m_bMenu );
			return	TRUE;
		}
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnSize( WNDMSGPARAM )
{
	// 嵟戝壔昞帵摍
	// WM_SYSCOMMAND偱傕椙偄偺偩偑丆僉儍僾僔儑儞傪僟僽儖僋儕僢僋偟偨偲偒偵棃側偄偺偱...
	switch( wParam ) {
		case	SIZE_MAXIMIZED:
			m_bZoomed = TRUE;
			{
			WINDOWPLACEMENT wpl;
			::GetWindowPlacement( m_hWnd, &wpl );
			m_WindowRect = wpl.rcNormalPosition;	// 尦偺埵抲傪僐僺乕
			}
			break;
		case	SIZE_RESTORED:
			m_bZoomed = FALSE;
			break;
		default:
			break;
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnSysCommand( WNDMSGPARAM )
{
	// 僄儈儏儗乕僔儑儞拞偺傒
	if( Emu.IsRunning() ) {
		// 僗僋儕乕儞僙乕僶乕偺婲摦偺梷巭
		if( wParam == SC_SCREENSAVE ) {
//			DEBUGOUT( "CMainFrame::OnSysCommand  SC_SCREENSAVE\n" );
			hResult = 1L;
			return	TRUE;
		}
		// 僨傿僗僾儗僀揹尮僆僼偺梷巭
		if( wParam == SC_MONITORPOWER ) {
//			DEBUGOUT( "CMainFrame::OnSysCommand  SC_MONITORPOWER\n" );
			hResult = 1L;
			return	TRUE;
		}
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnExitSizeMove( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnExitSizeMove\n" );

	if( !Config.general.bScreenMode ) {
		::GetWindowRect( m_hWnd, &m_WindowRect );
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnSetCursor( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnSetCursor\n" );
	if( LOWORD(lParam) == HTCLIENT ) {
		if( Emu.IsRunning() ) {
//			if( !Emu.IsPausing() && !Emu.IsEmuPausing() ) {
			if( !Emu.IsPausing() ) {
				if( DirectDraw.GetZapperMode() ) {
					m_bCursor = TRUE;
					::SetCursor( NULL );
					hResult = 1L;
					return	TRUE;
				} else {
					m_bCursor = FALSE;
					m_LastMovedTime = ::timeGetTime();
				}
			} else {
				m_bCursor = TRUE;
			}
		} else {
			m_bCursor = TRUE;
		}
	} else {
		m_bCursor = TRUE;
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnTimer( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnTimer\n" );
//	DEBUGOUT( "CMainFrame::OnTimer bCursor=%s bForcus=%s\n", m_bCursor?"TRUE":"FALSE", m_bForcus?"TRUE":"FALSE" );

	if( wParam == 0x0001 ) {
	// 儅僂僗僇乕僜儖梡僞僀儅乕
		RECT	rc;
		POINT	pt;
		::GetWindowRect( m_hWnd, &rc );
		::GetCursorPos( &pt );

		if( !m_bCursor && m_bForcus && pt.x >= rc.left && pt.x <= rc.right && pt.y >= rc.top && pt.y <= rc.bottom ) {
			if( Emu.IsRunning() && !Emu.IsPausing() ) {
				if( (::timeGetTime()-m_LastMovedTime) > 1500 ) {
					::SetCursor( NULL );
				}
			}
		}
	} else if( wParam == 0x0100 ) {
	// 僉乕僠僃僢僋梡僞僀儅乕
		// 傾僋僥傿僽僂僀儞僪僂偱側偄帪傗僟僀傾儘僌摍偑弌偰偄傞帪偼僉儍儞僙儖
//		if( m_bActivate && m_bKeyEnable )
		if( m_bKeyEnable )
			OnKeyControl();
	}

	return	FALSE;
}

WNDMSG	CMainFrame::OnCopyData( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnCopyData\n" );

	COPYDATASTRUCT* pcds = (COPYDATASTRUCT*)lParam;

	OnEmulationStart( (LPSTR)pcds->lpData, FALSE );

	return	TRUE;
}

WNDMSG	CMainFrame::OnErrorMessage( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnErrorMessage\n" );

	// 僒乕僠僟僀傾儘僌廔椆
	m_SearchDlg.Destroy();

	// 僷僞乕儞價儏傾廔椆
	//m_PatternView.Destroy();
	//m_NameTableView.Destroy();
	//m_PaletteView.Destroy();
	//m_MemoryView.Destroy();

	// NetPlay愗抐

	// 僄儈儏儗乕僔儑儞廔椆
	Emu.Stop();
	DELETEPTR( Nes );

	::MessageBox( m_hWnd, (LPCTSTR)lParam, "ERROR", MB_ICONERROR|MB_OK );

	::InvalidateRect( m_hWnd, NULL, TRUE );

	return	TRUE;
}

/*WNDCMD	CMainFrame::OnExit( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnExit\n" );

	::PostMessage( m_hWnd, WM_CLOSE, 0, 0 );
}*/

WNDCMD	CMainFrame::OnHelp( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnHelp\n" );

	string	sHelp = CPathlib::MakePath( CApp::GetModulePath(), "virtuanes.chm" );

	// 柺搢側傫偱ShellExecute偱戙梡
	::ShellExecute( HWND_DESKTOP, "open", sHelp.c_str(), NULL, NULL, SW_SHOWNORMAL );
//	::HtmlHelp( m_hWnd, "virtuanes.chm", HH_DISPLAY_TOPIC, NULL ); 
}

WNDCMD	CMainFrame::OnAbout( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnAbout\n" );
	CAboutDlg dlg;
	dlg.DoModal( m_hWnd );
}

/*WNDCMD	CMainFrame::OnFileOpen( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnFileOpen\n" );

	if( Emu.IsRunning() ) {
		Emu.Pause();
	}

	if( Config.general.bScreenMode ) {
		if( Emu.IsRunning() ) {
			OnFullScreenGDI( TRUE );
		}
	}

	OPENFILENAME	ofn;
	CHAR	szFile[_MAX_PATH];

	::ZeroMemory( szFile, sizeof(szFile) );
	::ZeroMemory( &ofn, sizeof(ofn) );

	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = m_hWnd;
	ofn.lpstrFile = szFile;
	ofn.nMaxFile = sizeof(szFile);

	CHAR	szTitle[256];
	CApp::LoadString( IDS_UI_OPENROM, szTitle, sizeof(szTitle) );
	ofn.lpstrTitle = szTitle;

	ofn.lpstrFilter = "All Support Types\0*.nes;*.fds;*.nsf;*.lzh;*.zip;*.rar;*.cab\0NES ROM (*.nes)\0*.nes\0Disk Image (*.fds)\0*.fds\0NES Music File (*.nsf)\0*.nsf\0Archive File\0*.lzh;*.zip;*.rar;*.cab\0";
	ofn.nFilterIndex = 1;
	ofn.Flags = OFN_READONLY|OFN_HIDEREADONLY|OFN_EXPLORER|OFN_PATHMUSTEXIST;

	if( Config.path.bRomPath ) {
		ofn.lpstrInitialDir = Config.path.szRomPath;
	} else {
		ofn.lpstrInitialDir = CApp::GetModulePath();
	}

	if( ::GetOpenFileName( &ofn ) ) {
		OnEmulationStart( szFile, FALSE );
		if( Config.general.bScreenMode && Emu.IsRunning() ) {
			OnChangeMenu( FALSE );
		}
	} else {
		if( Emu.IsRunning() ) {
			OnFullScreenGDI( FALSE );
		}
	}
	while( Emu.IsPausing() ) {
		Emu.Resume();
	}
}

WNDCMD	CMainFrame::OnFileClose( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnFileClose\n" );

	// 僒乕僠僟僀傾儘僌廔椆
	m_SearchDlg.Destroy();

	// 僷僞乕儞價儏傾廔椆
	m_PatternView.Destroy();
	m_NameTableView.Destroy();
	m_PaletteView.Destroy();
	m_MemoryView.Destroy();

	if( Emu.IsRunning() ) {
		Emu.Stop();
		DELETEPTR( Nes );

		// 嵞昤夋
		::InvalidateRect( m_hWnd, NULL, TRUE );
	}

	if( Config.general.bScreenMode ) {
		OnChangeMenu( TRUE );
	}

}
*/

WNDMSG	CMainFrame::OnCommandLine( WNDMSGPARAM )
{
	OnEmulationStart( (LPSTR)lParam, FALSE );

	return	TRUE;
}

/*WNDMSG	CMainFrame::OnLauncherCommand( WNDMSGPARAM )
{
	if( Config.general.bScreenMode ) {
		if( m_LauncherDlg.m_hWnd ) {
			::SetWindowPos( m_LauncherDlg.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
		}
	}
	::SetForegroundWindow( m_hWnd );

	OnEmulationStart( (LPSTR)lParam, FALSE );
	return	TRUE;
}

WNDMSG	CMainFrame::OnNetPlay( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnNetPlay\n" );
	hResult = NetPlay.WndProc( hWnd, wParam, lParam );
	return	TRUE;
}

WNDMSG	CMainFrame::OnNetPlayClose( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnNetPlayClose\n" );
	NetPlay.SetMsgWnd( NULL );
	NetPlay.Disconnect();

	LPCSTR	szErrStr = CApp::GetErrorString( IDS_ERROR_NETWORKDISCONNECT );
	::PostMessage( m_hWnd, WM_VNS_ERRORMSG, 0, (LPARAM)szErrStr );
	return	TRUE;
}

WNDMSG	CMainFrame::OnNetPlayError( WNDMSGPARAM )
{
//	DEBUGOUT( "CMainFrame::OnNetPlayError\n" );
	NetPlay.SetMsgWnd( NULL );
	NetPlay.Disconnect();

	LPCSTR	szErrStr = CApp::GetErrorString( IDS_ERROR_NETWORKERROR );
	::PostMessage( m_hWnd, WM_VNS_ERRORMSG, 0, (LPARAM)szErrStr );
	return	TRUE;
}
*/

/*WNDCMD	CMainFrame::OnRecentOpen( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnRecentOpen ID=%d\n", uID-ID_MRU_FILE0 );
//	DEBUGOUT( "Fname: \"%s\"\n", CRecent::GetName( (INT)uID-ID_MRU_FILE0 ) );

	OnEmulationStart( CRecent::GetName( (INT)uID-ID_MRU_FILE0 ), FALSE );
}

WNDCMD	CMainFrame::OnRecentOpenPath( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnRecentOpenPath ID=%d\n", uID-ID_MRU_FILE0 );
//	DEBUGOUT( "Fname: \"%s\"\n", CRecent::GetPath( (INT)uID-ID_MRU_FILE0 ) );

	if( Emu.IsRunning() ) {
		Emu.Pause();
	}

	if( Config.general.bScreenMode ) {
		if( Emu.IsRunning() ) {
			OnFullScreenGDI( TRUE );
		}
	}

	OPENFILENAME	ofn;
	CHAR	szFile[_MAX_PATH];

	::ZeroMemory( szFile, sizeof(szFile) );
	::ZeroMemory( &ofn, sizeof(ofn) );

	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = m_hWnd;
	ofn.lpstrFile = szFile;
	ofn.nMaxFile = sizeof(szFile);

	CHAR	szTitle[256];
	CApp::LoadString( IDS_UI_OPENROM, szTitle, sizeof(szTitle) );
	ofn.lpstrTitle = szTitle;

	ofn.lpstrFilter = "All Support Types\0*.nes;*.fds;*.nsf;*.lzh;*.zip;*.rar;*.cab\0NES ROM (*.nes)\0*.nes\0Disk Image (*.fds)\0*.fds\0NES Music File (*.nsf)\0*.nsf\0Archive File\0*.lzh;*.zip;*.rar;*.cab\0";
	ofn.nFilterIndex = 1;
	ofn.Flags = OFN_READONLY|OFN_HIDEREADONLY|OFN_EXPLORER|OFN_PATHMUSTEXIST;

	ofn.lpstrInitialDir = CRecent::GetPath( (INT)uID-ID_MRU_PATH0 );

	if( ::GetOpenFileName( &ofn ) ) {
		OnEmulationStart( szFile, FALSE );
		if( Config.general.bScreenMode && Emu.IsRunning() ) {
			OnChangeMenu( FALSE );
		}
	} else {
		if( Emu.IsRunning() ) {
			OnFullScreenGDI( FALSE );
		}
	}
	while( Emu.IsPausing() ) {
		Emu.Resume();
	}
}
*/

/*WNDMSG	CMainFrame::OnDropFiles( WNDMSGPARAM )
{
	DEBUGOUT( "CMainFrame::OnDropFiles\n" );
	::SetForegroundWindow( m_hWnd );

	CHAR szFile[_MAX_PATH];
	::DragQueryFile( (HDROP)wParam, 0, szFile, _MAX_PATH );
	::DragFinish( (HDROP)wParam );

	INT	ret;
	if( (ret = ROM::IsRomFile( szFile )) >= 0 ) {
DEBUGOUT( "ROM僼傽僀儖 CHK=%d\n", ret );
		if( ret == IDS_ERROR_ILLEGALHEADER ) {
			if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
				return	TRUE;

			OnEmulationStart( szFile, TRUE );
			return	TRUE;
		} else if( ret == 0 ) {
			OnEmulationStart( szFile, TRUE );
			return	TRUE;
		}
	}

	if( Emu.IsRunning() ) {
		if( (ret = NES::IsStateFile( szFile, Nes->rom )) >= 0 ) {
DEBUGOUT( "僗僥乕僩僼傽僀儖 CHK=%d\n", ret );
			if( ret == IDS_ERROR_ILLEGALSTATECRC ) {
				if( Config.emulator.bCrcCheck ) {
					if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
						return	TRUE;
				}
			}
			Emu.EventParam2( CEmuThread::EV_STATE_LOAD, (INT)szFile, -1 );
		} else
		if( (ret = NES::IsMovieFile( szFile, Nes->rom )) >= 0 ) {
DEBUGOUT( "儉乕價乕僼傽僀儖 CHK=%d\n", ret );
			if( ret == IDS_ERROR_ILLEGALMOVIEOLD ) {
				::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONHAND|MB_OK );
				return	TRUE;
			} else
			if( ret == IDS_ERROR_ILLEGALMOVIEVER ) {
				if( Config.emulator.bCrcCheck ) {
					if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
						return	TRUE;
				}
			} else
			if( ret == IDS_ERROR_ILLEGALMOVIECRC ) {
				if( Config.emulator.bCrcCheck ) {
					if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
						return	TRUE;
				}
			}
			Emu.EventParam( CEmuThread::EV_MOVIE_PLAY, (INT)szFile );
		}
	}

	return	TRUE;
}
*/

void	CMainFrame::OnEmulationStart( LPCSTR szFile, BOOL bChecked )
{
	// 僒乕僠僟僀傾儘僌廔椆
	m_SearchDlg.Destroy();

	// 僷僞乕儞價儏傾廔椆
//	m_PatternView.Destroy();
//	m_NameTableView.Destroy();
//	m_PaletteView.Destroy();
//	m_MemoryView.Destroy();

	// 僄儈儏儗乕僔儑儞廔椆
//	Emu.Stop();
//	DELETEPTR( Nes );


	try {
			if(szFile != NULL)
			{
				if( !bChecked ) 
				{
					INT	ret;
					if( (ret = ROM::IsRomFile( szFile )) != 0 ) 
					{
						// 抳柦揑側僄儔乕
						if( ret == IDS_ERROR_OPEN ) 
						{
							// xxx 僼傽僀儖傪奐偗傑偣傫
							::wsprintf( szErrorString, CApp::GetErrorString(ret), szFile );
							throw	szErrorString;
						}
						if( ret == IDS_ERROR_READ ) 
						{
							throw	CApp::GetErrorString(ret);
						}
						if( ret == IDS_ERROR_UNSUPPORTFORMAT ) 
						{
							throw	CApp::GetErrorString(ret);
						}

						// YES/NO僠僃僢僋
						if( ret == IDS_ERROR_ILLEGALHEADER ) 
						{
							if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
								return;
						}
					}
			}//end of if szFile is not NULL
		}

		// 僄儈儏儗乕僔儑儞廔椆
		Emu.Stop();
		DELETEPTR( Nes );

		if( (Nes = new NES(szFile, this)) ) 
		{
			if(szFile != NULL)
				CRecent::Add( szFile );

			if( Config.general.bScreenMode ) 
			{
				DirectDraw.SetFullScreenGDI( FALSE );
			} 
			else 
			{
				if( Config.emulator.bLoadFullscreen )
				{
					::PostMessage( m_hWnd, WM_COMMAND, (WPARAM)ID_FULLSCREEN, (LPARAM)0 );
				}
			}

			// 僄儈儏儗乕僔儑儞僗儗僢僪僗僞乕僩
			Emu.Start( m_hWnd, Nes );
		} 
		else 
		{
			// xxx 僼傽僀儖傪奐偗傑偣傫
			LPCSTR	szErrStr = CApp::GetErrorString( IDS_ERROR_OPEN );
			sprintf( szErrorString, szErrStr, szFile );
			throw	szErrorString;
		}
	} 
	catch( CHAR* str ) 
	{
		::strcpy( szErrorString, str );
		PostMessage( m_hWnd, WM_VNS_ERRORMSG, 0, (LPARAM)szErrorString );
#ifndef	_DEBUG
	} 
	catch(...) 
	{
		// 晄柧側僄儔乕偑敪惗偟傑偟偨
		::strcpy( szErrorString, CApp::GetErrorString( IDS_ERROR_UNKNOWN ) );
		PostMessage( m_hWnd, WM_VNS_ERRORMSG, 0, (LPARAM)szErrorString );
#endif
	}
}

/*WNDCMD	CMainFrame::OnRomInfo( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnRomInfo\n" );

	if( !Emu.IsRunning() || !Nes )
		return;

	CRomInfoDlg dlg;

	// 儊儞僶偺愝掕
	::strcpy( dlg.m_szName, Nes->rom->GetRomName() );
	dlg.m_nMapper = Nes->rom->GetMapperNo();
	dlg.m_nPRG    = Nes->rom->GetPROM_SIZE();
	dlg.m_nCHR    = Nes->rom->GetVROM_SIZE();
	dlg.m_bMirror  = Nes->rom->IsVMIRROR();
	dlg.m_bSram    = Nes->rom->IsSAVERAM();
	dlg.m_b4Screen = Nes->rom->Is4SCREEN();
	dlg.m_bTrainer = Nes->rom->IsTRAINER();

	if( Nes->rom->GetMapperNo() < 256 && Nes->rom->GetMapperNo() != 20 ) {
		dlg.m_dwCRC    = Nes->rom->GetPROM_CRC();
		dlg.m_dwCRCALL = Nes->rom->GetROM_CRC();
		dlg.m_dwCRCCHR = Nes->rom->GetVROM_CRC();
	}

	if( !m_bMenu )
		OnFullScreenGDI( TRUE );
	dlg.DoModal( m_hWnd );
	if( !m_bMenu )
		OnFullScreenGDI( FALSE );
}*/

/*WNDCMD	CMainFrame::OnWaveRecord( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnWaveRecord\n" );

	if( !Emu.IsRunning() || !Nes )
		return;

	if( Emu.IsWaveRecord() ) {
	// 榐壒拞偩偭偨傜巭傔傞
		Emu.Event( CEmuThread::EV_WAVEREC_STOP );
	} else {
	// 弶傔偰側傜
		string	tempstr;
		if( Config.path.bWavePath ) {
			CreateDirectory( Config.path.szWavePath, NULL );
			tempstr = CPathlib::MakePathExt( Config.path.szWavePath, Nes->rom->GetRomName(), "wav" );
		} else {
			tempstr = CPathlib::MakePathExt( Nes->rom->GetRomPath(), Nes->rom->GetRomName(), "wav" );
		}
//		DEBUGOUT( "Path: %s\n", tempstr.c_str() );

		OPENFILENAME	ofn;
		CHAR	szFile[_MAX_PATH];

		::strcpy( szFile, tempstr.c_str() );
		::ZeroMemory( &ofn, sizeof(ofn) );

		ofn.lStructSize = sizeof(ofn);
		ofn.hwndOwner = m_hWnd;
		ofn.lpstrFile = szFile;
		ofn.nMaxFile = sizeof(szFile);

		CHAR	szTitle[256];
		CApp::LoadString( IDS_UI_WAVERECORD, szTitle, sizeof(szTitle) );
		ofn.lpstrTitle = szTitle;

		ofn.lpstrFilter = "Wave Files(*.wav)\0*.wav\0All Files(*.*)\0*.*\0";
		ofn.nFilterIndex = 1;
		ofn.Flags = OFN_EXPLORER|OFN_PATHMUSTEXIST|OFN_OVERWRITEPROMPT;

		if( Config.path.bWavePath ) {
			ofn.lpstrInitialDir = Config.path.szWavePath;
		} else {
			ofn.lpstrInitialDir = CApp::GetModulePath();
		}

		if( !m_bMenu )
			Emu.EventParam( CEmuThread::EV_FULLSCREEN_GDI, TRUE );

		if( ::GetSaveFileName( &ofn ) ) {
			Emu.EventParam( CEmuThread::EV_WAVEREC_START, (INT)szFile );
		}

		if( !m_bMenu )
			Emu.EventParam( CEmuThread::EV_FULLSCREEN_GDI, FALSE );
	}
}*/

/*WNDCMD	CMainFrame::OnLauncher( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnLauncher\n" );

	if( !m_LauncherDlg.m_hWnd ) {
		m_LauncherDlg.Create( NULL );
	}

	if( Config.general.bScreenMode ) {
		if( Emu.IsRunning() ) {
			Emu.EventParam( CEmuThread::EV_FULLSCREEN_GDI, TRUE );
		} else {
			DirectDraw.SetFullScreenGDI( TRUE );
		}
		// 僟僀傾儘僌傪拞墰偵堏摦偡傞:)
		RECT	rcParent, rc;
		::GetWindowRect( m_hWnd, &rcParent );
		::GetWindowRect( m_LauncherDlg.m_hWnd, &rc );
		INT x = rcParent.left+(rcParent.right-rcParent.left)/2-(rc.right-rc.left)/2;
		INT y = rcParent.top +(rcParent.bottom-rcParent.top)/2-(rc.bottom-rc.top)/2;
		::SetWindowPos( m_LauncherDlg.m_hWnd, NULL, x, y, -1, -1, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE );

		::SetWindowPos( m_LauncherDlg.m_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
	} else {
		::SetWindowPos( m_LauncherDlg.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
	}

	::ShowWindow( m_LauncherDlg.m_hWnd, SW_SHOW );
}*/

/*WNDCMD	CMainFrame::OnNetPlayConnect( WNDCMDPARAM )
{
	if( !Emu.IsRunning() || !Nes )
		return;
	if( Nes->IsMoviePlay() || Nes->IsMovieRec() )
		return;
	if( NetPlay.IsConnect() )
		return;

	Emu.Pause();

	if( !m_bMenu )
		OnFullScreenGDI( TRUE );

	CNetPlayDlg	dlg;
	if( dlg.DoModal( m_hWnd ) == IDOK ) {
		Emu.Event( CEmuThread::EV_NETPLAY_START );
	}

	if( !m_bMenu )
		OnFullScreenGDI( FALSE );

	Emu.Resume();
}

WNDCMD	CMainFrame::OnNetPlayDisconnect( WNDCMDPARAM )
{
	if( !NetPlay.IsConnect() )
		return;

	Emu.Pause();
	NetPlay.Disconnect();
	Emu.Resume();
}*/

WNDCMD	CMainFrame::OnSearch( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnSearch\n" );

	if( !Emu.IsRunning() || !Nes )
		return;

	if( !m_SearchDlg.m_hWnd ) {
		m_SearchDlg.Create( NULL );
	}

	if( Config.general.bScreenMode ) {
		if( Emu.IsRunning() ) {
			Emu.EventParam( CEmuThread::EV_FULLSCREEN_GDI, TRUE );
		} else {
			DirectDraw.SetFullScreenGDI( TRUE );
		}
		// 僟僀傾儘僌傪拞墰偵堏摦偡傞:)
		RECT	rcParent, rc;
		::GetWindowRect( m_hWnd, &rcParent );
		::GetWindowRect( m_SearchDlg.m_hWnd, &rc );
		INT x = rcParent.left+(rcParent.right-rcParent.left)/2-(rc.right-rc.left)/2;
		INT y = rcParent.top +(rcParent.bottom-rcParent.top)/2-(rc.bottom-rc.top)/2;
		::SetWindowPos( m_SearchDlg.m_hWnd, NULL, x, y, -1, -1, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE );

		::SetWindowPos( m_SearchDlg.m_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
	} else {
		::SetWindowPos( m_SearchDlg.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
	}

	::ShowWindow( m_SearchDlg.m_hWnd, SW_SHOW );
}

WNDCMD	CMainFrame::OnCheat( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnCheat\n" );

	if( !Emu.IsRunning() || !Nes )
		return;

//	if( m_SearchDlg.m_hWnd ) {
//		if( ::IsWindowVisible( m_SearchDlg.m_hWnd ) )
//			return;
//	}

	if( !m_bMenu )
		OnFullScreenGDI( TRUE );

	CCheatCodeDlg	dlg;

	dlg.DoModal( m_hWnd );

	if( !m_bMenu )
		OnFullScreenGDI( FALSE );
}

WNDCMD	CMainFrame::OnGenie( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnGenie\n" );

	if( !Emu.IsRunning() || !Nes )
		return;

	Emu.Pause();

	string	tempstr;
	if( Config.path.bCheatPath ) {
		CreateDirectory( Config.path.szCheatPath, NULL );
		tempstr = CPathlib::MakePathExt( Config.path.szCheatPath, Emu.GetNES()->rom->GetRomName(), "gen" );
	} else {
		tempstr = CPathlib::MakePathExt( Emu.GetNES()->rom->GetRomPath(), Emu.GetNES()->rom->GetRomName(), "gen" );
	}

	OPENFILENAME	ofn;
	CHAR	szFile[_MAX_PATH];

	::strcpy( szFile, tempstr.c_str() );
	::ZeroMemory( &ofn, sizeof(ofn) );

	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = m_hWnd;
	ofn.lpstrFile = szFile;
	ofn.nMaxFile = sizeof(szFile);

	ofn.lpstrFilter = "GameGeine Files(*.gen)\0*.gen\0All Files(*.*)\0*.*\0";
	ofn.nFilterIndex = 1;
	ofn.Flags = OFN_EXPLORER|OFN_PATHMUSTEXIST;

	if( Config.path.bCheatPath ) {
		ofn.lpstrInitialDir = Config.path.szCheatPath;
	} else {
		ofn.lpstrInitialDir = CApp::GetModulePath();
	}

	CHAR	szTitle[256];
	CApp::LoadString( IDS_UI_LOADGENIECODE, szTitle, sizeof(szTitle) );
	ofn.lpstrTitle = szTitle;
	if( ::GetOpenFileName( &ofn ) ) {
		Nes->GenieLoad( szFile );
	}

	Emu.Resume();
}

WNDCMD	CMainFrame::OnCheatCommand( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnCheatCommand\n" );

	if( !Emu.IsRunning() || !Nes )
		return;

	if( uID == ID_CHEAT_ENABLE ) {
		Nes->SetCheatCodeAllFlag( TRUE, TRUE );
	} else
	if( uID == ID_CHEAT_DISABLE ) {
		Nes->SetCheatCodeAllFlag( FALSE, TRUE );
	}
}

WNDCMD	CMainFrame::OnEmulatorCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnEmulatorCfg\n" );
	CEmulatorDlg dlg;
	if( dlg.DoModal( m_hWnd ) == IDOK ) {
		Emu.SetPriority( Config.emulator.nPriority );
	}
}

WNDCMD	CMainFrame::OnGraphicsCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnGraphicsCfg\n" );
	CGraphicsDlg dlg;
	if( dlg.DoModal( m_hWnd ) == IDOK ) {
		Emu.Pause();

		// 僒乕僼僃僗傪儕價儖僪偡傞昁梫偑偁傞偐偺僠僃僢僋
		BOOL	bRebuildDirectDraw = FALSE;
		BOOL	bRebuildSurface = FALSE;

		if( DirectDraw.GetUseHEL() != Config.graphics.bUseHEL )
			bRebuildDirectDraw = TRUE;
		if( DirectDraw.GetSystemMemory() != Config.graphics.bSystemMemory )
			bRebuildSurface = TRUE;

		// DirectDraw堤呒偺愝掕
		DirectDraw.SetFlipMode     ( Config.graphics.bSyncDraw );
		DirectDraw.SetAspectMode   ( Config.graphics.bAspect );
		DirectDraw.SetAllLineMode  ( Config.graphics.bAllLine );
		DirectDraw.SetMaxZoom      ( Config.graphics.bFitZoom );
		DirectDraw.SetTVFrameMode  ( Config.graphics.bTVFrame );
		DirectDraw.SetScanlineMode ( Config.graphics.bScanline );
		DirectDraw.SetScanlineColor( Config.graphics.nScanlineColor );

		DirectDraw.SetUseHEL      ( Config.graphics.bUseHEL );
		DirectDraw.SetDoubleSize  ( Config.graphics.bDoubleSize );
		DirectDraw.SetSystemMemory( Config.graphics.bSystemMemory );

		DirectDraw.SetWindowVSyncMode( Config.graphics.bWindowVSync );

		if( Config.general.bScreenMode ) {
			// 僗僋儕乕儞儌乕僪偺曄峏偑偁偭偨応崌
			if( DirectDraw.m_dwDisplayWidth  != Config.graphics.dwDisplayWidth
			 || DirectDraw.m_dwDisplayHeight != Config.graphics.dwDisplayHeight
			 || DirectDraw.m_dwDisplayDepth  != Config.graphics.dwDisplayDepth
			 || DirectDraw.m_dwDisplayRate   != Config.graphics.dwDisplayRate ) {
				if( !bRebuildDirectDraw )
					DirectDraw.BeginDisplayChange();

				// 僗僋儕乕儞儌乕僪偺愝掕
				DirectDraw.m_dwDisplayWidth  = Config.graphics.dwDisplayWidth;
				DirectDraw.m_dwDisplayHeight = Config.graphics.dwDisplayHeight;
				DirectDraw.m_dwDisplayDepth  = Config.graphics.dwDisplayDepth;
				DirectDraw.m_dwDisplayRate   = Config.graphics.dwDisplayRate;

				DirectDraw.SetScreenMode( Config.general.bScreenMode );
				if( !bRebuildDirectDraw )
					DirectDraw.EndDisplayChange();

				::SetWindowPos( m_hWnd, HWND_NOTOPMOST, 0, 0, ::GetSystemMetrics(SM_CXSCREEN),
						::GetSystemMetrics(SM_CYSCREEN), SWP_SHOWWINDOW );

				// 僷儗僢僩
				::PostMessage( m_hWnd, WM_QUERYNEWPALETTE, 0, 0 );

				// 僒乕僼僃僗偼嵞峔抸偝傟偨偺偱擇搙偼峴傢側偄
				bRebuildSurface = FALSE;
			}
		} else {
			OnSetWindowSize();
		}
		if( bRebuildDirectDraw ) {
			DirectDraw.ReleaseDDraw();
			DirectDraw.InitialDDraw( m_hWnd );
			DirectDraw.InitialSurface( Config.general.bScreenMode );
		} else
		if( bRebuildSurface ) {
			DirectDraw.ReleaseSurface();
			DirectDraw.InitialSurface( Config.general.bScreenMode );
		}

		Emu.Resume();
	}
}

WNDCMD	CMainFrame::OnSoundCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnSoundCfg\n" );
	CSoundDlg dlg;

	DWORD	rate, bits;
	DirectSound.GetSamplingRate( rate, bits );
	BOOL	bSoundOn = DirectSound.IsStreamPlay();

	if( dlg.DoModal( m_hWnd ) == IDOK ) {
		if( Config.sound.nRate != (INT)rate
		 || Config.sound.nBits != (INT)bits
		 || Config.sound.nBufferSize != DirectSound.GetBufferSize() ) {
			Emu.Pause();
			if( bSoundOn ) {
				DirectSound.StreamStop();
			}

			DirectSound.ReleaseBuffer();

			DirectSound.SetSamplingRate( Config.sound.nRate, Config.sound.nBits );
			DirectSound.SetBufferSize( Config.sound.nBufferSize );

			if( DirectSound.InitialBuffer() ) {
				if( Nes ) {
					Nes->SoundSetup();
				}
				if( bSoundOn ) {
					DirectSound.StreamPlay();
					DirectSound.StreamPause();
				}
			}
			Emu.Resume();
		}
	}
}

WNDCMD	CMainFrame::OnControllerCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnControllerCfg\n" );
	CControllerDlg dlg;

	CWndHook::SetFiltering( TRUE );
	dlg.DoModal( m_hWnd );
	CWndHook::SetFiltering( FALSE );
}

WNDCMD	CMainFrame::OnShortcutCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnShortcutCfg\n" );
	CShortcutDlg dlg;

	CWndHook::SetFiltering( TRUE );

	if( dlg.DoModal( m_hWnd ) == IDOK ) {
		::SetMenu( m_hWnd, NULL );
		::DestroyMenu( m_hMenu );
		// 儊僯儏乕偺嵞儘乕僪
		m_hMenu = CApp::LoadMenu( IDR_MENU );
		CApp::SetMenu( m_hMenu );
		::SetMenu( m_hWnd, m_hMenu );
		OnRebuildMenu();
	}

	CWndHook::SetFiltering( FALSE );
}

/*WNDCMD	CMainFrame::OnFolderCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnFolderCfg\n" );
	CFolderDlg dlg;
	dlg.DoModal( m_hWnd );
}

WNDCMD	CMainFrame::OnLanguageCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnLanguageCfg\n" );
	CLanguageDlg dlg;

	if( dlg.DoModal( m_hWnd ) == IDOK ) {
		// 儔儞僠儍乕婲摦拞側傜偽暵偠傞
		BOOL	bLauncher = FALSE;
		BOOL	bLauncherVisible = FALSE;
		if( m_LauncherDlg.m_hWnd && ::IsWindow(m_LauncherDlg.m_hWnd) ) {
			bLauncherVisible = ::IsWindowVisible( m_LauncherDlg.m_hWnd );
			bLauncher = TRUE;
			m_LauncherDlg.Destroy();
		}

		// 僒乕僠僟僀傾儘僌偼暵偠傞
		if( m_SearchDlg.m_hWnd && ::IsWindow(m_SearchDlg.m_hWnd) ) {
			m_SearchDlg.Destroy();
		}

		// 儊僯儏乕偺攋婞
		::SetMenu( m_hWnd, NULL );
		::DestroyMenu( m_hMenu );

		// 尦偺僾儔僌僀儞偺奐曻
		::FreeLibrary( CApp::GetPlugin() );
		// 怴偟偄僾儔僌僀儞偺儘乕僪
		HINSTANCE hPlugin;
		if( !(hPlugin = ::LoadLibrary( CPlugin::GetPluginPath() )) ) {
			::MessageBox( m_hWnd, "Language plug-in load failed.", "VirtuaNES", MB_ICONERROR|MB_OK );
			::PostMessage( m_hWnd, WM_CLOSE, 0, 0 );
			return;
		}
		CApp::SetPlugin( hPlugin );
		// 儊僯儏乕偺嵞儘乕僪
		m_hMenu = CApp::LoadMenu( IDR_MENU );
		CApp::SetMenu( m_hMenu );
		::SetMenu( m_hWnd, m_hMenu );

		OnRebuildMenu();

		// 儔儞僠儍乕嵞婲摦
		if( bLauncher ) {
			m_LauncherDlg.Create( NULL );
			if( bLauncherVisible ) {
				::ShowWindow( m_LauncherDlg.m_hWnd, SW_SHOW );
			}
		}

		::SetForegroundWindow( m_hWnd );
	}
}
*/

WNDCMD	CMainFrame::OnMovieCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnMovieCfg\n" );
	CMovieDlg dlg;
	dlg.DoModal( m_hWnd );
}

WNDCMD	CMainFrame::OnGameOptionCfg( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnGameOptionCfg\n" );
	CGameOptionDlg dlg;
	dlg.DoModal( m_hWnd );
}

WNDCMD	CMainFrame::OnOnTop( WNDCMDPARAM )
{
	HMENU m_menu = GetMenu(m_hWnd);
	UINT state = GetMenuState(m_menu, 60000, MF_BYCOMMAND);
	if(state == MF_CHECKED)
	{
		m_OnTop = 0;
		SetWindowPos(m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE);
		CheckMenuItem(m_menu, 60000, MF_BYCOMMAND|MF_UNCHECKED);
	}
	else if(state == MF_UNCHECKED)
	{
		m_OnTop = 1;
		SetWindowPos(m_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE);
		CheckMenuItem(m_menu, 60000, MF_BYCOMMAND|MF_CHECKED);
	}
}

WNDCMD	CMainFrame::OnFullScreen( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnFullScreen\n" );

	// 婋尟側偺偱...
	Emu.Pause();

	Config.general.bScreenMode = !Config.general.bScreenMode;
//	OnChangeMenu( FALSE );

	if( !Config.general.bScreenMode ) {
	// FullScreen to Window mode
		DirectDraw.BeginDisplayChange();

		// 埵抲偺暅婣
		::SetWindowLong( m_hWnd, GWL_STYLE, m_StyleBackup );
		::SetWindowLong( m_hWnd, GWL_EXSTYLE, m_ExStyleBackup );
		::SetWindowPlacement( m_hWnd, &m_WindowPlacement );
	} else {
	// Window to FullScreen mode
		// 埵抲偺曐懚
//		m_bZoomed = ::IsZoomed( m_hWnd );
		m_WindowPlacement.length = sizeof(m_WindowPlacement);
		::GetWindowPlacement( m_hWnd, &m_WindowPlacement );

		m_StyleBackup   = ::GetWindowLong( m_hWnd, GWL_STYLE );
		m_ExStyleBackup = ::GetWindowLong( m_hWnd, GWL_EXSTYLE );
		::SetWindowLong( m_hWnd, GWL_STYLE, WS_POPUP );
		::SetWindowLong( m_hWnd, GWL_EXSTYLE, 0 );

		// 僗僋儕乕儞儌乕僪偺愝掕
		DirectDraw.m_dwDisplayWidth  = Config.graphics.dwDisplayWidth;
		DirectDraw.m_dwDisplayHeight = Config.graphics.dwDisplayHeight;
		DirectDraw.m_dwDisplayDepth  = Config.graphics.dwDisplayDepth;
		DirectDraw.m_dwDisplayRate   = Config.graphics.dwDisplayRate;

		DirectDraw.BeginDisplayChange();
	}

//	OnChangeMenu( FALSE );

	DirectDraw.SetScreenMode( Config.general.bScreenMode );
	DirectDraw.EndDisplayChange();

	if( !Config.general.bScreenMode ) {
		/*if( m_LauncherDlg.m_hWnd && ::IsWindow(m_LauncherDlg.m_hWnd) ) {
			::SetWindowPos( m_LauncherDlg.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE );	
		}*/

		OnChangeMenu( TRUE );
		::SetForegroundWindow( m_hWnd );

		::RedrawWindow( m_hWnd, NULL, NULL, RDW_FRAME|RDW_INVALIDATE );
	} else {
		if( Emu.IsRunning() ) {
			OnChangeMenu( FALSE );
		} else {
			OnChangeMenu( TRUE );
		}
		::SetWindowPos( m_hWnd, HWND_NOTOPMOST, 0, 0, ::GetSystemMetrics(SM_CXSCREEN),
				::GetSystemMetrics(SM_CYSCREEN), SWP_SHOWWINDOW );
	}

	// 婋尟側偺偱...
	Emu.Resume();
}

WNDCMD	CMainFrame::OnZoom( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnZoom %d\n", uID-ID_ZOOMx1 );

	// 僼儖僗僋儕乕儞帪偼柍帇
	if( Config.general.bScreenMode )
		return;

	Config.general.nScreenZoom = (INT)(uID-ID_ZOOMx1);
	OnSetWindowSize();

	CHAR	szStr[64];
	::wsprintf( szStr, "Screen Zoom *%d", uID-ID_ZOOMx1+1 );
	Emu.EventParam( CEmuThread::EV_MESSAGE_OUT, (LONG)szStr );
}

/*WNDCMD	CMainFrame::OnViewCommand( WNDCMDPARAM )
{
	if( !Emu.IsRunning() || !Nes )
		return;

	switch( uID ) {
		case	ID_VIEW_PATTERN:
			if( !m_PatternView.m_hWnd ) {
				m_PatternView.Create( HWND_DESKTOP );
			}
			::SetWindowPos( m_PatternView.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
			break;
		case	ID_VIEW_NAMETABLE:
			if( !m_NameTableView.m_hWnd ) {
				m_NameTableView.Create( HWND_DESKTOP );
			}
			::SetWindowPos( m_NameTableView.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
			break;
		case	ID_VIEW_PALETTE:
			if( !m_PaletteView.m_hWnd ) {
				m_PaletteView.Create( HWND_DESKTOP );
			}
			::SetWindowPos( m_PaletteView.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
			break;

		case	ID_VIEW_MEMORY:
			if( !m_MemoryView.m_hWnd ) {
				m_MemoryView.Create( HWND_DESKTOP );
			}
			::SetWindowPos( m_MemoryView.m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE );
			break;
		default:
			break;
	}
}
*/

WNDCMD	CMainFrame::OnEmuCommand( WNDCMDPARAM )
{
//	DEBUGOUT( "CMainFrame::OnEmuCommand %d\n", uID );

	switch( uID ) {
		case	ID_HWRESET:
			Emu.Event( CEmuThread::EV_HWRESET );
			break;
		case	ID_SWRESET:
			Emu.Event( CEmuThread::EV_SWRESET );
			break;
		case	ID_PAUSE:
			Emu.Event( CEmuThread::EV_EMUPAUSE );
			break;

		case	ID_ONEFRAME:
			Emu.Event( CEmuThread::EV_ONEFRAME );
			break;

		case	ID_THROTTLE:
			Emu.Event( CEmuThread::EV_THROTTLE );
			break;
		case	ID_FRAMESKIP_UP:
			Emu.Event( CEmuThread::EV_FRAMESKIP_UP );
			break;
		case	ID_FRAMESKIP_DOWN:
			Emu.Event( CEmuThread::EV_FRAMESKIP_DOWN );
			break;
		case	ID_FRAMESKIP_AUTO:
			Emu.Event( CEmuThread::EV_FRAMESKIP_AUTO );
			break;

		case	ID_STATE_UP:
			{
			CHAR	szStr[64];
			if( ++m_nStateSlot > 10-1 ) {
				m_nStateSlot = 0;
			}
			::wsprintf( szStr, "State Slot #%d", m_nStateSlot );
			Emu.EventParam( CEmuThread::EV_MESSAGE_OUT, (LONG)szStr );
			}
			break;
		case	ID_STATE_DOWN:
			{
			CHAR	szStr[64];
			if( --m_nStateSlot < 0 ) {
				m_nStateSlot = 10-1;
			}
			::wsprintf( szStr, "State Slot #%d", m_nStateSlot );
			Emu.EventParam( CEmuThread::EV_MESSAGE_OUT, (LONG)szStr );
			}
			break;

		case	ID_STATE_SLOT0:	case	ID_STATE_SLOT1:
		case	ID_STATE_SLOT2:	case	ID_STATE_SLOT3:
		case	ID_STATE_SLOT4:	case	ID_STATE_SLOT5:
		case	ID_STATE_SLOT6:	case	ID_STATE_SLOT7:
		case	ID_STATE_SLOT8:	case	ID_STATE_SLOT9:
			m_nStateSlot = (INT)(uID-ID_STATE_SLOT0);
			{
			CHAR	szStr[64];
			::wsprintf( szStr, "State Slot #%d", m_nStateSlot );
			Emu.EventParam( CEmuThread::EV_MESSAGE_OUT, (LONG)szStr );
			}
			break;

		/*
		case	ID_DISK_EJECT:
		case	ID_DISK_0A: case	ID_DISK_0B:
		case	ID_DISK_1A: case	ID_DISK_1B:
			Emu.EventParam( CEmuThread::EV_DISK_COMMAND, (LONG)(uID-ID_DISK_EJECT) );
			break;
		*/

		case	ID_MUTE_0: case	ID_MUTE_1: case	ID_MUTE_2: case	ID_MUTE_3:
		case	ID_MUTE_4: case	ID_MUTE_5: case	ID_MUTE_6: case	ID_MUTE_7:
		case	ID_MUTE_8: case	ID_MUTE_9: case	ID_MUTE_A: case	ID_MUTE_B:
		case	ID_MUTE_C: case	ID_MUTE_D: case	ID_MUTE_E: case	ID_MUTE_F:
			Emu.EventParam( CEmuThread::EV_SOUND_MUTE, (LONG)(uID-ID_MUTE_0) );
			break;

		/*
		case	ID_EXCTR_NONE:
		case	ID_EXCTR_PADDLE:
		case	ID_EXCTR_HYPERSHOT:
		case	ID_EXCTR_ZAPPER:
		case	ID_EXCTR_KEYBOARD:
		case	ID_EXCTR_CRAZYCLIMBER:
		case	ID_EXCTR_TOPRIDER:
			Emu.EventParam( CEmuThread::EV_EXCONTROLLER, uID-ID_EXCTR_NONE );
			break;
		*/

		case	ID_SNAPSHOT:
			Emu.Event( CEmuThread::EV_SNAPSHOT );
			break;

		case	ID_TVASPECT:
			Config.graphics.bAspect = !Config.graphics.bAspect;
			Emu.Pause();
			DirectDraw.SetAspectMode( Config.graphics.bAspect );
			if( !Config.general.bScreenMode )
				OnSetWindowSize();
			Emu.Resume();
			break;
		case	ID_SCANLINE:
			Config.graphics.bScanline = !Config.graphics.bScanline;
			Emu.Pause();
			DirectDraw.SetScanlineMode( Config.graphics.bScanline );
			Emu.Resume();
			break;
		case	ID_ALLLINE:
			Config.graphics.bAllLine = !Config.graphics.bAllLine;
			Emu.Pause();
			DirectDraw.SetAllLineMode( Config.graphics.bAllLine );
			if( !Config.general.bScreenMode )
				OnSetWindowSize();
			Emu.Resume();
			break;
		case	ID_ALLSPRITE:
			Config.graphics.bAllSprite = !Config.graphics.bAllSprite;
			break;
		case	ID_SYNCDRAW:
			Config.graphics.bSyncDraw = !Config.graphics.bSyncDraw;
			Emu.Pause();
			DirectDraw.SetFlipMode( Config.graphics.bSyncDraw );
			Emu.Resume();
			break;
		case	ID_FITSCREEN:
			Config.graphics.bFitZoom = !Config.graphics.bFitZoom;
			Emu.Pause();
			DirectDraw.SetMaxZoom( Config.graphics.bFitZoom );
			Emu.Resume();
			break;

		case	ID_TVFRAME:
			Config.graphics.bTVFrame = !Config.graphics.bTVFrame;
			Emu.Pause();
			DirectDraw.SetTVFrameMode( Config.graphics.bTVFrame );
			Emu.Resume();
			break;
		case	ID_FPSDISP:
			Config.graphics.bFPSDisp = !Config.graphics.bFPSDisp;
			break;

		case	ID_LEFTCLIP:
			Config.graphics.bLeftClip = !Config.graphics.bLeftClip;
			break;

		default:
			break;
	}
}

WNDCMD	CMainFrame::OnStateCommand( WNDCMDPARAM )
{
	if( !Emu.IsRunning() )
		return;

	CHAR	st[16];
	::wsprintf( st, "st%1X", m_nStateSlot );

	string	tempstr;
	if( Config.path.bStatePath ) {
		::CreateDirectory( Config.path.szStatePath, NULL );
		tempstr = CPathlib::MakePathExt( Config.path.szStatePath, Nes->rom->GetRomName(), st );
	} else {
		tempstr = CPathlib::MakePathExt( Nes->rom->GetRomPath(), Nes->rom->GetRomName(), st );
	}

	if( uID == ID_STATE_LOAD ) {
		INT	ret;
		if( (ret = NES::IsStateFile( tempstr.c_str(), Nes->rom )) >= 0 ) {
			if( ret == IDS_ERROR_ILLEGALSTATECRC ) {
				if( Config.emulator.bCrcCheck ) {
					if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
						return;
				}
			}
		}
		Emu.EventParam2( CEmuThread::EV_STATE_LOAD, (INT)tempstr.c_str(), m_nStateSlot );
	}
	if( uID == ID_STATE_SAVE ) {
		Emu.EventParam2( CEmuThread::EV_STATE_SAVE, (INT)tempstr.c_str(), m_nStateSlot );
	}
}

WNDCMD	CMainFrame::OnMovieCommand( WNDCMDPARAM )
{
	if( !Emu.IsRunning() )
		return;

	if( uID == ID_MOVIE_STOP ) {
		Emu.Event( CEmuThread::EV_MOVIE_STOP );
		return;
	}

	Emu.Pause();

	if( Config.general.bScreenMode && !m_bMenu ) {
		OnFullScreenGDI( TRUE );
	}

	string	tempstr;
	if( Config.path.bMoviePath ) {
		CreateDirectory( Config.path.szMoviePath, NULL );
		tempstr = CPathlib::MakePathExt( Config.path.szMoviePath, Nes->rom->GetRomName(), "vmv" );
	} else {
		tempstr = CPathlib::MakePathExt( Nes->rom->GetRomPath(), Nes->rom->GetRomName(), "vmv" );
	}
//	DEBUGOUT( "Path: %s\n", tempstr.c_str() );

	OPENFILENAME	ofn;
	CHAR	szFile[_MAX_PATH];

	::strcpy( szFile, tempstr.c_str() );
	::ZeroMemory( &ofn, sizeof(ofn) );

	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = m_hWnd;
	ofn.lpstrFile = szFile;
	ofn.nMaxFile = sizeof(szFile);

	ofn.lpstrFilter = "Movie Files(*.vmv)\0*.vmv\0All Files(*.*)\0*.*\0";
	ofn.nFilterIndex = 1;
	ofn.Flags = OFN_EXPLORER|OFN_PATHMUSTEXIST;

	if( Config.path.bMoviePath ) {
		ofn.lpstrInitialDir = Config.path.szMoviePath;
	} else {
		ofn.lpstrInitialDir = CApp::GetModulePath();
	}

	CHAR	szTitle[256];

	if( uID == ID_MOVIE_PLAY ) {
		CApp::LoadString( IDS_UI_PLAYMOVIE, szTitle, sizeof(szTitle) );
		ofn.lpstrTitle = szTitle;
		if( ::GetOpenFileName( &ofn ) ) {
			INT	ret;
			if( (ret = NES::IsMovieFile( szFile, Nes->rom )) >= 0 ) {
				if( ret == IDS_ERROR_ILLEGALMOVIEOLD ) {
					::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONHAND|MB_OK );
					goto	_Movie_Play_Failed;
				} else
				if( ret == IDS_ERROR_ILLEGALMOVIEVER ) {
					if( Config.emulator.bCrcCheck ) {
						if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
							goto	_Movie_Play_Failed;
					}
				} else
				if( ret == IDS_ERROR_ILLEGALMOVIECRC ) {
					if( Config.emulator.bCrcCheck ) {
						if( ::MessageBox( m_hWnd, CApp::GetErrorString(ret), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
							goto	_Movie_Play_Failed;
					}
				}

				Emu.EventParam( CEmuThread::EV_MOVIE_PLAY, (INT)szFile );
			}
_Movie_Play_Failed:;
		}
	}
	if( uID == ID_MOVIE_REC ) {
		ofn.Flags |= OFN_OVERWRITEPROMPT;
		CApp::LoadString( IDS_UI_RECMOVIE, szTitle, sizeof(szTitle) );
		ofn.lpstrTitle = szTitle;
		if( ::GetSaveFileName( &ofn ) ) {
			Emu.EventParam( CEmuThread::EV_MOVIE_REC, (INT)szFile );
		}
	}
	if( uID == ID_MOVIE_REC_APPEND ) {
		CApp::LoadString( IDS_UI_APPENDMOVIE, szTitle, sizeof(szTitle) );
		ofn.lpstrTitle = szTitle;
		if( ::GetOpenFileName( &ofn ) ) {
			INT	ret;
			if( (ret = NES::IsMovieFile( szFile, Nes->rom )) >= 0 ) {
				if( ret == IDS_ERROR_ILLEGALMOVIEOLD ) {
					::MessageBox( m_hWnd, CApp::GetErrorString(IDS_ERROR_ILLEGALMOVIEOLD_A), "VirtuaNES", MB_ICONHAND|MB_OK );
					goto	_Movie_Append_Failed;
				} else
				if( ret == IDS_ERROR_ILLEGALMOVIEVER ) {
					if( Config.emulator.bCrcCheck ) {
						if( ::MessageBox( m_hWnd, CApp::GetErrorString(IDS_ERROR_ILLEGALMOVIEVER_A), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
							goto	_Movie_Append_Failed;
					}
				} else
				if( ret == IDS_ERROR_ILLEGALMOVIECRC ) {
					if( Config.emulator.bCrcCheck ) {
						if( ::MessageBox( m_hWnd, CApp::GetErrorString(IDS_ERROR_ILLEGALMOVIECRC_A), "VirtuaNES", MB_ICONWARNING|MB_YESNO ) != IDYES )
							goto	_Movie_Append_Failed;
					}
				}

				Emu.EventParam( CEmuThread::EV_MOVIE_RECAPPEND, (INT)szFile );
			} else {
				// 怴婯嶌惉偲摨偠
				Emu.EventParam( CEmuThread::EV_MOVIE_REC, (INT)szFile );
			}
		}
_Movie_Append_Failed:;
	}

	if( Config.general.bScreenMode && !m_bMenu ) {
		OnFullScreenGDI( FALSE );
	}

	Emu.Resume();
}

void	CMainFrame::OnMovieInfo( WNDCMDPARAM )
{
	if( !Emu.IsRunning() || !Nes )
		return;

	if( !(Nes->IsMoviePlay() || Nes->IsMovieRec()) )
		return;

	CMovieInfoDlg dlg;

	// 儊儞僶偺愝掕
	Nes->GetMovieInfo( dlg.m_wRecVersion, dlg.m_wVersion, dlg.m_dwFrames, dlg.m_dwRerecordTimes );

	if( !m_bMenu )
		OnFullScreenGDI( TRUE );
	dlg.DoModal( m_hWnd );
	if( !m_bMenu )
		OnFullScreenGDI( FALSE );
}

void	CMainFrame::OnFullScreenGDI( BOOL bFlag )
{
	if( Config.general.bScreenMode ) {
		if( Emu.IsRunning() ) {
			Emu.EventParam( CEmuThread::EV_FULLSCREEN_GDI, (INT)bFlag );
		}
	}
}

void	CMainFrame::OnShowCursor( BOOL bFlag )
{
	if( !bFlag ) {
		if( m_bCursor ) {
			while( ::ShowCursor( FALSE ) >= 0 );
			m_bCursor = FALSE;
		}
	} else {
		if( !m_bCursor ) {
			while( ::ShowCursor( TRUE ) < 0 );
			m_bCursor = TRUE;
		}
	}
}

void	CMainFrame::OnChangeMenu( BOOL bFlag )
{
	if( Config.general.bScreenMode ) {
		if( Emu.IsRunning() ) {
			Emu.EventParam( CEmuThread::EV_FULLSCREEN_GDI, (INT)bFlag );
		} else {
			DirectDraw.SetFullScreenGDI( bFlag );
		}
	}

	if( m_bMenu ) {
		if( !bFlag ) {
			::SetMenu( m_hWnd, NULL );
			m_bMenu = FALSE;
		}
	} else {
		if( bFlag ) {
			::SetMenu( m_hWnd, m_hMenu );
			m_bMenu = TRUE;
			::DrawMenuBar( m_hWnd );
		}
	}

//	OnShowCursor( m_bMenu );
}

void	CMainFrame::OnSetWindowSize()
{
	// 僂僀儞僪僂儌乕僪偺帪偩偗
	if( Config.general.bScreenMode )
		return;

	LONG	width, height;

	if( !Config.graphics.bAspect )  width  = CDirectDraw::SCREEN_WIDTH;
	else				width  = (LONG)((CDirectDraw::SCREEN_WIDTH)*1.25);
	if( !Config.graphics.bAllLine ) height = CDirectDraw::SCREEN_HEIGHT-16;
	else				height = CDirectDraw::SCREEN_HEIGHT;

	width  *= Config.general.nScreenZoom+1;
	height *= Config.general.nScreenZoom+1;

	RECT	rcW, rcC;
	::GetWindowRect( m_hWnd, &rcW );
	::GetClientRect( m_hWnd, &rcC );

	rcW.right  += width -(rcC.right-rcC.left);
	rcW.bottom += height-(rcC.bottom-rcC.top);
	::SetWindowPos( m_hWnd, HWND_NOTOPMOST, rcW.left, rcW.top,
			rcW.right-rcW.left, rcW.bottom-rcW.top, SWP_NOZORDER );

	// 儊僯儏乕偑愜傝曉偝傟傞偲廲偑彫偝偔側傞帠偑偁傞偺偱嵞廋惓
	::GetClientRect( m_hWnd, &rcC );
	if( (rcC.bottom-rcC.top) != height ) {
		::GetWindowRect( m_hWnd, &rcW );
		::GetClientRect( m_hWnd, &rcC );

		rcW.right  += width -(rcC.right-rcC.left);
		rcW.bottom += height-(rcC.bottom-rcC.top);
		::SetWindowPos( m_hWnd, HWND_NOTOPMOST, rcW.left, rcW.top,
				rcW.right-rcW.left, rcW.bottom-rcW.top, SWP_NOZORDER );
	}


	// 嵟戝昞帵帪偵僒僀僘傪曄峏偟偨応崌偺嵟戝昞帵偺夝彍媦傃儃僞儞偺嵞昤夋
	LONG	style = ::GetWindowLong( m_hWnd, GWL_STYLE );
	style &= ~WS_MAXIMIZE;
	::SetWindowLong( m_hWnd, GWL_STYLE, style );
	::RedrawWindow( m_hWnd, NULL, NULL, RDW_FRAME|RDW_INVALIDATE );

	// 埵抲曐懚偟側偍偟
	m_bZoomed = FALSE;
	::GetWindowRect( m_hWnd, &m_WindowRect );
}

void	CMainFrame::OnRebuildMenu()
{
	CHAR	szMenuString[256];
	string	str;
	for( INT i = 0; CConfig::ShortcutKeyID[i*3+0]; i++ ) {
		::GetMenuString( m_hMenu, CConfig::ShortcutKeyID[i*3+0], szMenuString, 256, MF_BYCOMMAND );

		INT	id, code;
		id = CConfig::ShortcutKeyID[i*3+2];
		code = Config.shortcut.nShortCut[id];
		if( code != 0 ) {
			string	temp = szMenuString;
			temp = temp + "\t";
			temp = temp + Config.ShortcutToKeyName( code );

			::ModifyMenu( m_hMenu, CConfig::ShortcutKeyID[i*3+0], MF_BYCOMMAND|MF_STRING,
					CConfig::ShortcutKeyID[i*3+0], temp.c_str() );
		}
	}
}

void	CMainFrame::OnUpdateMenu( HMENU hMenu, UINT uID )
{
	BOOL	bEnable, bCheck;
	BOOL	bEmu = Emu.IsRunning();
	BOOL	bScn = Config.general.bScreenMode;

	switch( uID ) {
		case	ID_CLOSE:
		case	ID_ROMINFO:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_WAVERECORD:
			bCheck = bEmu && Emu.IsWaveRecord();
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			::CheckMenuItem( hMenu, uID, MF_BYCOMMAND|(bCheck?MF_CHECKED:MF_UNCHECKED) );
			break;

		case	ID_HWRESET:
		case	ID_SWRESET:
		case	ID_PAUSE:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;
		case	ID_STATE_LOAD:
		case	ID_STATE_SAVE:
			if( Nes ) {
				bEnable = bEmu && !Nes->rom->IsNSF();
			} else {
				bEnable = bEmu;
			}
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEnable?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_STATE_SLOT0:	case	ID_STATE_SLOT1:
		case	ID_STATE_SLOT2:	case	ID_STATE_SLOT3:
		case	ID_STATE_SLOT4:	case	ID_STATE_SLOT5:
		case	ID_STATE_SLOT6:	case	ID_STATE_SLOT7:
		case	ID_STATE_SLOT8:	case	ID_STATE_SLOT9:
			// 擔帪昞帵(寢峔柺搢倵)
			{
			CHAR	szMenuString[256];
			CHAR*	pToken;
			const UCHAR seps[] = " \t\0";	// 僙僷儗乕僞

			// ID斣崋偐傜僀儞僨僢僋僗傪扵偡
			for( INT i = 0; CConfig::ShortcutKeyID[i*3+0] != uID; i++ );

			::GetMenuString( m_hMenu, CConfig::ShortcutKeyID[i*3+0], szMenuString, 256, MF_BYCOMMAND );

			if( (pToken = (CHAR*)_mbstok( (UCHAR*)szMenuString, seps )) ) {
				string	str = pToken;

				if( Emu.IsRunning() && Nes ) {
					CHAR	temp[256];
					::wsprintf( temp, "st%1X", CConfig::ShortcutKeyID[i*3+0]-ID_STATE_SLOT0 );

					string	tempstr;
					if( Config.path.bStatePath ) {
						::CreateDirectory( Config.path.szStatePath, NULL );
						tempstr = CPathlib::MakePathExt( Config.path.szStatePath, Nes->rom->GetRomName(), temp );
					} else {
						tempstr = CPathlib::MakePathExt( Nes->rom->GetRomPath(), Nes->rom->GetRomName(), temp );
					}

					HANDLE	hFile = ::CreateFile( tempstr.c_str(), GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL );
					if( hFile != INVALID_HANDLE_VALUE ) {
						BY_HANDLE_FILE_INFORMATION fi;
						FILETIME	filetime;
						SYSTEMTIME	systime;
						::GetFileInformationByHandle( hFile, &fi );
						::FileTimeToLocalFileTime( &fi.ftLastWriteTime, &filetime );
						::FileTimeToSystemTime( &filetime, &systime );
						::wsprintf( temp,"%04d/%02d/%02d %02d:%02d:%02d", systime.wYear, systime.wMonth, systime.wDay, systime.wHour, systime.wMinute, systime.wSecond );
						::CloseHandle( hFile );

						str = str + " ";
						str = str + temp;
					} else {
						str = str + " not exist";
					}
				}

				INT	code = Config.shortcut.nShortCut[CConfig::ShortcutKeyID[i*3+2]];
				if( code ) {
					str = str + "\t";
					str = str + Config.ShortcutToKeyName( code );
				}

				::ModifyMenu( m_hMenu, CConfig::ShortcutKeyID[i*3+0], MF_BYCOMMAND|MF_STRING,
						CConfig::ShortcutKeyID[i*3+0], str.c_str() );
			}

			::CheckMenuItem( hMenu, uID, MF_BYCOMMAND|(m_nStateSlot==(INT)(uID-ID_STATE_SLOT0)?MF_CHECKED:MF_UNCHECKED) );
			}
			break;

		case	ID_ZOOMx1: case	ID_ZOOMx2:
		case	ID_ZOOMx3: case	ID_ZOOMx4:
			if( bScn ) ::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|MF_GRAYED );
			else	   ::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|MF_ENABLED );
			::CheckMenuItem( hMenu, uID, MF_BYCOMMAND|(Config.general.nScreenZoom==(uID-ID_ZOOMx1)?MF_CHECKED:MF_UNCHECKED) );
			break;

		case	ID_DISK_EJECT:
		case	ID_DISK_0A: case	ID_DISK_0B:
		case	ID_DISK_1A: case	ID_DISK_1B:
			if( Nes ) {
				INT no = Nes->GetDiskNo();
				bEnable = bEmu && !(Nes->rom->GetMapperNo() != 20 || Nes->rom->IsNSF() || (!no || (uID-ID_DISK_EJECT)>no));
			} else {
				bEnable = bEmu;
			}
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEnable?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_MOVIE_PLAY:
		case	ID_MOVIE_REC:
		case	ID_MOVIE_REC_APPEND:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|((bEmu)?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_MOVIE_STOP:
		case	ID_MOVIE_INFO:
			if( Nes ) {
				bEnable = Nes->IsMoviePlay() || Nes->IsMovieRec();
			} else {
				bEnable = bEmu;
			}
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEnable?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_SNAPSHOT:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_EXCTR_NONE:
		case	ID_EXCTR_PADDLE:
		case	ID_EXCTR_HYPERSHOT:
		case	ID_EXCTR_ZAPPER:
		case	ID_EXCTR_KEYBOARD:
		case	ID_EXCTR_CRAZYCLIMBER:
		case	ID_EXCTR_TOPRIDER:
		case	ID_EXCTR_SPACESHADOWGUN:
			if( Nes ) {
				if( Nes->pad->GetExController() == (uID-ID_EXCTR_NONE) )
					bCheck = TRUE;
				else
					bCheck = FALSE;
			} else {
				if( uID == ID_EXCTR_NONE )
					bCheck = TRUE;
				else
					bCheck = FALSE;
			}
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			::CheckMenuItem( hMenu, uID, MF_BYCOMMAND|(bCheck?MF_CHECKED:MF_UNCHECKED) );
			break;

		case	ID_FULLSCREEN:
			::CheckMenuItem( hMenu, uID, MF_BYCOMMAND|(bScn?MF_CHECKED:MF_UNCHECKED) );
			break;
		case	ID_FPSDISP:
			::CheckMenuItem( hMenu, uID, MF_BYCOMMAND|(Config.graphics.bFPSDisp?MF_CHECKED:MF_UNCHECKED) );
			break;
		case	ID_TVFRAME:
			::CheckMenuItem( hMenu, uID, MF_BYCOMMAND|(Config.graphics.bTVFrame?MF_CHECKED:MF_UNCHECKED) );
			break;

		case	ID_CFG_GAMEOPTION:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_SEARCH:
		case	ID_CHEAT:
		case	ID_CHEAT_ENABLE:
		case	ID_CHEAT_DISABLE:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;
#if	0
		case	ID_CHEAT:
			if( m_SearchDlg.m_hWnd )
				bEnable = bEmu && !::IsWindowVisible( m_SearchDlg.m_hWnd );
			else
				bEnable = bEmu;
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEnable?MF_ENABLED:MF_GRAYED) );
			break;
#endif
		case	ID_GENIE:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_VIEW_PATTERN:
		case	ID_VIEW_NAMETABLE:
		case	ID_VIEW_PALETTE:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;

		case	ID_VIEW_MEMORY:
		case	ID_VIEW_WATCH:
			::EnableMenuItem( hMenu, uID, MF_BYCOMMAND|(bEmu?MF_ENABLED:MF_GRAYED) );
			break;

		default:
			break;
	}
}

void	CMainFrame::OnKeyControl()
{
BYTE	KeyBuf[256+64*8];

	// 僞僀儅乕僀儀儞僩偑廳側傞帠偑偁傞偺偐偳偆偐傢偐傜傫偑偲傝偁偊偢
	if( m_bKeyChecking )
		return;

	m_bKeyChecking = TRUE;

	// 僉乕偺庢摼
	if( !Emu.IsRunning() ) {
		// 僄儈儏儗乕僔儑儞拞埲奜偼億乕儕儞僌偡傞
		DirectInput.Poll();
	}

	// 僩儕僈仌儕僺乕僩惗惉
	BYTE*	pSw  = (BYTE*)DirectInput.m_Sw;
	BYTE*	pOld = (BYTE*)m_KeyBuf;
	BYTE*	pCnt = (BYTE*)m_KeyCnt;
	BYTE*	pKey = (BYTE*)KeyBuf;

	for( INT i = 0; i < 256+64*8; i++ ) {
		pKey[i] = (pSw[i]&0x80) & ~(pOld[i]&0x80);
		if( pKey[i] ) {
			pCnt[i] = 14;
		} else if( pCnt[i] != 0 && (pSw[i]&0x80) ) {
			if( --pCnt[i] == 0 ) {
				pKey[i] = pSw[i]&0x80;
				pCnt[i] = 5;
			} else {
				pKey[i] = 0;
			}
		}
	}
	::memcpy( m_KeyBuf, DirectInput.m_Sw, sizeof(m_KeyBuf) );

	// 僀儀儞僩敪怣
	BOOL	bAltOnly = FALSE;
	if( Emu.IsRunning() ) {
		// 僼傽儈儕乕儀乕僔僢僋僉乕儃乕僪偺帪
		if( Emu.GetExController() == PAD::EXCONTROLLER_KEYBOARD )
			bAltOnly = TRUE;
		if( Emu.IsPausing() )
			bAltOnly = FALSE;
	}

	WORD	keyex = 0;
	if( DirectInput.m_Sw[DIK_LMENU] || DirectInput.m_Sw[DIK_RMENU] )
		keyex = CCfgShortCut::K_ALT;
	if( DirectInput.m_Sw[DIK_LCONTROL] || DirectInput.m_Sw[DIK_RCONTROL] )
		keyex = CCfgShortCut::K_CTRL;
	if( DirectInput.m_Sw[DIK_LSHIFT] || DirectInput.m_Sw[DIK_RSHIFT] )
		keyex = CCfgShortCut::K_SHIFT;

	WORD*	pShortCutKey = Config.shortcut.nShortCut;
	INT*	pShortCutKeyID = Config.ShortcutKeyID;

	if( bAltOnly ) {
		INT	no;
		WORD	key, key2;
		for( INT i = 0; pShortCutKeyID[i*3+0]; i++ ) {
			no = pShortCutKeyID[i*3+2];
			key = pShortCutKey[no];
			key2 = pShortCutKey[no+128];

			if( (key & 0xF000) == CCfgShortCut::K_ALT && keyex == CCfgShortCut::K_ALT && KeyBuf[key&0x0FFF] && (key&0x0FFF) ) {
				::PostMessage( m_hWnd, WM_COMMAND, (WPARAM)pShortCutKeyID[i*3+0], (LPARAM)0 );
			}
			if( (key2 & 0xF000) == CCfgShortCut::K_ALT && keyex == CCfgShortCut::K_ALT && KeyBuf[key2&0x0FFF] && (key2&0x0FFF) ) {
				::PostMessage( m_hWnd, WM_COMMAND, (WPARAM)pShortCutKeyID[i*3+0], (LPARAM)0 );
			}
		}
	} else {
		INT	no;
		WORD	key, key2;
		for( INT i = 0; pShortCutKeyID[i*3+0]; i++ ) {
			no = pShortCutKeyID[i*3+2];
			key = pShortCutKey[no];
			key2 = pShortCutKey[no+128];
			if( (key & 0xF000) == keyex && KeyBuf[key&0x0FFF] && (key&0x0FFF) ) {
				::PostMessage( m_hWnd, WM_COMMAND, (WPARAM)pShortCutKeyID[i*3+0], (LPARAM)0 );
			}
			if( (key2 & 0xF000) == keyex && KeyBuf[key2&0x0FFF] && (key2&0x0FFF) ) {
				::PostMessage( m_hWnd, WM_COMMAND, (WPARAM)pShortCutKeyID[i*3+0], (LPARAM)0 );
			}
		}
	}

	m_bKeyChecking = FALSE;
}


WNDCMD	CMainFrame::OnSelFont(WNDCMDPARAM)
{
	//LOGFONT lf;
	CHOOSEFONT cf;
	memset((void *)&cf, 0, sizeof(cf));
	cf.lStructSize = sizeof(CHOOSEFONT);
	cf.hwndOwner = m_hWnd;
	cf.lpLogFont = &DirectDraw.m_logFont;
	cf.iPointSize = 160;
	cf.Flags = CF_EFFECTS | CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT | CF_LIMITSIZE ;
	cf.rgbColors = RGB(0, 0, 0);
	cf.nFontType = REGULAR_FONTTYPE;
	cf.nSizeMax = 9;
	cf.nSizeMin = 9;

	if(ChooseFont(&cf) == 0)
	{
		memset((void *)&DirectDraw.m_logFont, 0, sizeof(DirectDraw.m_logFont));
		DirectDraw.m_logFont.lfHeight = 12;
		DirectDraw.m_logFont.lfWidth = 6;
		DirectDraw.m_logFont.lfCharSet = DEFAULT_CHARSET;
		DirectDraw.m_logFont.lfOutPrecision = OUT_TT_PRECIS;
		DirectDraw.m_logFont.lfClipPrecision = CLIP_DEFAULT_PRECIS;
		DirectDraw.m_logFont.lfQuality = DEFAULT_QUALITY;
		DirectDraw.m_logFont.lfPitchAndFamily = FIXED_PITCH|FF_DONTCARE;
		strcpy(DirectDraw.m_logFont.lfFaceName, "宋体");	
	}
	else
	{
		DirectDraw.m_logFont.lfHeight = 12;
		DirectDraw.m_logFont.lfWidth = 6;
	}

}


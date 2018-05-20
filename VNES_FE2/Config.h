//
// 設定保存クラス
//
#ifndef	__CCONFIG_INCLUDED__
#define	__CCONFIG_INCLUDED__

#define	WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdlib.h>
#include <dinput.h>

#include <string>
using namespace std;

#include "typedef.h"
#include "macro.h"

class	CCfgGeneral
{
public:
	BOOL	bOnTop;

	BOOL	bDoubleExecute;

	BOOL	bWindowSave;
	BOOL	bWindowZoom;
	RECT	rcWindowPos;

	BOOL	bScreenMode;
	BOOL	nScreenZoom;

	RECT	rcSearchDlgPos;

	//RECT	rcPatternViewPos;
	//RECT	rcNameTableViewPos;
	//RECT	rcPaletteViewPos;
	//RECT	rcMemoryViewPos;

	LOGFONT	stLogFont;

	void	Default() {
		bOnTop = TRUE;
		bWindowSave = FALSE;
		bWindowZoom = FALSE;
		rcWindowPos.left = rcWindowPos.right =
		rcWindowPos.top = rcWindowPos.bottom = 0;

		bDoubleExecute = TRUE;

		bScreenMode = FALSE;
		nScreenZoom = 0;

		rcSearchDlgPos.left = rcSearchDlgPos.right =
		rcSearchDlgPos.top = rcSearchDlgPos.bottom = 0;

//		rcPatternViewPos.left = rcPatternViewPos.right =
//		rcPatternViewPos.top = rcPatternViewPos.bottom = 0;
//		rcNameTableViewPos.left = rcNameTableViewPos.right =
//		rcNameTableViewPos.top = rcNameTableViewPos.bottom = 0;
//		rcPaletteViewPos.left = rcPaletteViewPos.right =
//		rcPaletteViewPos.top = rcPaletteViewPos.bottom = 0;
//		rcMemoryViewPos.left = rcMemoryViewPos.right =
//		rcMemoryViewPos.top = rcMemoryViewPos.bottom = 0;
	}
};

class	CCfgPath
{
public:
	BOOL	bRomPath;
	BOOL	bSavePath;
	BOOL	bStatePath;
	BOOL	bSnapshotPath;
	BOOL	bMoviePath;
	BOOL	bWavePath;
	BOOL	bCheatPath;

	CHAR	szRomPath     [_MAX_PATH];
	CHAR	szSavePath    [_MAX_PATH];
	CHAR	szStatePath   [_MAX_PATH];
	CHAR	szSnapshotPath[_MAX_PATH];
	CHAR	szMoviePath   [_MAX_PATH];
	CHAR	szWavePath    [_MAX_PATH];
	CHAR	szCheatPath   [_MAX_PATH];

	void	Default() {
		bRomPath = bSavePath =
		bStatePath = bSnapshotPath =
		bMoviePath = bWavePath =
		bCheatPath = FALSE;
		szRomPath[0] = szSavePath[0] =
		szStatePath[0] = szSnapshotPath[0] =
		szMoviePath[0] = szWavePath[0] =
		szCheatPath[0] = 0;
	}
};

class	CCfgEmulator
{
public:
	BOOL	bIllegalOp;
	BOOL	bAutoFrameSkip;
	BOOL	bThrottle;
	INT	nThrottleFPS;
	BOOL	bBackground;
	INT	nPriority;
	BOOL	bFourPlayer;
	BOOL	bCrcCheck;
	BOOL	bDiskThrottle;
	BOOL	bLoadFullscreen;

	void	Default() {
		bIllegalOp	= FALSE;
		bAutoFrameSkip	= TRUE;
		bThrottle	= TRUE;
		nThrottleFPS	= 120;		// 120FPS
		bBackground	= FALSE;
		nPriority	= 3;		// Normal
		bFourPlayer	= TRUE;		// TRUE:NES FALSE:Famicom
		bCrcCheck	= TRUE;
		bDiskThrottle	= TRUE;
		bLoadFullscreen	= FALSE;
	}
};

class	CCfgGraphics
{
public:
	BOOL	bAspect;
	BOOL	bAllSprite;
	BOOL	bAllLine;
	BOOL	bFPSDisp;
	BOOL	bTVFrame;
	BOOL	bScanline;
	INT	nScanlineColor;
	BOOL	bSyncDraw;
	BOOL	bFitZoom;

	BOOL	bLeftClip;

	BOOL	bWindowVSync;

	BOOL	bDoubleSize;
	BOOL	bSystemMemory;
	BOOL	bUseHEL;

	DWORD	dwDisplayWidth;
	DWORD	dwDisplayHeight;
	DWORD	dwDisplayDepth;
	DWORD	dwDisplayRate;

	BOOL	bPaletteFile;
	CHAR	szPaletteFile[_MAX_PATH];

	void	Default() {
		bAspect		= FALSE;
		bAllSprite	= TRUE;
		bAllLine	= FALSE;
		bFPSDisp	= FALSE;
		bTVFrame	= FALSE;
		bScanline	= FALSE;
		nScanlineColor	= 75;
		bSyncDraw	= FALSE;
		bFitZoom	= FALSE;

		bLeftClip	= TRUE;

		bWindowVSync	= FALSE;

		bDoubleSize	= FALSE;
		bSystemMemory	= FALSE;
		bUseHEL		= FALSE;

		dwDisplayWidth	= 640;
		dwDisplayHeight	= 480;
		dwDisplayDepth	= 16;
		dwDisplayRate	= 0;

		bPaletteFile = FALSE;
		szPaletteFile[0] = 0;
	}
};

class	CCfgSound
{
public:
	BOOL	bEnable;
	INT	nRate;
	INT	nBits;
	INT	nBufferSize;

	INT	nFilterType;

	//  0:Master
	//  1:Rectangle 1
	//  2:Rectangle 2
	//  3:Triangle
	//  4:Noise
	//  5:DPCM
	//  6:VRC6
	//  7:VRC7
	//  8:FDS
	//  9:MMC5
	// 10:N106
	// 11:FME7
	SHORT	nVolume[16];

	void	Default() {
		bEnable		= TRUE;
		nRate		= 22050;
		nBits		= 8;
		nBufferSize	= 4;

		nFilterType	= 0;

		for( INT i = 0; i < 16; i++ ) {
			nVolume[i] = 100;
		}
	}
};

class	CCfgShortCut
{
public:
	WORD	nShortCut[256];

	enum {
		K_ALT   = 0x8000,
		K_CTRL  = 0x4000,
		K_SHIFT = 0x2000,
	};

	void	Default() {
		for( INT i = 0; i < 256; i++ ) {
			nShortCut[i] = 0;
		}

		// Main controls
		nShortCut[  0] = 0;		// ID_OPEN
		nShortCut[  1] = 0;		// ID_CLOSE
		nShortCut[  2] = 0;		// ID_LAUNCHER
		nShortCut[  3] = 0;		// ID_NETPLAY_CONNECT
		nShortCut[  4] = 0;		// ID_NETPLAY_DISCONNECT
		nShortCut[  8] = 0;		// ID_ROMINFO
		nShortCut[  9] = 0;		// ID_WAVERECORD
		nShortCut[ 15] = 0;		// ID_EXIT

		// Emulation controls
		nShortCut[ 16] = DIK_F1;		// Hardware reset
		nShortCut[ 17] = DIK_F2;		// Software reset
		nShortCut[ 18] = DIK_F4;		// Hardware pause
		nShortCut[ 19] = 0;		// Throttle(toggled)
		nShortCut[ 20] = 0;		// Frame skip Auto
		nShortCut[ 21] = 0;		// Frame skip +
		nShortCut[ 22] = 0;		// Frame skip -

		nShortCut[ 23] = 0;		// One Frame step
		nShortCut[ 24] = 0;		// Throttle(Not toggle)

		// State controls
		nShortCut[ 32] = DIK_F7;			// State Load
		nShortCut[ 33] = DIK_F5;			// State Save
		nShortCut[ 34] = DIK_F9;		// State Slot +
		nShortCut[ 35] = DIK_F10;		// State Slot -

		// Disk controls
		nShortCut[ 48] = 0;			// Disk Eject
		nShortCut[ 49] = 0;			// Disk 0 Side A
		nShortCut[ 50] = 0;			// Disk 0 Side B
		nShortCut[ 51] = 0;			// Disk 1 Side A
		nShortCut[ 52] = 0;			// Disk 1 Side B

		// Movie controls
		nShortCut[ 56] = 0;		// Movie Play
		nShortCut[ 57] = 0;		// Movie Rec
		nShortCut[ 58] = 0;		// Movie Rec Append
		nShortCut[ 59] = 0;		// Movie Stop
		nShortCut[ 60] = 0;		// Movie Info

		// Screen controls
		nShortCut[ 64] = DIK_F12;		// Zoom x1
		nShortCut[ 65] = 0;		// Zoom x2
		nShortCut[ 66] = 0;		// Zoom x3
		nShortCut[ 67] = 0;		// Zoom x4
		nShortCut[ 68] = 0;		// Fullscreen

		// Sound controls
		nShortCut[ 72] = 0;	// Mute Master
		nShortCut[ 73] = 0;		// Mute Rectangle #1
		nShortCut[ 74] = 0;		// Mute Rectangle #2
		nShortCut[ 75] = 0;		// Mute Triangle
		nShortCut[ 76] = 0;		// Mute Noise
		nShortCut[ 77] = 0;		// Mute Dpcm
		nShortCut[ 78] = 0;		// Mute External #1
		nShortCut[ 79] = 0;		// Mute External #2
		nShortCut[ 80] = 0;		// Mute External #3
		nShortCut[ 81] = 0;		// Mute External #4
		nShortCut[ 82] = 0;		// Mute External #5
		nShortCut[ 83] = 0;	// Mute External #6
		nShortCut[ 84] = 0;	// Mute External #7
		nShortCut[ 85] = 0;	// Mute External #8

		// Other controls
		nShortCut[ 96] = DIK_P+K_CTRL;		// Snapshot
		nShortCut[ 97] = 0;		// FPSDISP
		nShortCut[ 98] = 0;	// TV Aspect
		nShortCut[ 99] = 0;	// TV frame
		nShortCut[100] = 0;		// Scanline
		nShortCut[101] = 0;		// Show 240 lines
		nShortCut[102] = 0;		// All sprites
		nShortCut[103] = 0;		// Sync draw
		nShortCut[104] = 0;	// Fit screen
		nShortCut[105] = 0;			// Left clip

		// Cheat
		nShortCut[110] = 0;	// Search
		nShortCut[111] = 0;	// Cheat
		nShortCut[112] = 0;		// Cheat enable
		nShortCut[113] = 0;		// Cheat disable
		nShortCut[114] = 0;			// GameGenie

		// Tools
		nShortCut[116] = 0;			// Pattern Viewer
		nShortCut[117] = 0;			// NameTable Viewer
		nShortCut[118] = 0;			// Palette Viewer

		nShortCut[119] = 0;			// Memory Viewer
	}
};

class	CCfgLanguage
{
public:
	LCID	lcID;
	CHAR	szLanguage[256];

	void	Default() {
		lcID = NULL;
		szLanguage[0] = 0;
	}
};

class	CCfgController
{
public:

	WORD	nButton[4][32];
	WORD	nRapid [4][ 2];
	WORD	nExButton[1][32];
	WORD	nNsfButton[32];

	void	Default() {
		DefaultController1();
		DefaultController2();
		DefaultController3();
		DefaultController4();
		DefaultExController0();
		DefaultNsfController();
	}

	void	DefaultController1() {
		nButton[0][ 0] = DIK_UP;
		nButton[0][ 1] = DIK_DOWN;
		nButton[0][ 2] = DIK_LEFT;
		nButton[0][ 3] = DIK_RIGHT;
		nButton[0][ 4] = DIK_X;	// A
		nButton[0][ 5] = DIK_Z;	// B
		nButton[0][ 6] = 0;	// A Rapid
		nButton[0][ 7] = 0;	// B Rapid
		nButton[0][ 8] = DIK_RSHIFT; // SELECT
		nButton[0][ 9] = DIK_RETURN; // START
		nButton[0][10] = 0;	// EX C
		nButton[0][11] = 0;	// EX D
		nButton[0][12] = 0;	// EX E
		nButton[0][13] = 0;	// EX F
		nButton[0][14] = 0;	// EX G
		nButton[0][15] = 0;	// EX H

		for( INT i = 16; i < 32; i++ )
			nButton[0][i] = 0;	// 2nd keys

		nRapid[0][0] = 0;	// A Rapid speed
		nRapid[0][1] = 0;	// B Rapid speed
	}

	void	DefaultController2() {
		nButton[1][ 0] = DIK_NUMPAD8;
		nButton[1][ 1] = DIK_NUMPAD2;
		nButton[1][ 2] = DIK_NUMPAD4;
		nButton[1][ 3] = DIK_NUMPAD6;
		nButton[1][ 4] = DIK_N;	// A
		nButton[1][ 5] = DIK_B;	// B
		nButton[1][ 6] = 0;	// A Rapid
		nButton[1][ 7] = 0;	// B Rapid
		nButton[1][ 8] = 0;	// SELECT
		nButton[1][ 9] = 0;	// START
		nButton[1][10] = DIK_M;	// Mic
		nButton[1][11] = 0;	// EX D
		nButton[1][12] = 0;	// EX E
		nButton[1][13] = 0;	// EX F
		nButton[1][14] = 0;	// EX G
		nButton[1][15] = 0;	// EX H

		for( INT i = 16; i < 32; i++ )
			nButton[1][i] = 0;	// 2nd keys

		nRapid[1][0] = 0;	// A Rapid speed
		nRapid[1][1] = 0;	// B Rapid speed
	}

	void	DefaultController3() {
		nButton[2][ 0] = 0;
		nButton[2][ 1] = 0;
		nButton[2][ 2] = 0;
		nButton[2][ 3] = 0;
		nButton[2][ 4] = 0;	// A
		nButton[2][ 5] = 0;	// B
		nButton[2][ 6] = 0;	// A Rapid
		nButton[2][ 7] = 0;	// B Rapid
		nButton[2][ 8] = 0;	// SELECT
		nButton[2][ 9] = 0;	// START
		nButton[2][10] = 0;	// EX C
		nButton[2][11] = 0;	// EX D
		nButton[2][12] = 0;	// EX E
		nButton[2][13] = 0;	// EX F
		nButton[2][14] = 0;	// EX G
		nButton[2][15] = 0;	// EX H

		for( INT i = 16; i < 32; i++ )
			nButton[2][i] = 0;	// 2nd keys

		nRapid[2][0] = 0;	// A Rapid speed
		nRapid[2][1] = 0;	// B Rapid speed
	}

	void	DefaultController4() {
		nButton[3][ 0] = 0;
		nButton[3][ 1] = 0;
		nButton[3][ 2] = 0;
		nButton[3][ 3] = 0;
		nButton[3][ 4] = 0;	// A
		nButton[3][ 5] = 0;	// B
		nButton[3][ 6] = 0;	// A Rapid
		nButton[3][ 7] = 0;	// B Rapid
		nButton[3][ 8] = 0;	// SELECT
		nButton[3][ 9] = 0;	// START
		nButton[3][10] = 0;	// EX C
		nButton[3][11] = 0;	// EX D
		nButton[3][12] = 0;	// EX E
		nButton[3][13] = 0;	// EX F
		nButton[3][14] = 0;	// EX G
		nButton[3][15] = 0;	// EX H

		for( INT i = 16; i < 32; i++ )
			nButton[3][i] = 0;	// 2nd keys

		nRapid[3][0] = 0;	// A Rapid speed
		nRapid[3][1] = 0;	// B Rapid speed
	}

	void	DefaultExController0() {
		// Crazy Climberｺﾝﾄﾛｰﾗ(実際には存在しない)
		nExButton[0][ 0] = 0;	// L up
		nExButton[0][ 1] = 0;	// L down
		nExButton[0][ 2] = 0;	// L left
		nExButton[0][ 3] = 0;	// L right
		nExButton[0][ 4] = 0;	// R up
		nExButton[0][ 5] = 0;	// R down
		nExButton[0][ 6] = 0;	// R left
		nExButton[0][ 7] = 0;	// R right
		nExButton[0][ 8] = 0;	// No use
		nExButton[0][ 9] = 0;	// No use
		nExButton[0][10] = 0;	// No use
		nExButton[0][11] = 0;	// No use
		nExButton[0][12] = 0;	// No use
		nExButton[0][13] = 0;	// No use
		nExButton[0][14] = 0;	// No use
		nExButton[0][15] = 0;	// No use

		for( INT i = 16; i < 32; i++ )
			nExButton[0][i] = 0;	// 2nd keys
	}

	void	DefaultNsfController() {
		nNsfButton[ 0] = DIK_UP;	// Play
		nNsfButton[ 1] = DIK_DOWN;	// Stop
		nNsfButton[ 2] = DIK_LEFT;	// Number -1
		nNsfButton[ 3] = DIK_RIGHT;	// Number +1
		nNsfButton[ 4] = DIK_PRIOR;	// Number +16
		nNsfButton[ 5] = DIK_NEXT;	// Number -16
		nNsfButton[ 6] = 0;		// No use
		nNsfButton[ 7] = 0;		// No use
		nNsfButton[ 8] = 0;		// No use
		nNsfButton[ 9] = 0;		// No use
		nNsfButton[10] = 0;		// No use
		nNsfButton[11] = 0;		// No use
		nNsfButton[12] = 0;		// No use
		nNsfButton[13] = 0;		// No use
		nNsfButton[14] = 0;		// No use
		nNsfButton[15] = 0;		// No use

		for( INT i = 16; i < 32; i++ )
			nNsfButton[i] = 0;	// 2nd keys
	}
};

class	CCfgMovie
{
public:
	BYTE	bUsePlayer[4];
	BOOL	bRerecord;
	BOOL	bLoopPlay;
	BOOL	bResetRec;
	BOOL	bPadDisplay;

	void	Default() {
		bUsePlayer[0] = 0xFF;
		bUsePlayer[1] = 0x00;
		bUsePlayer[2] = 0x00;
		bUsePlayer[3] = 0x00;
		bRerecord = TRUE;
		bLoopPlay = FALSE;
		bResetRec = FALSE;
		bPadDisplay = FALSE;
	}
};

class	CCfgLauncher
{
public:
	RECT	rcWindowPos;

	BYTE	bHeaderView [16];
	BYTE	nHeaderOrder[16];
	SHORT	nHeaderWidth[16];

	BYTE	bFolderUse[16];
	CHAR	szFolder[16][_MAX_PATH];

	CHAR	szLastSelect[_MAX_PATH];

	void	Default() {
		rcWindowPos.left = rcWindowPos.right =
		rcWindowPos.top = rcWindowPos.bottom = 0;

		INT	i;
		for( i = 0; i < 16; i++ ) {
			bHeaderView [i] = TRUE;
			nHeaderOrder[i] = i;
			nHeaderWidth[i] = (i==0)?160:48;
		}

		for( i = 0; i < 16; i++ ) {
			bFolderUse[i] = FALSE;
			szFolder[i][0] = '\0';
		}

		szLastSelect[0] = '\0';
	}
};

class	CConfig
{
public:
	CConfig() { Default(); }

	CCfgGeneral	general;
	CCfgPath	path;
	CCfgEmulator	emulator;
	CCfgGraphics	graphics;
	CCfgSound	sound;
	CCfgShortCut	shortcut;
	CCfgLanguage	language;
	CCfgController	controller;
	CCfgMovie	movie;
	CCfgLauncher	launcher;

	void	Default() {
		general.Default();
		path.Default();
		emulator.Default();
		graphics.Default();
		sound.Default();
		shortcut.Default();
		language.Default();
		controller.Default();
		movie.Default();
		launcher.Default();
	}

	void	Load();
	void	Save();

// Converter
	string	ShortcutToKeyName( INT nShortcut );

// Table
	static	INT	SamplingRateTable[];
	static	INT	ShortcutKeyID[];
protected:
private:
};

// 個別ゲームオプション
class	CGameOption
{
public:
	// for Cartridge
	void	Load( DWORD crc );
	void	Save( LPCSTR name, DWORD crc );
	// for FDS
	void	Load( DWORD gid, DWORD mid );
	void	Save( LPCSTR name, DWORD gid, DWORD mid );

	// Default保存
	INT	defRenderMethod;
	INT	defIRQtype;
	BOOL	defFrameIRQ;
	BOOL	defVideoMode;

	// データ
	INT	nRenderMethod;
	INT	nIRQtype;
	BOOL	bFrameIRQ;
	BOOL	bVideoMode;
protected:
private:
};

extern	CConfig		Config;
extern	CGameOption	GameOption;

#endif // !__CCONFIG_INCLUDED__

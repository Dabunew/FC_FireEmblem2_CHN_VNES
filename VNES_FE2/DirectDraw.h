//
// DirectDraw class
//
#ifndef	__DIRECTDRAW_INCLUDED__
#define	__DIRECTDRAW_INCLUDED__

#define	WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <mmsystem.h>
#include <ddraw.h>
#include <stdio.h>
#include <stdlib.h>

#include "typedef.h"
#include "macro.h"

class	CDirectDraw
{
public:

	LOGFONT	m_logFont;

	// パブリックメンバ変数
	typedef	struct	tagDISPLAYMODE {
		DWORD		dwWidth;
		DWORD		dwHeight;
		DWORD		dwDepth;
		DWORD		dwRate;
	} DISPLAYMODE, *LPDISPLAYMODE;

	typedef	struct	tagPALBUF {
		BYTE		r;
		BYTE		g;
		BYTE		b;
	} PALBUF, *LPPALBUF;

	// 利用可能なディスプレイモード
	enum { DD_DISPLAYMODEMAX = 256 };

	INT		m_DisplayModeNum;	// ディスプレイモード数
	DISPLAYMODE	m_DisplayModeBuf[DD_DISPLAYMODEMAX];

	// フルスクリーンディスプレイモード
	DWORD	m_dwDisplayWidth;
	DWORD	m_dwDisplayHeight;
	DWORD	m_dwDisplayDepth;
	DWORD	m_dwDisplayRate;

	// Screen size
	enum { SCREEN_WIDTH = 256, SCREEN_HEIGHT = 240 };

	// Render screen size
	enum { RENDER_WIDTH = SCREEN_WIDTH+16, RENDER_HEIGHT = SCREEN_HEIGHT };

	// パブリックメンバ関数
	CDirectDraw();
	virtual ~CDirectDraw();

	BOOL	InitialDDraw( HWND hWnd );
	void	ReleaseDDraw();

	BOOL	InitialSurface( BOOL bScreenMode );
	BOOL	ReleaseSurface();
	BOOL	RestoreSurface();

	BOOL	BeginDisplayChange();
	BOOL	EndDisplayChange();
	BOOL	OnChangeDisplayMode();

	BOOL	CalcPaletteTable();
	void	SetPaletteTable( LPBYTE pal );
	void	SetPaletteFile( LPCTSTR fname );

	BOOL	SetFullScreenGDI( BOOL bMode );

	void	ViewChars(LPDIRECTDRAWSURFACE7 lpDDSurface = NULL);
	void	ViewChars(unsigned char * lpRdr = NULL);

	void	SetUseHEL( BOOL bUseHEL ) {
		m_bUseHEL = bUseHEL;
	}
	void	SetSystemMemory( BOOL bSystemMemory ) {
		m_bSystemMemory = bSystemMemory;
	}
	void	SetScreenMode( BOOL bScreenMode ) {
		m_bScreenMode = bScreenMode;
	}
	void	SetFlipMode( BOOL bFlip ) {
		m_bFlip = bFlip;
	}
	void	SetAspectMode( BOOL bAspect ) {
		m_bAspect = bAspect;
	}
	void	SetAllLineMode( BOOL bAllLine ) {
		m_bAllLine = bAllLine;
	}
	void	SetMaxZoom( BOOL bZoom ) {
		m_bMaxZoom = bZoom;
	}
	void	SetDoubleSize( BOOL bDoubleSize ) {
		m_bDoubleSize = bDoubleSize;
	}
	void	SetTVFrameMode( BOOL bTVFrame ) {
		m_bTVFrameMode = bTVFrame;
	}
	void	SetScanlineMode( BOOL bScanline ) {
		m_bScanlineMode = bScanline;
	}
	void	SetScanlineColor( INT nScanlineColor ) {
		m_nScanlineColor = nScanlineColor;
	}

	void	SetZapperMode( BOOL bZapper ) {
		m_bZapper = bZapper;
	}
	void	SetZapperDrawMode( BOOL bDraw ) {
		m_bZapperDraw = bDraw;
	}

	void	SetWindowVSyncMode( BOOL bVSync ) {
		m_bWindowVSync = bVSync;
	}

	BOOL	GetUseHEL()		{ return m_bUseHEL; }
	BOOL	GetSystemMemory()	{ return m_bSystemMemory; }
	BOOL	GetScreenMode()		{ return m_bScreenMode; }
	BOOL	GetFlipMode()		{ return m_bFlip; }
	BOOL	GetAspectMode()		{ return m_bAspect; }
	BOOL	GetAllLineMode()	{ return m_bAllLine; }
	BOOL	GetMaxZoom()		{ return m_bMaxZoom; }
	BOOL	GetDoubleSize()		{ return m_bDoubleSize; }
	BOOL	GetTVFrameMode()	{ return m_bTVFrameMode; }
	BOOL	GetScanlineMode()	{ return m_bScanlineMode; }
	INT	GetScanlineColor()	{ return m_nScanlineColor; }
	BOOL	GetZapperMode()		{ return m_bZapper; }
	BOOL	GetZapperDrawMode()	{ return m_bZapperDraw; }

	BOOL	GetWindowVSyncMode()	{ return m_bWindowVSync; }

	LPBYTE	GetRenderScreen()	{ return m_lpRender; }

	// For Zapper
	BYTE	GetZapperHit();
	void	GetZapperPos( LONG& x, LONG& y );
	void	SetZapperPos( LONG x, LONG y );

	void	SetInfoString( LPCSTR str );
	void	SetMessageString( LPCSTR str );

	void	RealizePalette();

	void	OnScreenDraw();
	void	SetPaletteMode( INT nMode, BOOL bMono );
	void	GetPaletteData( RGBQUAD* rgb );

	void	Blt();
	void	Flip();

protected:
	// プロテクトメンバ変数

	// General
	HWND	m_hWnd;			// Window handle
	BOOL	m_bChangeMode;		// Screen Mode changing flag
	BOOL	m_bDraw;		// Drawing flag
	BOOL	m_bNoRestore;

	BOOL	m_bMessage;		// Message display?
	DWORD	m_dwMessageTime;	// Message display times

	// Options
	BOOL	m_bUseHEL;		// HALを使わずHELを使用
	BOOL	m_bSystemMemory;	// サーフェスにシステムメモリを使用

	BOOL	m_bScreenMode;		// FALSE:Window  TRUE:Fullscreen
	BOOL	m_bFlip;		// フリップ使用
	BOOL	m_bGDI;			// GDI
	BOOL	m_bAspect;		// TVアスペクト補正
	BOOL	m_bAllLine;		// 全ライン表示
	BOOL	m_bMaxZoom;		// 画面全体まで拡大
	BOOL	m_bDoubleSize;		// ２倍サイズレンダリング
	BOOL	m_bTVFrameMode;		// TV枠表示
	BOOL	m_bScanlineMode;	// スキャンラインモード
	INT		m_nScanlineColor;	// スキャンラインカラー
	BOOL	m_bZapper;		// Zapper
	BOOL	m_bZapperDraw;		// Zapper Sight drawing

	BOOL	m_bWindowVSync;		// Wait for VSync(Window mode)

	// DirectDraw object
	LPDIRECTDRAW7		m_lpDD;		// DirectDraw Object
	
	
	LPDIRECTDRAWSURFACE7	m_lpDDPrimary;	// PrimarySurface
	LPDIRECTDRAWSURFACE7	m_lpDDBack;	// BackSurface
	LPDIRECTDRAWSURFACE7	m_lpDDRender;	// RenderSurface
	LPDIRECTDRAWSURFACE7	m_lpDDAscii;	// AsciiSurface
	LPDIRECTDRAWSURFACE7	m_lpDDZapper;	// ZapperSurface
	LPDIRECTDRAWSURFACE7	m_lpDDTV;	// TV枠(^^;
	LPDIRECTDRAWSURFACE7	m_lpDDChar;	// ｵ･ｸﾐﾎﾄﾎﾄﾗﾖ;
	LPDIRECTDRAWCLIPPER	m_lpDDClipper;	// Clipper
	LPDIRECTDRAWCLIPPER	m_lpDDClipper2;	// Render window Clipper
	LPDIRECTDRAWPALETTE	m_lpDDPalette;	// DirectDraw Palette

	LPBYTE			m_lpRender;

	// パレット
	INT	m_nPaletteMode;
	BOOL	m_bMonoMode;
	BOOL	m_bPaletteUpdate;	// Palette Update?

	static	PALBUF	m_PalDefault[64];	// デフォルトパレット
	PALBUF		m_PaletteBuf[64];	// パレットバッファ

	typedef	struct	tagLOGPAL {
		WORD		palVersion;
		WORD		palNumEntries;
		PALETTEENTRY	pe[256];
	} LOGPAL, *LPLOGPAL;

	HPALETTE	m_hPalette;	// Windows GDI Palette handle
	LOGPAL		m_logPalette;	// Windows GDI Palette

	// 256色モード用
	RGBQUAD	m_cpPalette[8][64*2];	// Color
	RGBQUAD	m_mpPalette[8][64*2];	// Monochrome

	// ピクセルフォーマットに変換したパレット
	DWORD	m_cnPalette[8][256];	// Color
	DWORD	m_csPalette[8][256];	// Color/Scanline
	DWORD	m_mnPalette[8][256];	// Monochrome
	DWORD	m_msPalette[8][256];	// Monochrome/Scanline

	// For Zapper
	LONG	m_ZapperPosX, m_ZapperPosY;

	// For Infomation&Message
	CHAR	m_szInfo[32+1];
	CHAR	m_szMess[32+1];

	// プロテクトメンバ関数
	static	HRESULT	WINAPI	EnumModesCallback( LPDDSURFACEDESC2 lpDDSurfaceDesc, LPVOID lpContext );

	void	RenderString( INT x, INT y, LPCSTR str );

	BOOL	RestoreSurfaceSub( LPDIRECTDRAWSURFACE7 lpSurface );
	void	GetBitMask( DWORD val, int& shift, int& bits );

	void	SetLZSSChar( LPBYTE lpLZ, LPDIRECTDRAWSURFACE7 lpDDSurface );

	void	Render8Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd );
	void	Render16Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd, LPBYTE lpPal, LPBYTE lpPalScan );
	void	Render24Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd, LPBYTE lpPal, LPBYTE lpPalScan );
	void	Render32Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd, LPBYTE lpPal, LPBYTE lpPalScan );

	void	LineRender8_Normal( LPBYTE lpRdr, LPBYTE lpScn, DWORD dwMask, LONG width );
	void	LineRender8_Double( LPBYTE lpRdr, LPBYTE lpScn, DWORD dwMask, LONG width );
	void	LineRender16_Normal( LPBYTE lpRdr, LPWORD lpScn, LPDWORD lpPal, LONG width );
	void	LineRender16_Double( LPBYTE lpRdr, LPWORD lpScn, LPDWORD lpPal, LONG width );
	void	LineRender24_Normal( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width );
	void	LineRender24_Double( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width );
	void	LineRender32_Normal( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width );
	void	LineRender32_Double( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width );

private:
	// プライベートメンバ変数
	// プライベートメンバ関数
};

extern	CDirectDraw	DirectDraw;

#endif	// !__DIRECTDRAW_INCLUDED__

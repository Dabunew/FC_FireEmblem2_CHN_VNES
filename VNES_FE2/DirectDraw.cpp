//
// DirectDraw class
//
#include "DebugOut.h"
#include "DirectDraw.h"
#include "COM.h"

#include "MMU.h"

#include "Chars.h"

CDirectDraw	DirectDraw;

//////////////////////////////////////////////////////////////////////
// デフォルトパレット
//////////////////////////////////////////////////////////////////////
CDirectDraw::PALBUF	CDirectDraw::m_PalDefault[] = {
	0x7F, 0x7F, 0x7F, 0x20, 0x00, 0xB0, 0x28, 0x00, 0xB8, 0x60, 0x10, 0xA0, 0x98, 0x20, 0x78, 0xB0,
	0x10, 0x30, 0xA0, 0x30, 0x00, 0x78, 0x40, 0x00, 0x48, 0x58, 0x00, 0x38, 0x68, 0x00, 0x38, 0x6C,
	0x00, 0x30, 0x60, 0x40, 0x30, 0x50, 0x80, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,

	0xBC, 0xBC, 0xBC, 0x40, 0x60, 0xF8, 0x40, 0x40, 0xFF, 0x90, 0x40, 0xF0, 0xD8, 0x40, 0xC0, 0xD8,
	0x40, 0x60, 0xE0, 0x50, 0x00, 0xC0, 0x70, 0x00, 0x88, 0x88, 0x00, 0x50, 0xA0, 0x00, 0x48, 0xA8,
	0x10, 0x48, 0xA0, 0x68, 0x40, 0x90, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,

	0xFF, 0xFF, 0xFF, 0x60, 0xA0, 0xFF, 0x50, 0x80, 0xFF, 0xA0, 0x70, 0xFF, 0xF0, 0x60, 0xFF, 0xFF,
	0x60, 0xB0, 0xFF, 0x78, 0x30, 0xFF, 0xA0, 0x00, 0xE8, 0xD0, 0x20, 0x98, 0xE8, 0x00, 0x70, 0xF0,
	0x40, 0x70, 0xE0, 0x90, 0x60, 0xD0, 0xE0, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,

	0xFF, 0xFF, 0xFF, 0x90, 0xD0, 0xFF, 0xA0, 0xB8, 0xFF, 0xC0, 0xB0, 0xFF, 0xE0, 0xB0, 0xFF, 0xFF,
	0xB8, 0xE8, 0xFF, 0xC8, 0xB8, 0xFF, 0xD8, 0xA0, 0xFF, 0xF0, 0x90, 0xC8, 0xF0, 0x80, 0xA0, 0xF0,
	0xA0, 0xA0, 0xFF, 0xC8, 0xA0, 0xFF, 0xF0, 0xA0, 0xA0, 0xA0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
};

//////////////////////////////////////////////////////////////////////
// キャラクタ
//////////////////////////////////////////////////////////////////////
static	BYTE	lzSight[] = {
#include "lzSight.h"
};
static	BYTE	lzAscii[] = {
#include "lzAscii.h"
};
static	BYTE	lzTVlayer[] = {
#include "lzTVlayer.h"
};

// メモリから１バイト取ってくる
static __inline INT	mgetc( LPBYTE& inbuf, INT& bufsize )
{
	if( bufsize <= 0 )
		return	-1;
	bufsize--;
	return	(INT)*inbuf++;
}

//
// LZSS解凍ルーチン
//
#define	N	4096
#define	F	18

static	void	LZdecode( BYTE *inbuf, BYTE *outbuf, INT length )
{
INT	i, j, k, r, c;
WORD	flags;
BYTE	text[N+F-1];
INT	outsize = 0;

	for( i = 0; i < N - F; i++ )
		text[i] = 0;
	r = N - F;
	flags = 0;

	for(;;) {
		if( ((flags >>= 1) & 256) == 0 ) {
			if( (c = mgetc( inbuf, length )) < 0 )
				break;
			flags = c | 0xff00;
		}
		if( flags & 1 ) {
			if( (c = mgetc( inbuf, length )) < 0 )
				break;
			*outbuf++ = c;
			text[r++] = c;
			r &= (N-1);

			outsize++;
		} else {
			if( (i = mgetc( inbuf, length )) < 0 )
				break;
			if( (j = mgetc( inbuf, length )) < 0 )
				break;
			i |= ((j & 0xf0)<<4);
			j  =  (j & 0x0f)+2;
			for( k = 0; k <= j; k++ ) {
				c = text[(i+k) & (N-1)];
				*outbuf++ = c;
				text[r++] = c;
				r &= (N-1);

				outsize++;
			}
		}
	}
}

//////////////////////////////////////////////////////////////////////
// 構築/消滅
//////////////////////////////////////////////////////////////////////

CDirectDraw::CDirectDraw()
{
	// General
	m_hWnd           = NULL;

	m_bUseHEL        = FALSE;
	m_bSystemMemory  = FALSE;

	m_bScreenMode    = FALSE;
	m_bGDI           = FALSE;
	m_bAspect        = FALSE;
	m_bAllLine       = FALSE;
	m_bMaxZoom       = FALSE;
	m_bDoubleSize    = FALSE;
	m_bTVFrameMode   = FALSE;
	m_bScanlineMode  = FALSE;
	m_nScanlineColor = 75;
	m_bZapper        = FALSE;
	m_bZapperDraw    = FALSE;

	m_ZapperPosX = m_ZapperPosY = -1;

	m_bWindowVSync   = FALSE;

	m_bChangeMode    = FALSE;
	m_bDraw          = FALSE;

	m_bNoRestore     = FALSE;

	memset((void *)&m_logFont, 0, sizeof(m_logFont));
	m_logFont.lfHeight = 12;
	m_logFont.lfWidth = 6;
	m_logFont.lfCharSet = DEFAULT_CHARSET;
	m_logFont.lfOutPrecision = OUT_TT_PRECIS;
	m_logFont.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	m_logFont.lfQuality = DEFAULT_QUALITY;
	m_logFont.lfPitchAndFamily = FIXED_PITCH|FF_DONTCARE;
	strcpy(m_logFont.lfFaceName, "ﾋﾎﾌ�");

	// DirectDraw
	m_lpDD          = NULL;
	m_lpDDPrimary   = NULL;
	m_lpDDBack      = NULL;
	m_lpDDRender    = NULL;
	m_lpDDAscii     = NULL;		//ﾏﾔﾊｾﾌ睫ｾﾎﾄﾗﾖ
	m_lpDDZapper    = NULL;
	m_lpDDTV        = NULL;		//ﾏﾔﾊｾﾓﾎﾏｷｻｭﾃ�
	m_lpDDClipper   = NULL;
	m_lpDDClipper2  = NULL;
	m_lpDDPalette   = NULL;

	m_lpRender = NULL;

	m_hPalette = NULL;

	m_DisplayModeNum = 0;

	// デフォルト解像度
	m_dwDisplayWidth  = 640;
	m_dwDisplayHeight = 480;
	m_dwDisplayDepth  = 16;
	m_dwDisplayRate   = 0;	// リフレッシュレート指定無し

	// パレット
	m_nPaletteMode = 0;
	m_bMonoMode = FALSE;

	memcpy( m_PaletteBuf, m_PalDefault, sizeof(m_PaletteBuf) );

	ZeroMemory( m_cpPalette, sizeof(m_cpPalette) );
	ZeroMemory( m_mpPalette, sizeof(m_mpPalette) );
	ZeroMemory( m_cnPalette, sizeof(m_cnPalette) );
	ZeroMemory( m_csPalette, sizeof(m_csPalette) );
	ZeroMemory( m_mnPalette, sizeof(m_mnPalette) );
	ZeroMemory( m_msPalette, sizeof(m_msPalette) );
}

CDirectDraw::~CDirectDraw()
{
	ReleaseDDraw();
}

//////////////////////////////////////////////////////////////////////
// メンバ関数
//////////////////////////////////////////////////////////////////////

// コールバック関数
HRESULT	WINAPI	CDirectDraw::EnumModesCallback( LPDDSURFACEDESC2 lpDDSurfaceDesc, LPVOID lpContext )
{
	// Thisポインタを受け取る
	CDirectDraw* pDD = (CDirectDraw*)lpContext;

	// 256x240以下の解像度は除外(ex. 320x200)
	if( lpDDSurfaceDesc->dwWidth < 256 || lpDDSurfaceDesc->dwHeight < 240 )
		return	DDENUMRET_OK;

	// 8,16,24,32bit以外は除外(4bitがあるね^^;)
	if( !(lpDDSurfaceDesc->ddpfPixelFormat.dwRGBBitCount == 8
	   || lpDDSurfaceDesc->ddpfPixelFormat.dwRGBBitCount == 16
	   || lpDDSurfaceDesc->ddpfPixelFormat.dwRGBBitCount == 24
	   || lpDDSurfaceDesc->ddpfPixelFormat.dwRGBBitCount == 32) )
		return	DDENUMRET_OK;

	// スクエアサイズの解像度以外は除外(ex. 640x400)
	if( lpDDSurfaceDesc->dwWidth*3 != lpDDSurfaceDesc->dwHeight*4 )
		return	DDENUMRET_OK;

	// サイズ／ビット深度／リフレッシュレート／ピクセルフォーマットのコピー
	pDD->m_DisplayModeBuf[ pDD->m_DisplayModeNum ].dwWidth  = lpDDSurfaceDesc->dwWidth;
	pDD->m_DisplayModeBuf[ pDD->m_DisplayModeNum ].dwHeight = lpDDSurfaceDesc->dwHeight;
	pDD->m_DisplayModeBuf[ pDD->m_DisplayModeNum ].dwDepth  = lpDDSurfaceDesc->ddpfPixelFormat.dwRGBBitCount;
	pDD->m_DisplayModeBuf[ pDD->m_DisplayModeNum ].dwRate   = lpDDSurfaceDesc->dwRefreshRate;

	// 列挙数が超えた場合はキャンセル
	if( ++pDD->m_DisplayModeNum > CDirectDraw::DD_DISPLAYMODEMAX-1 ) {
//		DEBUGOUT( "CDirectDraw:Maximum display modes over.\n" );
		return	DDENUMRET_CANCEL;
	}
	return	DDENUMRET_OK;
}

#define	COMUSE	TRUE

// DirectDrawの初期化
BOOL	CDirectDraw::InitialDDraw( HWND hWnd )
{
	try {
		// DirectDrawオブジェクトの作成
#if	!COMUSE
		if( !m_bUseHEL ) {
			if( DirectDrawCreateEx(NULL, (LPVOID*)&m_lpDD, IID_IDirectDraw7, NULL) != DD_OK ) {
				m_lpDD = NULL;
				throw	"CDirectDraw:DirectDrawCreateEx failed.";
			}
		} else {
			if( DirectDrawCreateEx((LPGUID)DDCREATE_EMULATIONONLY, (LPVOID*)&m_lpDD, IID_IDirectDraw7, NULL) != DD_OK ) {
				m_lpDD = NULL;
				throw	"CDirectDraw:DirectDrawCreateEx failed.";
			}
		}
#else
		// COM的利用(ライブラリをリンクするだけで起動が遅くなるので....)
		COM::AddRef();
		if( ::CoCreateInstance( CLSID_DirectDraw, NULL, CLSCTX_ALL, IID_IDirectDraw7, (LPVOID*)&m_lpDD ) != S_OK ) {
			m_lpDD = NULL;
			throw	"CDirectDraw:CoCreateInstance failed.";
		}
		if( !m_bUseHEL ) {
			if( m_lpDD->Initialize( NULL ) != DD_OK )
				throw	"CDirectDraw:IDirectDraw7->Initialize failed.";
		} else {
			if( m_lpDD->Initialize( (LPGUID)DDCREATE_EMULATIONONLY ) != DD_OK )	// HEL
				throw	"CDirectDraw:IDirectDraw7->Initialize failed.";
		}
#endif
		// モニタ解像度の調査
		DDSURFACEDESC2	ddsd;
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		if( m_lpDD->GetDisplayMode( &ddsd ) != DD_OK )
			throw	"CDirectDraw:GetDisplayMode failed.";

		if( ddsd.ddpfPixelFormat.dwRGBBitCount < 8 )
			throw	"CDirectDraw:Unsupported display mode.";

		// 利用可能なディスプレイモードの取得
		m_DisplayModeNum = 0;
		if( m_lpDD->EnumDisplayModes(DDEDM_REFRESHRATES, NULL, (LPVOID)this, (LPDDENUMMODESCALLBACK2)EnumModesCallback) != DD_OK )
			throw	"CDirectDraw:EnumDisplayModes failed.";
		if( !m_DisplayModeNum )
			throw	"CDirectDraw:No display modes available.";

		m_hWnd = hWnd;

		// デフォルトパレット設定
		memcpy( m_PaletteBuf, m_PalDefault, sizeof(m_PalDefault) );

		// Render screen
		if( !m_lpRender ) {
			if( !(m_lpRender = (LPBYTE)malloc( RENDER_WIDTH*RENDER_HEIGHT )) )
				throw	"CDirectDraw::Out of memory.";

			FillMemory( m_lpRender, RENDER_WIDTH*RENDER_HEIGHT, 0x3F );
		}
	} catch( char *str ) {
		m_DisplayModeNum = 0;
		RELEASE( m_lpDD );
		::MessageBox( hWnd, str, "ERROR", MB_ICONERROR|MB_OK );

		return	FALSE;
	}

	return	TRUE;
}

// DirectDrawの開放
void	CDirectDraw::ReleaseDDraw()
{
	ReleaseSurface();

	if( m_lpDD ) {
		RELEASE( m_lpDD );
	}
#if	COMUSE
	COM::Release();
#endif

	FREE( m_lpRender );
//	m_hWnd = NULL;
}

// サーフェスの初期化
BOOL	CDirectDraw::InitialSurface( BOOL bScreenMode )
{
INT	i;
DDSURFACEDESC2	ddsd;
DDSCAPS2	ddscaps;
DDBLTFX		ddbltfx;

	try {
		if( !m_lpDD )
			throw "CDirectDraw:DirectDraw object uninitialized.";

		m_bGDI = FALSE;
		m_bScreenMode = bScreenMode;
		m_bMessage = FALSE;

		if( !m_bScreenMode ) {
		// ウインドウ
			// 協調モード
			if( m_lpDD->SetCooperativeLevel( m_hWnd, DDSCL_NORMAL ) != DD_OK )
				throw	"CDirectDraw:SetCooperativeLevel failed.";

			// プライマリサーフェス
			ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
			ddsd.dwSize = sizeof(DDSURFACEDESC2);
			ddsd.dwFlags = DDSD_CAPS;
			ddsd.ddsCaps.dwCaps = DDSCAPS_PRIMARYSURFACE;
			if( m_lpDD->CreateSurface( &ddsd, &m_lpDDPrimary, NULL ) != DD_OK )
				throw	"CDirectDraw:CreateSurface failed.";

			// クリッパーの作成
			if( m_lpDD->CreateClipper( 0, &m_lpDDClipper, NULL ) != DD_OK )
				throw	"CDirectDraw:CreateClipper failed.";

			m_lpDDClipper->SetHWnd( 0, m_hWnd );
			m_lpDDPrimary->SetClipper( m_lpDDClipper );
			RELEASE( m_lpDDClipper );
		} else {
		// フルスクリーン
			// 排他モード
			if( m_lpDD->SetCooperativeLevel( m_hWnd, DDSCL_ALLOWREBOOT|DDSCL_EXCLUSIVE|DDSCL_FULLSCREEN ) != DD_OK )
				throw	"CDirectDraw:SetCooperativeLevel error";

			// 画面解像度の設定
			if( m_lpDD->SetDisplayMode( m_dwDisplayWidth, m_dwDisplayHeight,
						    m_dwDisplayDepth, m_dwDisplayRate,
							DDSDM_STANDARDVGAMODE ) != DD_OK ) {
				// 失敗したらリフレッシュレート指定を無くしてもう一度やってみる
				if( m_lpDD->SetDisplayMode( m_dwDisplayWidth, m_dwDisplayHeight,
							    m_dwDisplayDepth, 0,
								DDSDM_STANDARDVGAMODE ) != DD_OK ) {
					throw	"CDirectDraw:SetDisplayMode failed.";
				} else {
					m_dwDisplayRate = 0;
				}
			}

			// プライマリサーフェスとバックサーフェスの作成
			ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
			ddsd.dwSize = sizeof(DDSURFACEDESC2);
			ddsd.dwFlags = DDSD_CAPS | DDSD_BACKBUFFERCOUNT;
			ddsd.ddsCaps.dwCaps = DDSCAPS_PRIMARYSURFACE | DDSCAPS_FLIP | DDSCAPS_COMPLEX;
			ddsd.dwBackBufferCount = 1;
			if( m_lpDD->CreateSurface(&ddsd, &m_lpDDPrimary, NULL) != DD_OK )
				throw	"CDirectDraw:CreateSurface failed.";

			ZeroMemory( &ddscaps, sizeof(DDSCAPS2) );
			ddscaps.dwCaps = DDSCAPS_BACKBUFFER;
			if( m_lpDDPrimary->GetAttachedSurface(&ddscaps, &m_lpDDBack) != DD_OK )
				throw	"CDirectDraw:GetAttachedSurface failed.";
		}

		// offscreen surface
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		ddsd.dwFlags = DDSD_WIDTH | DDSD_HEIGHT | DDSD_CAPS;
		if( !m_bSystemMemory ) {
			ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_VIDEOMEMORY;
		} else {
			ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_SYSTEMMEMORY;
		}
		ddsd.dwWidth   = SCREEN_WIDTH*2;
		ddsd.dwHeight  = SCREEN_HEIGHT*2;
		if( m_lpDD->CreateSurface(&ddsd, &m_lpDDRender, NULL) != DD_OK ) {
			// システムメモリにも取れないってか....
			if( ddsd.ddsCaps.dwCaps & DDSCAPS_SYSTEMMEMORY )
				throw	"CDirectDraw:CreateSurface failed.";

			// ビデオメモリに取れない時，しょうがないのでシステムメモリに取る
			ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_SYSTEMMEMORY;
			if( m_lpDD->CreateSurface(&ddsd, &m_lpDDRender, NULL) != DD_OK )
				throw	"CDirectDraw:CreateSurface failed.";
		}

		typedef	struct tagDDRGNDATA {
			RGNDATAHEADER	rdh;
			RECT		rect[8];
		} DDRGNDATA;

		DDRGNDATA	rgn;
		rgn.rdh.dwSize = sizeof(RGNDATAHEADER);
		rgn.rdh.iType  = RDH_RECTANGLES;
		rgn.rdh.nCount = 1;
		rgn.rdh.nRgnSize = sizeof(RECT);
		rgn.rdh.rcBound.left   = 0;
		rgn.rdh.rcBound.top    = 0;
		rgn.rdh.rcBound.right  = ddsd.dwWidth;
		rgn.rdh.rcBound.bottom = ddsd.dwHeight;
		rgn.rect[0].left   = 0;
		rgn.rect[0].top    = 0;
		rgn.rect[0].right  = ddsd.dwWidth;
		rgn.rect[0].bottom = ddsd.dwHeight;

		// Rendering surface clipper
		if( m_lpDD->CreateClipper( 0, &m_lpDDClipper2, NULL ) != DD_OK )
			throw	"CDirectDraw:CreateClipper failed.";

		m_lpDDClipper2->SetClipList( (LPRGNDATA)&rgn, 0 );
		m_lpDDRender->SetClipper( m_lpDDClipper2 );



		// Char surface(2character)
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		ddsd.dwFlags = DDSD_WIDTH | DDSD_HEIGHT | DDSD_CAPS;
		ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_SYSTEMMEMORY;		
		ddsd.dwWidth   = 16;
		ddsd.dwHeight  = 16;
		if( m_lpDD->CreateSurface(&ddsd, &m_lpDDChar, NULL) != DD_OK )
			throw	"CDirectDraw:CreateSurface failed.";
		//Color key
		DDCOLORKEY	ddck;
		ZeroMemory( &ddck, sizeof(DDCOLORKEY) );
		ddck.dwColorSpaceLowValue  = 0;
		ddck.dwColorSpaceHighValue = 0;
		m_lpDDChar->SetColorKey( DDCKEY_SRCBLT, &ddck );

		// Ascii surface(6x6x64character)
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		ddsd.dwFlags = DDSD_WIDTH | DDSD_HEIGHT | DDSD_CAPS;
//		ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_VIDEOMEMORY;
		ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_SYSTEMMEMORY;
		ddsd.dwWidth   = 6*8;
		ddsd.dwHeight  = 6*8;
		if( m_lpDD->CreateSurface(&ddsd, &m_lpDDAscii, NULL) != DD_OK )
			throw	"CDirectDraw:CreateSurface failed.";
		// Color key
		ZeroMemory( &ddck, sizeof(DDCOLORKEY) );
		ddck.dwColorSpaceLowValue  = 0;
		ddck.dwColorSpaceHighValue = 0;
		m_lpDDAscii->SetColorKey( DDCKEY_SRCBLT, &ddck );

		// Zapper surface
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		ddsd.dwFlags = DDSD_WIDTH | DDSD_HEIGHT | DDSD_CAPS;
//		ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_VIDEOMEMORY;
		ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_SYSTEMMEMORY;
		ddsd.dwWidth   = 16;
		ddsd.dwHeight  = 16;
		if( m_lpDD->CreateSurface(&ddsd, &m_lpDDZapper, NULL) != DD_OK )
			throw	"CDirectDraw:CreateSurface failed.";
		// Color key
		ZeroMemory( &ddck, sizeof(DDCOLORKEY) );
		ddck.dwColorSpaceLowValue  = 0;
		ddck.dwColorSpaceHighValue = 0;
		m_lpDDZapper->SetColorKey( DDCKEY_SRCBLT, &ddck );

		// TV
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		ddsd.dwFlags = DDSD_WIDTH | DDSD_HEIGHT | DDSD_CAPS;
		if( !m_bSystemMemory ) {
			ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_VIDEOMEMORY;
		} else {
			ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_SYSTEMMEMORY;
		}
		ddsd.dwWidth   = 512;
		ddsd.dwHeight  = 480;
		if( m_lpDD->CreateSurface(&ddsd, &m_lpDDTV, NULL) != DD_OK ) 
		{
			// システムメモリにも取れないってか....
			if( ddsd.ddsCaps.dwCaps & DDSCAPS_SYSTEMMEMORY )
				throw	"CDirectDraw:CreateSurface failed.";

			// ビデオメモリに取れない時，しょうがないのでシステムメモリに取る
			ddsd.ddsCaps.dwCaps = DDSCAPS_OFFSCREENPLAIN | DDSCAPS_SYSTEMMEMORY;
			if( m_lpDD->CreateSurface(&ddsd, &m_lpDDTV, NULL) != DD_OK )
				throw	"CDirectDraw:CreateSurface failed.";
		}

		// Color key
		ZeroMemory( &ddck, sizeof(DDCOLORKEY) );
		ddck.dwColorSpaceLowValue  = 0;
		ddck.dwColorSpaceHighValue = 0;
		m_lpDDTV->SetColorKey( DDCKEY_SRCBLT, &ddck );

		// DirectDraw/GDI Palette
		ZeroMemory( &m_logPalette, sizeof(m_logPalette) );
		m_logPalette.palVersion    = 0x0300;
		m_logPalette.palNumEntries = 256;

		HDC	hdc = ::GetDC( NULL );
		GetSystemPaletteEntries( hdc, 0, 256, m_logPalette.pe );
		ReleaseDC( NULL, hdc );

		for( i = 0; i < 10; i++ ) {
			m_logPalette.pe[i    ].peFlags = PC_EXPLICIT;
			m_logPalette.pe[i+246].peFlags = PC_EXPLICIT;
		}
		for( i = 10; i < 246; i++ ) {
			m_logPalette.pe[i].peRed   = 0;
			m_logPalette.pe[i].peGreen = 0;
			m_logPalette.pe[i].peBlue  = 0;
			if( i >= 0x10 && i < 0x20 ) {
				m_logPalette.pe[i].peRed   = (i-0x10)*0x10;
				m_logPalette.pe[i].peGreen = (i-0x10)*0x10;
				m_logPalette.pe[i].peBlue  = (i-0x10)*0x10;
				m_logPalette.pe[i].peFlags = PC_RESERVED;
			} else if( i >= 0x40 && i < 0xC0 ) {
				m_logPalette.pe[i].peFlags = PC_RESERVED;
			} else {
				m_logPalette.pe[i].peFlags = PC_NOCOLLAPSE;
			}
		}

		// Surface clear
		ddbltfx.dwSize = sizeof(DDBLTFX);
		ddbltfx.dwFillColor = 0;
		if( m_lpDDBack ) {
			m_lpDDBack->Blt(NULL, NULL, NULL, DDBLT_COLORFILL|DDBLT_WAIT, &ddbltfx);
		}
		m_lpDDRender->Blt (NULL, NULL, NULL, DDBLT_COLORFILL|DDBLT_WAIT, &ddbltfx);

		// Palette object
		DDSURFACEDESC2	ddsd;
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		m_lpDDPrimary->GetSurfaceDesc(&ddsd);

		if( ddsd.ddpfPixelFormat.dwRGBBitCount == 8 ) {
			if( !m_bScreenMode ) {
				if( (m_hPalette = CreatePalette( (LOGPALETTE *)&m_logPalette )) == NULL )
					throw	"CDirectDraw:CreatePalette failed.";
			} else {
				if( m_lpDD->CreatePalette( DDPCAPS_8BIT, m_logPalette.pe, &m_lpDDPalette, NULL ) != DD_OK )
					throw	"CDirectDraw:CreatePalette error";
				if( m_lpDDPrimary->SetPalette( m_lpDDPalette ) != DD_OK )
					throw	"CDirectDraw:SetPalette failed.";
			}
		}

		// Palette calculate
		CalcPaletteTable();
		// Palette Realize
		RealizePalette();
		m_bPaletteUpdate = TRUE;

		// Character setup
		SetLZSSChar( lzAscii, m_lpDDAscii );
		SetLZSSChar( lzSight, m_lpDDZapper );
		SetLZSSChar( lzTVlayer, m_lpDDTV );
	} catch( char *str ) {
		ReleaseSurface();

//		DEBUGOUT( "%s\n", str );
		::MessageBox( m_hWnd, str, "ERROR", MB_ICONERROR|MB_OK );

		return	FALSE;
	}
//	DEBUGOUT( "CDirectDraw:InitialSurface complete.\n" );

	return	TRUE;
}

// サーフェスの開放
BOOL	CDirectDraw::ReleaseSurface( void )
{
	if( !m_lpDD )
		return	FALSE;

	GDIDELETE( m_hPalette );
	RELEASE( m_lpDDPalette );

	RELEASE( m_lpDDClipper2 );
	RELEASE( m_lpDDClipper );

	RELEASE( m_lpDDTV );
	RELEASE( m_lpDDZapper );
	RELEASE( m_lpDDAscii );
	RELEASE( m_lpDDRender );
	RELEASE( m_lpDDBack );
	RELEASE( m_lpDDPrimary );
	RELEASE( m_lpDDChar );	
	return	TRUE;
}

// サーフェスのリストアサブ
BOOL	CDirectDraw::RestoreSurfaceSub( LPDIRECTDRAWSURFACE7 lpSurface )
{
	if( lpSurface ) {
		if( lpSurface->IsLost() == DDERR_SURFACELOST ) {
			// ロストしていたらリストアする
//			DEBUGOUT( "CDirectDraw:Restore surface...." );
			if( lpSurface->Restore() == DD_OK ) {
//				DEBUGOUT( "Ok.\n" );
			} else {
//				DEBUGOUT( "Failed.\n" );
			}
			return	FALSE;
		}
	}
	return	TRUE;
}

// サーフェスのリストア
BOOL	CDirectDraw::RestoreSurface()
{
	if( !m_lpDD ) return FALSE;

	if( m_bNoRestore )
		return	FALSE;

	BOOL	bRet = TRUE;

	// ロストしてたら内容は失われているので描画をキャンセルしてね
	if( !RestoreSurfaceSub( m_lpDDRender ) ) {
		bRet = FALSE;
	}
	if( !RestoreSurfaceSub( m_lpDDBack ) ) {
		bRet = FALSE;
	}
	if( !RestoreSurfaceSub( m_lpDDPrimary ) ) {
		bRet = FALSE;
	}
	if( !RestoreSurfaceSub( m_lpDDAscii ) ) {
		SetLZSSChar( lzAscii, m_lpDDAscii );
		bRet = FALSE;
	}
	if( !RestoreSurfaceSub( m_lpDDZapper ) ) {
		SetLZSSChar( lzSight, m_lpDDZapper );
		bRet = FALSE;
	}
	if( !RestoreSurfaceSub( m_lpDDTV ) ) {
		SetLZSSChar( lzTVlayer, m_lpDDTV );
		bRet = FALSE;
	}
	if( !RestoreSurfaceSub( m_lpDDChar ) ) {
		bRet = FALSE;
	}
	return	bRet;
}

//
// ディスプレイモード変更開始
//
BOOL	CDirectDraw::BeginDisplayChange()
{
	if( !m_lpDD )
		return	FALSE;

	// 既に変更中やないか
	if( m_bChangeMode )
		return	FALSE;

	// 変更中じゃ
	m_bChangeMode = TRUE;

	if( m_bScreenMode ) {
		m_lpDD->RestoreDisplayMode();
	}

	// サーフェスの開放
	ReleaseSurface();

	// 協調モード
	m_lpDD->SetCooperativeLevel( NULL, DDSCL_NORMAL );

	return	TRUE;
}

//
// ディスプレイモード変更終了とサーフェスの再構築
//
BOOL	CDirectDraw::EndDisplayChange()
{
	if( !m_lpDD || !m_bChangeMode )
		return	FALSE;

	// Rebuild surface
	if( !InitialSurface( m_bScreenMode ) ) {
		m_bChangeMode = FALSE;
		return	FALSE;
	}

	// Wait
	::Sleep( 250 );

	// Change mode complete
	m_bChangeMode = FALSE;

	return	TRUE;
}

//
// WM_DISPLAYCHANGEハンドラから呼ばれる
//
BOOL	CDirectDraw::OnChangeDisplayMode()
{
	// 自分自身で変更している時はキャンセル
	if( m_bChangeMode )
		return	TRUE;

	// DirectDrawオブジェクトが無ければ意味無し
	if( !m_lpDD )
		return	FALSE;

//	DEBUGOUT( "CDirectDraw:OnChangeDisplayMode\n" );

	// 協調モードのチェック
	HRESULT	hr = m_lpDD->TestCooperativeLevel();

	if( !m_bScreenMode ) {
	// ウインドウ
		if( hr == DDERR_EXCLUSIVEMODEALREADYSET ) {
			// 他に排他モード野郎がおるのでおしまい
//			DEBUGOUT( "CDirectDraw:DDERR_EXCLUSIVEMODEALREADYSET\n" );
			// リストアしちゃいやん
			m_bNoRestore = TRUE;
			return	TRUE;
		} else if( hr == DDERR_WRONGMODE || hr == DD_OK ) {
			// 普通のディスプレイモードの変更(何故かWindowメッセージだとDD_OKな時がある)
//			DEBUGOUT( "CDirectDraw:DDERR_WRONGMODE\n" );
			// リストアしてええよ
			m_bNoRestore = FALSE;

			// サーフェスの開放
			ReleaseSurface();

			// サーフェスの初期化
			BOOL	bRet = InitialSurface( m_bScreenMode );

			if( bRet ) {
//				DEBUGOUT( "CDirectDraw:InitialSurface ok.\n" );
			} else {
//				DEBUGOUT( "CDirectDraw:InitialSurface failed.\n" );
			}

			return	bRet;
		} else if( hr == DDERR_INVALIDOBJECT ) {
//			DEBUGOUT( "CDirectDraw:DDERR_INVALIDOBJECT\n" );
			return	FALSE;
		} else if( hr == DDERR_NOEXCLUSIVEMODE ) {
//			DEBUGOUT( "CDirectDraw:DDERR_NOEXCLUSIVEMODE\n" );
		} else {
//			DEBUGOUT( "CDirectDraw:Unknown error. hr=%08Xた\n", hr );
		}
		// 良く分からんのでとりあえずサーフェスをリストア
		m_bNoRestore = FALSE;
		RestoreSurface();

		return	TRUE;
	} else {
	// フルスクリーン
		if( hr == DDERR_NOEXCLUSIVEMODE ) {
			// フルスクリーンモードから抜けた時
//			DEBUGOUT( "CDirectDraw:DDERR_NOEXCLUSIVEMODE\n" );
			return	TRUE;
		}

		if( hr == DD_OK ) {
			// リストアしろって事か...
			// サーフェスをリストア
			RestoreSurface();

			return	TRUE;
		}

		// 不明な原因
		return	FALSE;
	}

	return	TRUE;
}


// ビット位置の取得
void	CDirectDraw::GetBitMask( DWORD val, int& shift, int& bits )
{
	shift = 0;
	while( !(val & (1<<shift)) && (shift<32) ) {
		shift++;
	}

	bits = 32;
	while( !(val & (1<<(bits-1))) && (bits>0) ) {
		bits--;
	}
	bits = bits - shift;
}

static	float	PalConvTbl[][3] = {
	1.00f, 1.00f, 1.00f,
	1.00f, 0.80f, 0.73f,
	0.73f, 1.00f, 0.70f,
	0.76f, 0.78f, 0.58f,
	0.86f, 0.80f, 1.00f,
	0.83f, 0.68f, 0.85f,
	0.67f, 0.77f, 0.83f,
	0.68f, 0.68f, 0.68f,
//	1.00f, 1.00f, 1.00f,
};

// パレットテーブルの計算
BOOL	CDirectDraw::CalcPaletteTable()
{
INT	i, j;

	if( !m_lpDD || !m_lpDDPrimary )
		return	FALSE;

	DDSURFACEDESC2	ddsd;
	ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
	ddsd.dwSize = sizeof(DDSURFACEDESC2);
	ddsd.dwFlags = DDSD_PIXELFORMAT;
	if( m_lpDDPrimary->GetSurfaceDesc(&ddsd) != DD_OK )
		throw "CDirectDraw:GetSurfaceDesc error";

	DWORD	Rn, Gn, Bn, Yn;
	DWORD	Rs, Gs, Bs, Ys;
	INT	Rbit, Gbit, Bbit;
	INT	Rsft, Gsft, Bsft;

	if( ddsd.ddpfPixelFormat.dwRGBBitCount != 8 ) {
		GetBitMask( ddsd.ddpfPixelFormat.dwRBitMask, Rsft, Rbit );
		GetBitMask( ddsd.ddpfPixelFormat.dwGBitMask, Gsft, Gbit );
		GetBitMask( ddsd.ddpfPixelFormat.dwBBitMask, Bsft, Bbit );
	}

	for( j = 0; j < 8; j++ ) {
		for( i = 0; i < 64; i++ ) {
			// Normal
			Rn = (DWORD)(PalConvTbl[j][0]*m_PaletteBuf[i].r);
			Gn = (DWORD)(PalConvTbl[j][1]*m_PaletteBuf[i].g);
			Bn = (DWORD)(PalConvTbl[j][2]*m_PaletteBuf[i].b);
			// Scanline
			Rs = (DWORD)(PalConvTbl[j][0]*m_PaletteBuf[i].r*m_nScanlineColor/100.0f);
			Gs = (DWORD)(PalConvTbl[j][1]*m_PaletteBuf[i].g*m_nScanlineColor/100.0f);
			Bs = (DWORD)(PalConvTbl[j][2]*m_PaletteBuf[i].b*m_nScanlineColor/100.0f);

			m_cpPalette[j][i+0x00].rgbRed   = (BYTE)Rn;
			m_cpPalette[j][i+0x00].rgbGreen = (BYTE)Gn;
			m_cpPalette[j][i+0x00].rgbBlue  = (BYTE)Bn;
			m_cpPalette[j][i+0x40].rgbRed   = (BYTE)Rs;
			m_cpPalette[j][i+0x40].rgbGreen = (BYTE)Gs;
			m_cpPalette[j][i+0x40].rgbBlue  = (BYTE)Bs;

			m_cnPalette[j][i] = ((Rn>>(8-Rbit))<<Rsft)|((Gn>>(8-Gbit))<<Gsft)|((Bn>>(8-Bbit))<<Bsft);
			m_csPalette[j][i] = ((Rs>>(8-Rbit))<<Rsft)|((Gs>>(8-Gbit))<<Gsft)|((Bs>>(8-Bbit))<<Bsft);

			// Monochrome
			Rn = (DWORD)(PalConvTbl[j][0]*m_PaletteBuf[i&0x30].r);
			Gn = (DWORD)(PalConvTbl[j][1]*m_PaletteBuf[i&0x30].g);
			Bn = (DWORD)(PalConvTbl[j][2]*m_PaletteBuf[i&0x30].b);
			// Scanline
			Rs = (DWORD)(PalConvTbl[j][0]*m_PaletteBuf[i&0x30].r*m_nScanlineColor/100.0f);
			Gs = (DWORD)(PalConvTbl[j][1]*m_PaletteBuf[i&0x30].g*m_nScanlineColor/100.0f);
			Bs = (DWORD)(PalConvTbl[j][2]*m_PaletteBuf[i&0x30].b*m_nScanlineColor/100.0f);

			Yn = (DWORD)(0.299f * Rn + 0.587f * Gn + 0.114f * Bn);
			Ys = (DWORD)(0.299f * Rs + 0.587f * Gs + 0.114f * Bs);
			if( Yn > 0xFF ) Yn = 0xFF;
			if( Ys > 0xFF ) Ys = 0xFF;

			m_mpPalette[j][i+0x00].rgbRed = m_mpPalette[j][i+0x00].rgbGreen = m_mpPalette[j][i+0x00].rgbBlue = (BYTE)Yn;
			m_mpPalette[j][i+0x40].rgbRed = m_mpPalette[j][i+0x40].rgbGreen = m_mpPalette[j][i+0x40].rgbBlue = (BYTE)Ys;

			m_cnPalette[j][i+0x40] = ((Yn>>(8-Rbit))<<Rsft)|((Yn>>(8-Gbit))<<Gsft)|((Yn>>(8-Bbit))<<Bsft);
			m_csPalette[j][i+0x40] = ((Ys>>(8-Rbit))<<Rsft)|((Ys>>(8-Gbit))<<Gsft)|((Ys>>(8-Bbit))<<Bsft);

			m_mnPalette[j][i] = ((Yn>>(8-Rbit))<<Rsft)|((Yn>>(8-Gbit))<<Gsft)|((Yn>>(8-Bbit))<<Bsft);
			m_msPalette[j][i] = ((Ys>>(8-Rbit))<<Rsft)|((Ys>>(8-Gbit))<<Gsft)|((Ys>>(8-Bbit))<<Bsft);
		}
	}

	return	TRUE;
}

// パレットテーブルの設定
void	CDirectDraw::SetPaletteTable( LPBYTE pal )
{
	if( pal )
		memcpy( m_PaletteBuf, pal, sizeof(m_PaletteBuf) );
	else
		memcpy( m_PaletteBuf, m_PalDefault, sizeof(m_PaletteBuf) );

	CalcPaletteTable();

	m_bPaletteUpdate = TRUE;
}

// パレットファイルの更新
void	CDirectDraw::SetPaletteFile( LPCTSTR fname )
{
	// パレットファイルの更新
	if( strlen( fname ) > 0 ) {
		FILE	*fp;
		if( (fp = ::fopen( fname, "rb" )) ) {
			BYTE	palbuf[192];
			// サイズ分読み込み
			if( ::fread( palbuf, 192, 1, fp ) == 1 ) {
				// パレットの変更と計算
				SetPaletteTable( palbuf );
			} else {
				// 読みきれなかった時はデフォルト
				SetPaletteTable( NULL );
			}
			FCLOSE(fp);
		} else {
			// 開けなかった時はデフォルト
			SetPaletteTable( NULL );
		}
	} else {
		// 名前が無い時はデフォルト
		SetPaletteTable( NULL );
	}
}

// フルスクリーンモードでのGDIウインドウ表示設定
BOOL	CDirectDraw::SetFullScreenGDI( BOOL bMode )
{
	// 念の為チェック
	if( !m_lpDD || !m_lpDDPrimary )
		return	FALSE;

	if( m_bScreenMode ) {
		if( !m_bGDI ) {
			if( bMode ) {
				RELEASE( m_lpDDClipper ); // 一応
				if( m_lpDD->CreateClipper(0, &m_lpDDClipper, NULL) == DD_OK ) {
					m_lpDDClipper->SetHWnd( 0, m_hWnd );
					m_lpDDPrimary->SetClipper( m_lpDDClipper );
					if( m_lpDD->FlipToGDISurface() == DD_OK ) {
						m_bGDI = TRUE;
					} else {
						RELEASE( m_lpDDClipper );
						return	FALSE;
					}
				}
			}
		} else {
			if( !bMode ) {
				RELEASE( m_lpDDClipper );
				m_bGDI = FALSE;
			}
		}
	}
	return	TRUE;
}

void	CDirectDraw::RealizePalette()
{
	if( !m_lpDD || !m_lpDDPrimary )
		return;

	if( !m_bScreenMode ) {
		DDSURFACEDESC2	ddsd;
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		m_lpDDPrimary->GetSurfaceDesc(&ddsd);

		if( ddsd.ddpfPixelFormat.dwRGBBitCount == 8 ) {
			HDC	hdc;
			hdc = ::GetDC( m_hWnd );
			::SelectPalette( hdc, m_hPalette, FALSE );
			::RealizePalette( hdc );
			::ReleaseDC( m_hWnd, hdc );

			m_bPaletteUpdate = TRUE;
		}
	}
}

// 描画(Windowsメッセージ用)
void	CDirectDraw::OnScreenDraw()
{
	if( !m_bScreenMode ) {
	// Window mode
		Blt();
		Flip();
	} else {
	// Fullscreen mode
		if( m_bGDI ) {
			Blt();
			Flip();
		}
	}
}

void	CDirectDraw::SetPaletteMode( INT nMode, BOOL bMono )
{
	if( (m_nPaletteMode != nMode) || (m_bMonoMode != bMono) ) {
		m_bPaletteUpdate = TRUE;
	}
	m_nPaletteMode = nMode;
	m_bMonoMode    = bMono;
}

// パレットテーブルの取得
void	CDirectDraw::GetPaletteData( RGBQUAD* rgb )
{
INT	i;
	if( !m_bMonoMode ) {
		for( i = 0; i < 64; i++ ) {
			rgb[i     ] = m_cpPalette[m_nPaletteMode][i];
			rgb[i+0x40] = m_mpPalette[m_nPaletteMode][i];
		}
	} else {
		for( i = 0; i < 64; i++ ) {
			rgb[i     ] = m_mpPalette[m_nPaletteMode][i];
			rgb[i+0x40] = m_mpPalette[m_nPaletteMode][i];
		}
	}
}

void	CDirectDraw::Blt()
{
INT	i;
DDSURFACEDESC2	ddsd;

	if( !m_lpDD || !m_lpDDPrimary )
		return;
	// ScreenMode changing?
	if( m_bChangeMode )
		return;
	// Surface lost check & restore
	if( !RestoreSurface() )
		return;

	ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
	ddsd.dwSize = sizeof(DDSURFACEDESC2);
	m_lpDDPrimary->GetSurfaceDesc(&ddsd);

	// Palette copy
	if( ddsd.ddpfPixelFormat.dwRGBBitCount == 8 ) {
		if( !m_bMonoMode ) {
			for( i = 0; i < 128; i++ ) {
				m_logPalette.pe[i+0x40].peRed   = m_cpPalette[m_nPaletteMode][i].rgbRed;
				m_logPalette.pe[i+0x40].peGreen = m_cpPalette[m_nPaletteMode][i].rgbGreen;
				m_logPalette.pe[i+0x40].peBlue  = m_cpPalette[m_nPaletteMode][i].rgbBlue;
			}
		} else {
			for( i = 0; i < 128; i++ ) {
				m_logPalette.pe[i+0x40].peRed   = m_mpPalette[m_nPaletteMode][i].rgbRed;
				m_logPalette.pe[i+0x40].peGreen = m_mpPalette[m_nPaletteMode][i].rgbGreen;
				m_logPalette.pe[i+0x40].peBlue  = m_mpPalette[m_nPaletteMode][i].rgbBlue;
			}
		}
		if( !m_bScreenMode ) {
			::AnimatePalette( m_hPalette, 0, 256, m_logPalette.pe );
		}
	}

	LPBYTE	lpRdr = m_lpRender+8;

	LPBYTE	lpPal;
	LPBYTE	lpPalScan;

	if( !m_bMonoMode ) {
		lpPal     = (LPBYTE)m_cnPalette[m_nPaletteMode];
		lpPalScan = (LPBYTE)m_csPalette[m_nPaletteMode];
	} else {
		lpPal     = (LPBYTE)m_mnPalette[m_nPaletteMode];
		lpPalScan = (LPBYTE)m_msPalette[m_nPaletteMode];
	}

	if(m_bScanlineMode)
		ViewChars(lpRdr);
	
	ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
	ddsd.dwSize  = sizeof(DDSURFACEDESC2);
	
	if( m_lpDDRender->Lock( NULL, &ddsd, 0, NULL ) != DD_OK )
		throw	"CDirectDraw:Lock error";

	if( ddsd.ddpfPixelFormat.dwRGBBitCount == 8 ) 
	{
		Render8Bit( lpRdr, ddsd );
	} 
	else if( ddsd.ddpfPixelFormat.dwRGBBitCount == 16 ) 
	{
		Render16Bit( lpRdr, ddsd, lpPal, lpPalScan );
	} 
	else if( ddsd.ddpfPixelFormat.dwRGBBitCount == 24 ) 
	{
		Render24Bit( lpRdr, ddsd, lpPal, lpPalScan );
	} 
	else if( ddsd.ddpfPixelFormat.dwRGBBitCount == 32 ) 
	{
		Render32Bit( lpRdr, ddsd, lpPal, lpPalScan );
	}
	m_lpDDRender->Unlock( NULL );

	if( m_bTVFrameMode ) {
		RECT	rcW;
		// Size calculate
		rcW.left   = 0;
		rcW.top    = 0;
		rcW.right  = SCREEN_WIDTH;
		rcW.bottom = SCREEN_HEIGHT;
		if( m_bDoubleSize ) 
		{
			rcW.right *= 2;
		}
		if( m_bScanlineMode || m_bDoubleSize ) 
		{
			rcW.bottom *= 2;
		}


		m_lpDDRender->Blt( &rcW, m_lpDDTV, NULL, DDBLT_KEYSRC | DDBLT_ROP, NULL );
	}


	if(!m_bScanlineMode)
		ViewChars(m_lpDDRender);

	// Infomation string
	if( strlen(m_szInfo) > 0 ) {
		INT	x, y;
		if( !m_bDoubleSize ) {
			x = 8;
		} else {
			x = 16;
		}
		if( !m_bAllLine ) {
			if( !m_bScanlineMode && !m_bDoubleSize ) {
				y = 12;
			} else {
				y = 24;
			}
		} else {
			if( !m_bScanlineMode && !m_bDoubleSize ) {
				y = 4;
			} else {
				y = 8;
			}
		}
		RenderString( x, y, m_szInfo );
	}
	// Message string
	if( m_bMessage ) {
		if( timeGetTime()-m_dwMessageTime > 1500 ) {
			m_bMessage = FALSE;
		}
		if( strlen(m_szMess) > 0 ) {
			INT	x, y;
			if( !m_bDoubleSize ) {
				x = 8;
			} else {
				x = 16;
			}
			if( !m_bAllLine ) {
				if( !m_bScanlineMode && !m_bDoubleSize ) {
					y = SCREEN_HEIGHT-18;
				} else {
					y = SCREEN_HEIGHT*2-36;
				}
			} else {
				if( !m_bScanlineMode && !m_bDoubleSize ) {
					y = SCREEN_HEIGHT-10;
				} else {
					y = SCREEN_HEIGHT*2-20;
				}
			}
			RenderString( x, y, m_szMess );
		}
	}

	if( m_bZapper && m_bZapperDraw ) {
		if( m_ZapperPosX >= 0 && m_ZapperPosX < SCREEN_WIDTH && m_ZapperPosY >= 0 && m_ZapperPosY < SCREEN_HEIGHT ) {
			RECT	rcS, rcZ;

			SetRect( &rcZ, 0, 0, 16, 16 );
			if( !m_bDoubleSize ) {
				rcS.left   = m_ZapperPosX-(rcZ.right-rcZ.left)/2;
				rcS.right  = m_ZapperPosX+(rcZ.right-rcZ.left)/2;
			} else {
				rcS.left   = m_ZapperPosX*2-(rcZ.right-rcZ.left);
				rcS.right  = m_ZapperPosX*2+(rcZ.right-rcZ.left);
			}
			if( !m_bDoubleSize && !m_bScanlineMode ) {
				rcS.top    = m_ZapperPosY-(rcZ.bottom-rcZ.top)/2;
				rcS.bottom = m_ZapperPosY+(rcZ.bottom-rcZ.top)/2;
			} else {
				rcS.top    = m_ZapperPosY*2-(rcZ.bottom-rcZ.top);
				rcS.bottom = m_ZapperPosY*2+(rcZ.bottom-rcZ.top);
			}

			m_lpDDRender->Blt( &rcS, m_lpDDZapper, &rcZ, DDBLT_KEYSRC, NULL );
		}
	}
}

void	CDirectDraw::Flip()
{
	if( !m_lpDD || !m_lpDDPrimary )
		return;
	// ScreenMode changing?
	if( m_bChangeMode )
		return;
	// Surface lost check & restore
	if( !RestoreSurface() )
		return;

	DDSURFACEDESC2	ddsd;
	ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
	ddsd.dwSize = sizeof(DDSURFACEDESC2);
	m_lpDDPrimary->GetSurfaceDesc(&ddsd);

	RECT	rcS, rcC;

	// Size calculate
	if( !m_bDoubleSize ) 
	{
		rcS.left  = 0;
		rcS.right = SCREEN_WIDTH;
	} 
	else 
	{
		rcS.left  = 0;
		rcS.right = SCREEN_WIDTH*2;
	}
	if( !m_bAllLine ) 
	{
		rcS.top    = 8;
		rcS.bottom = SCREEN_HEIGHT-8;
	} 
	else 
	{
		rcS.top    = 0;
		rcS.bottom = SCREEN_HEIGHT;
	}
	if( m_bScanlineMode || m_bDoubleSize ) 
	{
		rcS.top    *= 2;
		rcS.bottom *= 2;
	}

	if( !m_bScreenMode ) 
	{
	// Window mode
		::GetClientRect( m_hWnd, &rcC );
		::ClientToScreen( m_hWnd, (POINT*)&rcC.left );
		::ClientToScreen( m_hWnd, (POINT*)&rcC.right );

		if( ddsd.ddpfPixelFormat.dwRGBBitCount == 8 ) 
		{
			HDC	hdc;
			if( m_lpDDPrimary->GetDC( &hdc ) == DD_OK ) 
			{
				::SelectPalette( hdc, m_hPalette, FALSE );
				::RealizePalette( hdc );
				m_lpDDPrimary->ReleaseDC( hdc );
			}
		}
	} 
	else 
	{
	// Fullscreen mode
		if( !m_bMaxZoom ) 
		{
			DDBLTFX	ddbltfx;
			ddbltfx.dwSize = sizeof(DDBLTFX);
			ddbltfx.dwFillColor = 0;
			m_lpDDBack->Blt(NULL, NULL, NULL, DDBLT_COLORFILL|DDBLT_WAIT, &ddbltfx);

			// Position offset caluclate
			LONG	swidth, sheight;
			LONG	dwidth, dheight;
			LONG	hmul, vmul;

			if( !m_bAspect )  swidth = SCREEN_WIDTH;
			else		  swidth = 320;
			if( !m_bAllLine ) sheight = SCREEN_HEIGHT-16;
			else		  sheight = SCREEN_HEIGHT;

			dwidth  = (LONG)ddsd.dwWidth;
			dheight = (LONG)ddsd.dwHeight;
			hmul = dwidth  / swidth;
			vmul = dheight / sheight;

			if( vmul < hmul ) hmul = vmul;
			else		  vmul = hmul;

			rcC.left   = (dwidth -swidth *hmul)/2;
			rcC.top    = (dheight-sheight*vmul)/2;
			rcC.right  = rcC.left+swidth *hmul;
			rcC.bottom = rcC.top +sheight*vmul;
		} 
		else 
		{
			// Maximum zoom
			rcC.left   = 0;
			rcC.top    = 0;
			rcC.right  = (LONG)ddsd.dwWidth;
			rcC.bottom = (LONG)ddsd.dwHeight;
		}
	}

	if( !m_bScreenMode ) 
	{
	// Window mode
		if( m_bWindowVSync ) 
		{
			HRESULT	hr;
			while( TRUE ) 
			{
				hr = m_lpDD->WaitForVerticalBlank( DDWAITVB_BLOCKBEGIN, NULL );
				if( hr == DD_OK )
					break;
				if( hr == DDERR_SURFACELOST ) 
				{
					break;
				}
				if( hr != DDERR_WASSTILLDRAWING )
				{
					break;
				}
			}

			m_lpDDPrimary->Blt( NULL, m_lpDDBack, NULL, DDBLT_WAIT, NULL );
		}

		m_lpDDPrimary->Blt( &rcC, m_lpDDRender, &rcS, DDBLT_WAIT, NULL );

		if( ddsd.ddpfPixelFormat.dwRGBBitCount == 8 ) {
			if( m_bPaletteUpdate ) 
			{
				m_bPaletteUpdate = FALSE;
				::AnimatePalette( m_hPalette, 0, 256, m_logPalette.pe );
			}
		}
	} 
	else 
	{
	// Fullscreen mode
		m_lpDDBack->Blt( &rcC, m_lpDDRender, &rcS, DDBLT_WAIT, NULL );

		if( ddsd.ddpfPixelFormat.dwRGBBitCount == 8 ) 
		{
			if( m_bPaletteUpdate ) 
			{
				m_bPaletteUpdate = FALSE;
				m_lpDDPalette->SetEntries( 0, 0, 256, m_logPalette.pe );
			}
		}

		if( m_bFlip )
		{
			if( !m_bGDI )
			{
				HRESULT	hr;
				while( TRUE )
				{
					hr = m_lpDDPrimary->Flip( NULL, DDFLIP_WAIT );
					if( hr == DD_OK )
						break;
					if( hr == DDERR_SURFACELOST )
					{
						break;
					}
					if( hr != DDERR_WASSTILLDRAWING )
					{
						break;
					}
				}
			} 
			else 
			{
				HRESULT	hr;
				while( TRUE ) 
				{
					hr = m_lpDD->WaitForVerticalBlank( DDWAITVB_BLOCKBEGIN, NULL );
					if( hr == DD_OK )
						break;
					if( hr == DDERR_SURFACELOST ) 
					{
						break;
					}
					if( hr != DDERR_WASSTILLDRAWING ) 
					{
						break;
					}
				}
				m_lpDDPrimary->Blt( NULL, m_lpDDBack, NULL, DDBLT_WAIT, NULL );
			}
		} 
		else 
		{
			m_lpDDPrimary->Blt( NULL, m_lpDDBack, NULL, DDBLT_WAIT, NULL );
		}
	}
}

BYTE	CDirectDraw::GetZapperHit()
{
	if( m_bZapper ) {
		if( m_ZapperPosX >= 0 && m_ZapperPosX < SCREEN_WIDTH && m_ZapperPosY >= 0 && m_ZapperPosY < SCREEN_HEIGHT ) {
			BYTE	c = m_lpRender[8+m_ZapperPosX+RENDER_WIDTH*m_ZapperPosY];
			DWORD	Yn = (DWORD)(0.299f * m_cpPalette[0][c].rgbRed + 0.587f * m_cpPalette[0][c].rgbGreen + 0.114f * m_cpPalette[0][c].rgbBlue);
			if( Yn > 0xFF )
				Yn = 0xFF;
			return	(BYTE)Yn;
		}
	}
	return	0x00;
}

void	CDirectDraw::GetZapperPos( LONG& x, LONG& y )
{
	x = y = -1;

	if( !m_lpDD || !m_lpDDPrimary )
		return;
	// ScreenMode changing?
	if( m_bChangeMode )
		return;

	if( !m_bZapper )
		return;

	RECT	rcS, rcC;

	// Size calculate
	rcS.left  = 0;
	rcS.right = SCREEN_WIDTH;
	if( !m_bAllLine ) {
		rcS.top    = 8;
		rcS.bottom = SCREEN_HEIGHT-8;
	} else {
		rcS.top    = 0;
		rcS.bottom = SCREEN_HEIGHT;
	}
	if( m_bScanlineMode ) {
		rcS.top    *= 2;
		rcS.bottom *= 2;
	}

	if( !m_bScreenMode ) {
	// Window mode
		::GetClientRect( m_hWnd, &rcC );
		::ClientToScreen( m_hWnd, (POINT*)&rcC.left );
		::ClientToScreen( m_hWnd, (POINT*)&rcC.right );
	} else {
	// Fullscreen mode
		DDSURFACEDESC2	ddsd;
		ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
		ddsd.dwSize = sizeof(DDSURFACEDESC2);
		m_lpDDPrimary->GetSurfaceDesc(&ddsd);

		if( !m_bMaxZoom ) {
			// Position offset caluclate
			LONG	swidth, sheight;
			LONG	dwidth, dheight;
			LONG	hmul, vmul;

			if( !m_bAspect )  swidth = SCREEN_WIDTH;
			else		  swidth = 320;
			if( !m_bAllLine ) sheight = SCREEN_HEIGHT-16;
			else		  sheight = SCREEN_HEIGHT;

			dwidth  = (LONG)ddsd.dwWidth;
			dheight = (LONG)ddsd.dwHeight;
			hmul = dwidth  / swidth;
			vmul = dheight / sheight;

			if( vmul < hmul ) hmul = vmul;
			else		  vmul = hmul;

			rcC.left   = (dwidth -swidth *hmul)/2;
			rcC.top    = (dheight-sheight*vmul)/2;
			rcC.right  = rcC.left+swidth *hmul;
			rcC.bottom = rcC.top +sheight*vmul;
		} else {
			// Maximum zoom
			rcC.left   = 0;
			rcC.top    = 0;
			rcC.right  = (LONG)ddsd.dwWidth;
			rcC.bottom = (LONG)ddsd.dwHeight;
		}

	}

	POINT	mp;
	FLOAT	hz, vz;

	::GetCursorPos( &mp );

	if( mp.x >= rcC.left && mp.x < rcC.right
	 && mp.y >= rcC.top && mp.y < rcC.bottom ) {
		hz = (FLOAT)(rcS.right-rcS.left)/(rcC.right-rcC.left);
		vz = (FLOAT)(rcS.bottom-rcS.top)/(rcC.bottom-rcC.top);
		x = (LONG)((mp.x-rcC.left)*hz)+rcS.left;
		if( !m_bScanlineMode ) {
			y = (LONG)((mp.y-rcC.top) *vz)+rcS.top;
		} else {
			y = (LONG)(((mp.y-rcC.top) *vz)+rcS.top)/2;
		}

		if( x > SCREEN_WIDTH-1 )
			x = SCREEN_WIDTH-1;
		if( y > SCREEN_HEIGHT-1 )
			y = SCREEN_HEIGHT-1;
	} else {
		x = y = -1;
	}

	m_ZapperPosX = x;
	m_ZapperPosY = y;
}

void	CDirectDraw::SetZapperPos( LONG x, LONG y )
{
	m_ZapperPosX = x;
	m_ZapperPosY = y;
}

// Infomation string
void	CDirectDraw::SetInfoString( LPCSTR str )
{
	if( str ) {
		if( strlen(str) > 32 ) {
			memcpy( m_szInfo, str, 32 );
			m_szInfo[32] = '\0';
		} else {
			strcpy( m_szInfo, str );
		}
	} else {
		m_szInfo[0] = '\0';
	}
}

// Message string
void	CDirectDraw::SetMessageString( LPCSTR str )
{
	if( str ) {
		if( strlen(str) > 32 ) {
			memcpy( m_szMess, str, 32 );
			m_szMess[32] = '\0';
		} else {
			strcpy( m_szMess, str );
		}
		m_bMessage = TRUE;
		m_dwMessageTime = ::timeGetTime();
	} else {
		m_bMessage = FALSE;
		m_szMess[0] = '\0';
	}
}

void	CDirectDraw::RenderString( INT x, INT y, LPCSTR str )
{
	if( !m_lpDD || !m_lpDDPrimary )
		return;
	// ScreenMode changing?
	if( m_bChangeMode )
		return;

	if( !str )
		return;

	RECT	rcS, rcW;
	INT	ch;

	rcW.left   = x;
	rcW.top    = y;
	if( !m_bDoubleSize ) {
		rcW.right  = x+6;
	} else {
		rcW.right  = x+12;
	}
	if( !m_bScanlineMode && !m_bDoubleSize ) {
		rcW.bottom = y+6;
	} else {
		rcW.bottom = y+12;
	}

	while( *str ) {
		ch = toupper(*str)-0x20;
		rcS.left   = (ch%8)*6;
		rcS.right  = rcS.left+6;
		rcS.top    = (ch/8)*6;
		rcS.bottom = rcS.top+6;

		m_lpDDRender->Blt( &rcW, m_lpDDAscii, &rcS, DDBLT_KEYSRC, NULL );

		if( !m_bDoubleSize ) {
			rcW.left  += 6;
			rcW.right += 6;
		} else {
			rcW.left  += 12;
			rcW.right += 12;
		}
		str++;
	}
}

// LZSS BMPからサーフェスへのコピー
void	CDirectDraw::SetLZSSChar( LPBYTE lpLZ, LPDIRECTDRAWSURFACE7 lpDDSurface )
{
LPBYTE	lpBuf = NULL;
LPBITMAPINFOHEADER pbi;
LPBYTE	lpPix;
HDC	hDC;

	if( !(lpBuf = (LPBYTE)malloc( *((LONG*)lpLZ) )) )
		throw	"Out of memory.";

	LZdecode( &lpLZ[8], lpBuf, *((LONG*)&lpLZ[4]) );

	pbi = (LPBITMAPINFOHEADER)lpBuf;
	if( pbi->biBitCount < 16 )
		lpPix = ((LPBYTE)pbi)+pbi->biSize+(1<<pbi->biBitCount)*sizeof(RGBQUAD);
	else
		lpPix = (LPBYTE)pbi+pbi->biSize;

	DDSURFACEDESC2	ddsd;
	ZeroMemory( &ddsd, sizeof(DDSURFACEDESC2) );
	ddsd.dwSize = sizeof(DDSURFACEDESC2);
	lpDDSurface->GetSurfaceDesc(&ddsd);

	// BMPは高さがマイナスな事があるので....
	LONG	height = pbi->biHeight;
	if( height < 0 )
		height = -height;

	if( lpDDSurface->GetDC( &hDC ) == DD_OK ) {
		StretchDIBits( hDC, 0, 0, ddsd.dwWidth, ddsd.dwHeight,
				    0, 0, pbi->biWidth, height, lpPix, (LPBITMAPINFO)lpBuf, DIB_RGB_COLORS, SRCCOPY );
		lpDDSurface->ReleaseDC( hDC );
	}
	FREE( lpBuf );
}

// 8Bit surface
void	CDirectDraw::Render8Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd )
{
	LPBYTE	lpScn = (LPBYTE)ddsd.lpSurface;

	if( !m_bDoubleSize ) {
		if( !m_bScanlineMode ) {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender8_Normal( lpRdr, lpScn, 0x40404040, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPBYTE)lpScn + ddsd.lPitch;
			}
		} else {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender8_Normal( lpRdr, lpScn, 0x40404040, SCREEN_WIDTH/8 );
				lpScn = (LPBYTE)lpScn + ddsd.lPitch;
				LineRender8_Normal( lpRdr, lpScn, 0x80808080, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPBYTE)lpScn + ddsd.lPitch;
			}
		}
	} else {
		if( !m_bScanlineMode ) {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender8_Double( lpRdr, lpScn, 0x40404040, SCREEN_WIDTH/8 );
				lpScn = (LPBYTE)lpScn + ddsd.lPitch;
				LineRender8_Double( lpRdr, lpScn, 0x40404040, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPBYTE)lpScn + ddsd.lPitch;
			}
		} else {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender8_Double( lpRdr, lpScn, 0x40404040, SCREEN_WIDTH/8 );
				lpScn = (LPBYTE)lpScn + ddsd.lPitch;
				LineRender8_Double( lpRdr, lpScn, 0x80808080, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPBYTE)lpScn + ddsd.lPitch;
			}
		}
	}
}

// 16Bit surface
void	CDirectDraw::Render16Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd, LPBYTE lpPal, LPBYTE lpPalScan )
{
	LPWORD	lpScn = (LPWORD)ddsd.lpSurface;

	if( !m_bDoubleSize ) 
	{
		if( !m_bScanlineMode ) 
		{
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) 
			{
				//lpRdrﾊﾇﾃｿﾌ�ﾏﾟﾉﾏｵ羞ﾄﾊﾗﾖｸﾕ�

				LineRender16_Normal( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		} 
		else 
		{
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) 
			{
				//ﾉｨﾃ靠ﾟｷｽﾊｽ｣ｬﾃｿｻｭﾒｻﾐﾐｱｳｾｰｼﾓﾒｻﾐﾐﾉｨﾃ霆ﾒｶﾈﾏﾟ
				LineRender16_Normal( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender16_Normal( lpRdr, lpScn, (LPDWORD)lpPalScan, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		}
	} 
	else 
	{
		if( !m_bScanlineMode ) 
		{
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) 
			{
				LineRender16_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender16_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		} 
		else 
		{
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) 
			{
				LineRender16_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender16_Double( lpRdr, lpScn, (LPDWORD)lpPalScan, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		}
	}
}

// 24Bit surface
void	CDirectDraw::Render24Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd, LPBYTE lpPal, LPBYTE lpPalScan )
{
	LPDWORD	lpScn = (LPDWORD)ddsd.lpSurface;

	if( !m_bDoubleSize ) {
		if( !m_bScanlineMode ) {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender24_Normal( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		} else {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender24_Normal( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender24_Normal( lpRdr, lpScn, (LPDWORD)lpPalScan, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		}
	} else {
		if( !m_bScanlineMode ) {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender24_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender24_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		} else {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender24_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender24_Double( lpRdr, lpScn, (LPDWORD)lpPalScan, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		}
	}
}


// 32Bit surface
void	CDirectDraw::Render32Bit( LPBYTE lpRdr, DDSURFACEDESC2& ddsd, LPBYTE lpPal, LPBYTE lpPalScan )
{
	LPDWORD	lpScn = (LPDWORD)ddsd.lpSurface;

	if( !m_bDoubleSize ) {
		if( !m_bScanlineMode ) {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender32_Normal( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		} else {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender32_Normal( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender32_Normal( lpRdr, lpScn, (LPDWORD)lpPalScan, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		}
	} else {
		if( !m_bScanlineMode ) {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender32_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender32_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		} else {
			for( INT i = 0; i < SCREEN_HEIGHT; i++ ) {
				LineRender32_Double( lpRdr, lpScn, (LPDWORD)lpPal, SCREEN_WIDTH/8 );
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
				LineRender32_Double( lpRdr, lpScn, (LPDWORD)lpPalScan, SCREEN_WIDTH/8 );
				lpRdr += RENDER_WIDTH;
				lpScn = (LPDWORD)((LPBYTE)lpScn + ddsd.lPitch);
			}
		}
	}
}

// ラインレンダラ
// １ライン転送(ノーマル)
void	CDirectDraw::LineRender8_Normal( LPBYTE lpRdr, LPBYTE lpScn, DWORD dwMask, LONG width )
{
	_asm {
		mov	esi, lpRdr
		mov	edi, lpScn
		mov	ecx, width
		mov	edx, dwMask
_Trans8_010:
		mov	eax, [esi+0]
		mov	ebx, [esi+4]
		and	eax, 0x3F3F3F3F
		and	ebx, 0x3F3F3F3F
		or	eax, edx
		or	ebx, edx
		mov	[edi+0], eax
		mov	[edi+4], ebx

		add	esi, 8
		add	edi, 8
		dec	ecx
		jnz	_Trans8_010
	}
}
// １ライン転送(横２倍)
void	CDirectDraw::LineRender8_Double( LPBYTE lpRdr, LPBYTE lpScn, DWORD dwMask, LONG width )
{
	_asm {
		mov	esi, lpRdr
		mov	edi, lpScn
		mov	ecx, width
		mov	edx, dwMask
_Trans8_110:
		movzx	eax, BYTE PTR [esi+0]
		movzx	ebx, BYTE PTR [esi+1]
		and	al, 0x3F
		and	bl, 0x3F
		mov	ah, al
		mov	bh, bl
		shl	ebx, 16
		or	eax, ebx
		or	eax, edx
		mov	[edi+0], eax

		movzx	eax, BYTE PTR [esi+2]
		movzx	ebx, BYTE PTR [esi+3]
		and	al, 0x3F
		and	bl, 0x3F
		mov	ah, al
		mov	bh, bl
		shl	ebx, 16
		or	eax, ebx
		or	eax, edx
		mov	[edi+4], eax

		movzx	eax, BYTE PTR [esi+4]
		movzx	ebx, BYTE PTR [esi+5]
		and	al, 0x3F
		and	bl, 0x3F
		mov	ah, al
		mov	bh, bl
		shl	ebx, 16
		or	eax, ebx
		or	eax, edx
		mov	[edi+8], eax

		movzx	eax, BYTE PTR [esi+6]
		movzx	ebx, BYTE PTR [esi+7]
		and	al, 0x3F
		and	bl, 0x3F
		mov	ah, al
		mov	bh, bl
		shl	ebx, 16
		or	eax, ebx
		or	eax, edx
		mov	[edi+12], eax

		add	esi, 8
		add	edi, 2*8
		dec	ecx
		jnz	_Trans8_110
	}
}

// １ライン転送(ノーマル)
void	CDirectDraw::LineRender16_Normal( LPBYTE lpRdr, LPWORD lpScn, LPDWORD lpPal, LONG width )
{
	_asm {
		mov	eax, lpRdr
		mov	esi, lpScn
		mov	edi, lpPal

		mov	ecx, width
_Trans16_010:
		push	ecx

		movzx	ebx, BYTE PTR [eax+1]
		movzx	edx, WORD PTR [edi+ebx*4]
		shl	edx,16
		movzx	ebx, BYTE PTR [eax+0]
		movzx	ecx, WORD PTR [edi+ebx*4]
		or	edx, ecx
		mov	[esi], edx

		movzx	ebx, BYTE PTR [eax+3]
		movzx	edx, WORD PTR [edi+ebx*4]
		shl	edx,16
		movzx	ebx, BYTE PTR [eax+2]
		movzx	ecx, WORD PTR [edi+ebx*4]
		or	edx, ecx
		mov	[esi+4], edx

		movzx	ebx, BYTE PTR [eax+5]
		movzx	edx, WORD PTR [edi+ebx*4]
		shl	edx,16
		movzx	ebx, BYTE PTR [eax+4]
		movzx	ecx, WORD PTR [edi+ebx*4]
		or	edx, ecx
		mov	[esi+8], edx

		movzx	ebx, BYTE PTR [eax+7]
		movzx	edx, WORD PTR [edi+ebx*4]
		shl	edx,16
		movzx	ebx, BYTE PTR [eax+6]
		movzx	ecx, WORD PTR [edi+ebx*4]
		or	edx, ecx
		mov	[esi+12], edx

		add	eax, 8
		add	esi, 2*8
		pop	ecx
		dec	ecx
		jnz	_Trans16_010
	}
}
// １ライン転送(横２倍)
void	CDirectDraw::LineRender16_Double( LPBYTE lpRdr, LPWORD lpScn, LPDWORD lpPal, LONG width )
{
	_asm {
		mov	eax, lpRdr
		mov	esi, lpScn
		mov	edi, lpPal

		mov	ecx, width
_Trans16_110:
		push	ecx

		movzx	ebx, BYTE PTR [eax+0]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+0], edx

		movzx	ebx, BYTE PTR [eax+1]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+4], edx

		movzx	ebx, BYTE PTR [eax+2]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+8], edx

		movzx	ebx, BYTE PTR [eax+3]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+12], edx

		movzx	ebx, BYTE PTR [eax+4]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+16], edx

		movzx	ebx, BYTE PTR [eax+5]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+20], edx

		movzx	ebx, BYTE PTR [eax+6]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+24], edx

		movzx	ebx, BYTE PTR [eax+7]
		movzx	edx, WORD PTR [edi+ebx*4]
		mov	ecx, edx
		shl	edx,16
		or	edx, ecx
		mov	[esi+28], edx

		add	eax, 8
		add	esi, 4*8
		pop	ecx
		dec	ecx
		jnz	_Trans16_110
	}
}

// １ライン転送(ノーマル)
void	CDirectDraw::LineRender24_Normal( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width )
{
	_asm {
		mov	eax, lpRdr		// レンダリングスクリーン
		mov	esi, lpScn		// サーフェス
		mov	edi, lpPal

		mov	ecx, width		// 8ドット単位
_Trans24_010:
		push	ecx

		movzx	ebx, BYTE PTR [eax+0]
		mov	ecx, [edi+ebx*4]	// 0CBA
		shl	ecx, 8			// CBA0

		movzx	ebx, BYTE PTR [eax+1]
		mov	edx, [edi+ebx*4]	// 0cba
		mov	cl, dl			// CBAa
		ror	ecx, 8
		mov	[esi], ecx

		movzx	ebx, BYTE PTR [eax+2]
		mov	ecx, [edi+ebx*4]	// 0CBA
		shl	edx, 8			// cba0
		mov	dx, cx			// cbBA
		ror	edx, 16			// BAcb
		mov	[esi+4], edx

		movzx	ebx, BYTE PTR [eax+3]
		mov	edx, [edi+ebx*4]	// 0cba
		shl	edx, 8			// cba0
		ror	ecx, 16
		mov	dl, cl			// cbaC
		mov	[esi+8], edx

		movzx	ebx, BYTE PTR [eax+4]
		mov	ecx, [edi+ebx*4]	// 0CBA
		rol	ecx, 8			// CBA0

		movzx	ebx, BYTE PTR [eax+5]
		mov	edx, [edi+ebx*4]	// 0cba
		mov	cl, dl			// CBAa
		ror	ecx, 8
		mov	[esi+12], ecx

		movzx	ebx, BYTE PTR [eax+6]
		mov	ecx, [edi+ebx*4]	// 0CBA
		shl	edx, 8			// cba0
		mov	dx, cx			// cbBA
		ror	edx, 16			// BAcb
		mov	[esi+16], edx

		movzx	ebx, BYTE PTR [eax+7]
		mov	edx, [edi+ebx*4]	// 0cba
		shl	edx, 8			// cba0
		ror	ecx, 16
		mov	dl, cl			// cbaC
		mov	[esi+20], edx

		add	eax, 8
		add	esi, 3*8
		pop	ecx
		dec	ecx
		jnz	_Trans24_010
	}
}
// １ライン転送(横２倍)
void	CDirectDraw::LineRender24_Double( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width )
{
	_asm {
		mov	eax, lpRdr		// レンダリングスクリーン
		mov	esi, lpScn		// サーフェス
		mov	edi, lpPal

		mov	ecx, width		// 6ドット単位(8Byte)
_Trans24_110:
		push	ecx

		movzx	ebx, BYTE PTR [eax+0]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+0], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+4], cx		// BC <- CB

		movzx	ebx, BYTE PTR [eax+1]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+6], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+10], cx		// BC <- CB

		movzx	ebx, BYTE PTR [eax+2]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+12], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+16], cx		// BC <- CB

		movzx	ebx, BYTE PTR [eax+3]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+18], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+22], cx		// BC <- CB

		movzx	ebx, BYTE PTR [eax+4]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+24], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+28], cx		// BC <- CB

		movzx	ebx, BYTE PTR [eax+5]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+30], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+34], cx		// BC <- CB

		movzx	ebx, BYTE PTR [eax+6]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+36], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+40], cx		// BC <- CB

		movzx	ebx, BYTE PTR [eax+7]
		mov	ecx, [edi+ebx*4]	// 0CBA <- ABC0
		mov	edx, ecx		// 0CBA
		rol	edx, 8			// cba0
		mov	dl, cl			// cbaA
		ror	edx, 8			// Acba
		mov	[esi+42], ecx		// abcA <- Acba
		ror	ecx, 8			// A0CB
		mov	[esi+46], cx		// BC <- CB

		add	eax, 8
		add	esi, 6*8
		pop	ecx
		dec	ecx
		jnz	_Trans24_110
	}
}

// １ライン転送(ノーマル)
void	CDirectDraw::LineRender32_Normal( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width )
{
	_asm {
		mov	eax, lpRdr		// レンダリングスクリーン
		mov	esi, lpScn		// サーフェス
		mov	edi, lpPal

		mov	ecx, width		// 8ドット単位
_Trans32_010:
		push	ecx

		movzx	ebx, BYTE PTR [eax+0]
		mov	edx, [edi+ebx*4]
		mov	[esi], edx
		movzx	ebx, BYTE PTR [eax+1]
		mov	edx, [edi+ebx*4]
		mov	[esi+4], edx

		movzx	ebx, BYTE PTR [eax+2]
		mov	edx, [edi+ebx*4]
		mov	[esi+8], edx
		movzx	ebx, BYTE PTR [eax+3]
		mov	edx, [edi+ebx*4]
		mov	[esi+12], edx

		movzx	ebx, BYTE PTR [eax+4]
		mov	edx, [edi+ebx*4]
		mov	[esi+16], edx
		movzx	ebx, BYTE PTR [eax+5]
		mov	edx, [edi+ebx*4]
		mov	[esi+20], edx

		movzx	ebx, BYTE PTR [eax+6]
		mov	edx, [edi+ebx*4]
		mov	[esi+24], edx
		movzx	ebx, BYTE PTR [eax+7]
		mov	edx, [edi+ebx*4]
		mov	[esi+28], edx

		add	eax, 8
		add	esi, 4*8
		pop	ecx
		dec	ecx
		jnz	_Trans32_010
	}
}
// １ライン転送(横２倍)
void	CDirectDraw::LineRender32_Double( LPBYTE lpRdr, LPDWORD lpScn, LPDWORD lpPal, LONG width )
{
	_asm {
		mov	eax, lpRdr		// レンダリングスクリーン
		mov	esi, lpScn		// サーフェス
		mov	edi, lpPal

		mov	ecx, width		// 8ドット単位
_Trans32_110:
		push	ecx

		movzx	ebx, BYTE PTR [eax+0]
		mov	edx, [edi+ebx*4]
		mov	[esi], edx
		mov	[esi+4], edx
		movzx	ebx, BYTE PTR [eax+1]
		mov	edx, [edi+ebx*4]
		mov	[esi+8], edx
		mov	[esi+12], edx

		movzx	ebx, BYTE PTR [eax+2]
		mov	edx, [edi+ebx*4]
		mov	[esi+16], edx
		mov	[esi+20], edx
		movzx	ebx, BYTE PTR [eax+3]
		mov	edx, [edi+ebx*4]
		mov	[esi+24], edx
		mov	[esi+28], edx

		movzx	ebx, BYTE PTR [eax+4]
		mov	edx, [edi+ebx*4]
		mov	[esi+32], edx
		mov	[esi+36], edx
		movzx	ebx, BYTE PTR [eax+5]
		mov	edx, [edi+ebx*4]
		mov	[esi+40], edx
		mov	[esi+44], edx

		movzx	ebx, BYTE PTR [eax+6]
		mov	edx, [edi+ebx*4]
		mov	[esi+48], edx
		mov	[esi+52], edx
		movzx	ebx, BYTE PTR [eax+7]
		mov	edx, [edi+ebx*4]
		mov	[esi+56], edx
		mov	[esi+60], edx

		add	eax, 8
		add	esi, 8*8
		pop	ecx
		dec	ecx
		jnz	_Trans32_110
	}
}



void	CDirectDraw::ViewChars( unsigned char * lpRdr)	//ﾊ鰔ｴｸ｡ｶｯｵﾄﾎﾄｱｾ
{
	//if(!lpDDSurface)
	//	lpDDSurface = m_lpDDRender;
	
	memset(lpRdr, 15, 625);

	HDC m_dc;
	HFONT m_font, oldfont;

	//if(m_lpDDChar)
	{
		int i, j = 0;
		unsigned int pos;
		BOOL istext = 0;
		unsigned char tileid, bktileid=0,count,offset;
		unsigned int gridx, offsetx = (unsigned char)(CPU_MEM_BANK[0][0xD1]);
		unsigned int offsety = (unsigned char)(CPU_MEM_BANK[0][0xD0]);
		unsigned int LATCH_OUT = 0xFE;
		unsigned int LATCH_IN  = 0xFD;
		unsigned int BASE_VRAM = 0x2000+(0x400*(((unsigned char)CPU_MEM_BANK[0][0xD3])&0x03));
		unsigned int NEXT_VRAM = BASE_VRAM+0x400;
		
		if((CPU_MEM_BANK[0][0x61]==0)&&(CPU_MEM_BANK[0][0x62]==0)) 
			istext = 1;
		if(CPU_MEM_BANK[0][0x63]==1)
		{
			LATCH_OUT = 0xFC;
			LATCH_IN  = 0xFB;
			istext = 0;
		}

		gridx = offsetx>>3;

		RGBQUAD	m_Palette[256];
		DirectDraw.GetPaletteData( m_Palette );

		for(j = 0; j < 30 ; j ++)	//scan memory
		{
			count=0;
			offset=0;
			for(i = gridx; i < gridx + 32; i ++)
			{
				pos = (i & 0x1F) + (j << 5);

				if(i < 32)
					tileid = PPU_MEM_BANK[(pos + BASE_VRAM)>>10][pos&0x03FF];
				else
					tileid = PPU_MEM_BANK[(pos + NEXT_VRAM)>>10][pos&0x03FF];
				
				if(tileid == LATCH_OUT)
					istext = 0;

				else if(istext == 1)
				{
					//ｺｺﾗﾖﾏﾔﾊｾ
					if(tileid < MAX_GROUP)
					{
						bktileid = tileid;
						offset+=8;
					}
					else if(tileid < MAX_INDEX)
					{
						//ｷﾃﾎﾊﾊ�ﾐﾔｱ�｣ｬｻ�ｵﾃﾎﾄﾗﾖﾑﾕﾉｫ
						int ataddr;

						if(i < 32)
							ataddr = BASE_VRAM + 0x03C0 + ((j-(j&3))<<1) + (i>>2);
						else
							ataddr = NEXT_VRAM + 0x03C0 + ((j-(j&3))<<1) + ((i-32)>>2);

						unsigned char atbyte = PPU_MEM_BANK[ataddr>>10][ataddr&0x03FF];
						unsigned char color;

						switch(((i-gridx)&3)+((j&3)<<2))
						{
						case 0:
						case 1:
						case 4:
						case 5:
							//bit 0,1
							color = (((atbyte&0x03)<<2)|1);
							break;	
						case 2:
						case 3:
						case 6:
						case 7:
							//bit 2,3
							color = ((atbyte&0x0C)|1);
							break;	
						case 8:
						case 9:
						case 12:
						case 13:
							//bit 4,5
							color = (((atbyte&0x30)>>2)|1);
							break;
						case 10:
						case 11:
						case 14:
						case 15:
							//bit 6,7
							color = (((atbyte&0xC0)>>4)|1);
							break;
						}
						
						DDBLTFX ddbltfx;
						ddbltfx.dwSize = sizeof(DDBLTFX);
						ddbltfx.dwFillColor = 0;
						m_lpDDChar->Blt(NULL, NULL, NULL, DDBLT_COLORFILL|DDBLT_WAIT, &ddbltfx);

						if(m_lpDDChar->GetDC(&m_dc) == DD_OK)
						{
							m_font = CreateFontIndirect(&m_logFont);
							oldfont = (HFONT)SelectObject(m_dc, m_font);

							SetBkColor(m_dc, RGB(0, 0, 0));

							SetTextColor(m_dc, RGB(m_Palette[BGPAL[color]].rgbRed, m_Palette[BGPAL[color]].rgbGreen, m_Palette[BGPAL[color]].rgbBlue));

							HGDIOBJ oldobj = SelectObject(m_dc, GetStockObject(BLACK_BRUSH));
							Rectangle(m_dc, 0, 0, 12, 12);
							SelectObject(m_dc, oldobj);
							TextOut(m_dc, 0, 0, txt[bktileid][tileid], 2);
							int ty = (((j + 30)<<3) - offsety)%240-2;
							int tx = (i<<3) - offsetx+(count<<2)-offset;
							for(int r=0; r<12; r++)
								for(int c=0; c< 12; c++)
								{
									if(GetPixel(m_dc, c, r)!=0)
										lpRdr[((ty+r)<<8)+((ty+r)<<4)+c+tx] = BGPAL[color];
								}
						
							SelectObject(m_dc, oldfont);
							DeleteObject(m_font);

							m_lpDDChar->ReleaseDC( m_dc );
						}
						count++;
					}
					else
					{
						offset=0;
						count=0;
					}
				}
				else if((tileid == LATCH_IN))
				{
					istext = 1;
				}
			}
		}
	}
}

void	CDirectDraw::ViewChars( LPDIRECTDRAWSURFACE7 lpDDSurface )	//ﾊ鰔ｴｸ｡ｶｯｵﾄﾎﾄｱｾ
{
	if(!lpDDSurface)
		lpDDSurface = m_lpDDRender;
	
	HDC m_dc;
	HFONT m_font, oldfont;

	if(m_lpDDChar)
	{
		int i, j = 0;
		unsigned int pos;
		BOOL istext = 0;
		unsigned char tileid, bktileid=0,count,offset;
		unsigned int gridx, offsetx = (unsigned char)(CPU_MEM_BANK[0][0xD1]);
		unsigned int offsety = (unsigned char)(CPU_MEM_BANK[0][0xD0]);
		unsigned int LATCH_OUT = 0xFE;
		unsigned int LATCH_IN  = 0xFD;
		unsigned int BASE_VRAM = 0x2000+(0x400*(((unsigned char)CPU_MEM_BANK[0][0xD3])&0x03));
		unsigned int NEXT_VRAM = BASE_VRAM+0x400;
		
		if((CPU_MEM_BANK[0][0x61]==0)&&(CPU_MEM_BANK[0][0x62]==0)) 
			istext = 1;
		if(CPU_MEM_BANK[0][0x63]==1)
		{
			LATCH_OUT = 0xFC;
			LATCH_IN  = 0xFB;
			istext = 0;
		}

		gridx = offsetx>>3;

		RGBQUAD	m_Palette[256];
		DirectDraw.GetPaletteData( m_Palette );

		for(j = 0; j < 30 ; j ++)	//scan memory
		{
			count=0;
			offset=0;
			for(i = gridx; i < gridx + 32; i ++)
			{
				pos = (i & 0x1F) + (j << 5);

				if(i < 32)
					tileid = PPU_MEM_BANK[(pos + BASE_VRAM)>>10][pos&0x03FF];
				else
					tileid = PPU_MEM_BANK[(pos + NEXT_VRAM)>>10][pos&0x03FF];
				
				if(tileid == LATCH_OUT)
					istext = 0;

				else if(istext == 1)
				{
					//ｺｺﾗﾖﾏﾔﾊｾ
					if(tileid < MAX_GROUP)
					{
						bktileid = tileid;
						offset+=8;
					}
					else if(tileid < MAX_INDEX)
					{
						//ｷﾃﾎﾊﾊ�ﾐﾔｱ�｣ｬｻ�ｵﾃﾎﾄﾗﾖﾑﾕﾉｫ
						int ataddr;

						if(i < 32)
							ataddr = BASE_VRAM + 0x03C0 + ((j-(j&3))<<1) + (i>>2);
						else
							ataddr = NEXT_VRAM + 0x03C0 + ((j-(j&3))<<1) + ((i-32)>>2);

						unsigned char atbyte = PPU_MEM_BANK[ataddr>>10][ataddr&0x03FF];
						unsigned char color;

						switch(((i-gridx)&3)+((j&3)<<2))
						{
						case 0:
						case 1:
						case 4:
						case 5:
							//bit 0,1
							color = (((atbyte&0x03)<<2)|1);
							break;	
						case 2:
						case 3:
						case 6:
						case 7:
							//bit 2,3
							color = ((atbyte&0x0C)|1);
							break;	
						case 8:
						case 9:
						case 12:
						case 13:
							//bit 4,5
							color = (((atbyte&0x30)>>2)|1);
							break;
						case 10:
						case 11:
						case 14:
						case 15:
							//bit 6,7
							color = (((atbyte&0xC0)>>4)|1);
							break;
						}
						
						DDBLTFX ddbltfx;
						ddbltfx.dwSize = sizeof(DDBLTFX);
						ddbltfx.dwFillColor = 0;
						m_lpDDChar->Blt(NULL, NULL, NULL, DDBLT_COLORFILL|DDBLT_WAIT, &ddbltfx);

						if(m_lpDDChar->GetDC(&m_dc) == DD_OK)
						{
							m_font = CreateFontIndirect(&m_logFont);
							oldfont = (HFONT)SelectObject(m_dc, m_font);

							SetBkMode(m_dc, TRANSPARENT);
							SetBkColor(m_dc, RGB(0, 0, 0));
							if(!m_bScanlineMode)
							{
								SetTextColor(m_dc, RGB(m_Palette[BGPAL[color]].rgbRed, m_Palette[BGPAL[color]].rgbGreen, m_Palette[BGPAL[color]].rgbBlue));
							}
							else
							{
								float ratio = (float)m_nScanlineColor/100;
								SetTextColor(m_dc, RGB(m_Palette[BGPAL[color]].rgbRed*ratio, m_Palette[BGPAL[color]].rgbGreen*ratio, m_Palette[BGPAL[color]].rgbBlue*ratio));
							}

							HGDIOBJ oldobj = SelectObject(m_dc, GetStockObject(BLACK_BRUSH));
							Rectangle(m_dc, 0, 0, 12, 12);
							SelectObject(m_dc, oldobj);
							TextOut(m_dc, 0, 0, txt[bktileid][tileid], 2);
						
							SelectObject(m_dc, oldfont);
							DeleteObject(m_font);

							m_lpDDChar->ReleaseDC( m_dc );
						}


						RECT from, to;
						from.left = from.top = 0;
						from.right = from.bottom = 12;

						if( !m_bDoubleSize )
						{
							to.left = (i<<3) - offsetx+(count<<2)-offset;
							to.right = to.left + 12;	
						}
						else
						{
							to.left = (((i<<3) - offsetx+(count<<2)-offset)<<1);
							to.right = to.left + 24;							
						}
						if( !m_bScanlineMode && !m_bDoubleSize )
						{
							to.top = (((j + 30)<<3) - offsety)%240-2;
							to.bottom = to.top + 12;							
						}
						else
						{
							to.top = ((((j + 30)<<3) - offsety)%240 - 2)*2;
							to.bottom = to.top + 24;	
						}

						lpDDSurface->Blt(&to, m_lpDDChar, &from, DDBLT_KEYSRC|DDBLT_WAIT , NULL );

						count++;
					}
					else
					{
						offset=0;
						count=0;
					}
				}
				else if((tileid == LATCH_IN))
				{
					istext = 1;
				}
			}
		}
	}
}

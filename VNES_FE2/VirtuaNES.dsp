# Microsoft Developer Studio Project File - Name="VirtuaNES" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=VirtuaNES - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "VirtuaNES.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "VirtuaNES.mak" CFG="VirtuaNES - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "VirtuaNES - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "VirtuaNES - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE "VirtuaNES - Win32 Profile" (based on "Win32 (x86) Application")
!MESSAGE "VirtuaNES - Win32 Release_Debugout" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "VirtuaNES - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ""
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /I "Zlib" /I "." /I "NES" /I "NES\Mapper" /I "NES\ApuEx" /I "NES\ApuEx\emu2413" /I "NES\PadEx" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x411 /d "NDEBUG"
# ADD RSC /l 0x411 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib winmm.lib comctl32.lib imm32.lib dinput.lib shlwapi.lib /nologo /subsystem:windows /map /machine:I386 /out:"Ê¥»ð»ÕÕÂÍâ´«.exe"
# SUBTRACT LINK32 /nodefaultlib

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ""
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /YX /FD /GZ /c
# ADD CPP /nologo /MT /W3 /Gm /GX /ZI /Od /I "." /I "NES" /I "NES\Mapper" /I "NES\ApuEx" /I "NES\ApuEx\emu2413" /I "NES\PadEx" /I "Zlib" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /YX /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x411 /d "_DEBUG"
# ADD RSC /l 0x411 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib winmm.lib comctl32.lib imm32.lib dinput.lib shlwapi.lib libcmtd.lib /nologo /subsystem:windows /map /debug /machine:I386 /pdbtype:sept
# SUBTRACT LINK32 /nodefaultlib

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Profile"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "VirtuaNES___Win32_Profile"
# PROP BASE Intermediate_Dir "VirtuaNES___Win32_Profile"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ""
# PROP Intermediate_Dir "Profile"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /I "." /I "NES" /I "NES\Mapper" /I "NES\ApuEx" /I "NES\PadEx" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /Z7 /O2 /I "." /I "NES" /I "NES\Mapper" /I "NES\ApuEx" /I "NES\ApuEx\emu2413" /I "NES\PadEx" /I "Zlib" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x411 /d "NDEBUG"
# ADD RSC /l 0x411 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib winmm.lib comctl32.lib imm32.lib dinput.lib /nologo /subsystem:windows /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib winmm.lib comctl32.lib imm32.lib dinput.lib shlwapi.lib /nologo /subsystem:windows /profile /map:"VirtuaNES.map" /debug /machine:I386
# SUBTRACT LINK32 /nodefaultlib

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Release_Debugout"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "VirtuaNES___Win32_Release_Debugout"
# PROP BASE Intermediate_Dir "VirtuaNES___Win32_Release_Debugout"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ""
# PROP Intermediate_Dir "Release_Debugout"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /I "." /I "NES" /I "NES\Mapper" /I "NES\ApuEx" /I "NES\PadEx" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_DEBUGOUT" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /I "." /I "NES" /I "NES\Mapper" /I "NES\ApuEx" /I "NES\ApuEx\emu2413" /I "NES\PadEx" /I "Zlib" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_DEBUGOUT" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x411 /d "NDEBUG"
# ADD RSC /l 0x411 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib winmm.lib comctl32.lib imm32.lib dinput.lib /nologo /subsystem:windows /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib winmm.lib comctl32.lib imm32.lib dinput.lib shlwapi.lib /nologo /subsystem:windows /map /machine:I386
# SUBTRACT LINK32 /debug /nodefaultlib

!ENDIF 

# Begin Target

# Name "VirtuaNES - Win32 Release"
# Name "VirtuaNES - Win32 Debug"
# Name "VirtuaNES - Win32 Profile"
# Name "VirtuaNES - Win32 Release_Debugout"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AboutDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\App.cpp
# End Source File
# Begin Source File

SOURCE=.\Archive.cpp
# End Source File
# Begin Source File

SOURCE=.\CheatDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\Com.cpp
# End Source File
# Begin Source File

SOURCE=.\Config.cpp
# End Source File
# Begin Source File

SOURCE=.\ControllerDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\Crclib.cpp
# End Source File
# Begin Source File

SOURCE=.\DirectDraw.cpp
# End Source File
# Begin Source File

SOURCE=.\DirectInput.cpp
# End Source File
# Begin Source File

SOURCE=.\DirectSound.cpp
# End Source File
# Begin Source File

SOURCE=.\EmulatorDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\EmuThread.cpp
# End Source File
# Begin Source File

SOURCE=.\GameOptionDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\GraphicsDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\MainFrame.cpp
# End Source File
# Begin Source File

SOURCE=.\MMTimer.cpp
# End Source File
# Begin Source File

SOURCE=.\MovieDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\MovieInfoDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\Pathlib.cpp
# End Source File
# Begin Source File

SOURCE=.\Plugin.cpp
# End Source File
# Begin Source File

SOURCE=.\Recent.cpp
# End Source File
# Begin Source File

SOURCE=.\Registry.cpp
# End Source File
# Begin Source File

SOURCE=.\ShortcutDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\SoundDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\VirtuaNES.rc
# End Source File
# Begin Source File

SOURCE=.\WaveRec.cpp
# End Source File
# Begin Source File

SOURCE=.\WinMain.cpp
# End Source File
# Begin Source File

SOURCE=.\Wnd.cpp
# End Source File
# Begin Source File

SOURCE=.\WndHook.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AboutDlg.h
# End Source File
# Begin Source File

SOURCE=.\App.h
# End Source File
# Begin Source File

SOURCE=.\Archive.h
# End Source File
# Begin Source File

SOURCE=.\CheatDlg.h
# End Source File
# Begin Source File

SOURCE=.\Com.h
# End Source File
# Begin Source File

SOURCE=.\Config.h
# End Source File
# Begin Source File

SOURCE=.\ControllerDlg.h
# End Source File
# Begin Source File

SOURCE=.\Crclib.h
# End Source File
# Begin Source File

SOURCE=.\DirectDraw.h
# End Source File
# Begin Source File

SOURCE=.\DirectInput.h
# End Source File
# Begin Source File

SOURCE=.\DirectSound.h
# End Source File
# Begin Source File

SOURCE=.\EmulatorDlg.h
# End Source File
# Begin Source File

SOURCE=.\EmuThread.h
# End Source File
# Begin Source File

SOURCE=.\GameOptionDlg.h
# End Source File
# Begin Source File

SOURCE=.\GraphicsDlg.h
# End Source File
# Begin Source File

SOURCE=.\lzAscii.h
# End Source File
# Begin Source File

SOURCE=.\lzSight.h
# End Source File
# Begin Source File

SOURCE=.\lzTVlayer.h
# End Source File
# Begin Source File

SOURCE=.\Macro.h
# End Source File
# Begin Source File

SOURCE=.\MainFrame.h
# End Source File
# Begin Source File

SOURCE=.\MMTimer.h
# End Source File
# Begin Source File

SOURCE=.\MovieDlg.h
# End Source File
# Begin Source File

SOURCE=.\MovieInfoDlg.h
# End Source File
# Begin Source File

SOURCE=.\Pathlib.h
# End Source File
# Begin Source File

SOURCE=.\Plugin.h
# End Source File
# Begin Source File

SOURCE=.\Recent.h
# End Source File
# Begin Source File

SOURCE=.\Registry.h
# End Source File
# Begin Source File

SOURCE=.\resource.h
# End Source File
# Begin Source File

SOURCE=.\ShortcutDlg.h
# End Source File
# Begin Source File

SOURCE=.\SoundDlg.h
# End Source File
# Begin Source File

SOURCE=.\Typedef.h
# End Source File
# Begin Source File

SOURCE=.\VirtuaNESres.h
# End Source File
# Begin Source File

SOURCE=.\WaveRec.h
# End Source File
# Begin Source File

SOURCE=.\Wnd.h
# End Source File
# Begin Source File

SOURCE=.\WndHook.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\CheatImageList.bmp
# End Source File
# Begin Source File

SOURCE=.\fe2_h3_end.ico
# End Source File
# Begin Source File

SOURCE=.\VirtuaNES.ico
# End Source File
# End Group
# Begin Group "NES"

# PROP Default_Filter ""
# Begin Group "Mapper"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\NES\Mapper\Mapper.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\Mapper\Mapper.h
# End Source File
# Begin Source File

SOURCE=.\NES\Mapper\Mapper010.cpp

!IF  "$(CFG)" == "VirtuaNES - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Profile"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Release_Debugout"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\NES\Mapper\Mapper010.h

!IF  "$(CFG)" == "VirtuaNES - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Profile"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Release_Debugout"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\NES\Mapper\MapperNSF.cpp

!IF  "$(CFG)" == "VirtuaNES - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Profile"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Release_Debugout"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\NES\Mapper\MapperNSF.h

!IF  "$(CFG)" == "VirtuaNES - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Profile"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Release_Debugout"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# End Group
# Begin Group "ApuEx"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\NES\ApuEX\APU_INTERFACE.h
# End Source File
# Begin Source File

SOURCE=.\NES\ApuEX\APU_INTERNAL.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\ApuEX\APU_INTERNAL.h
# End Source File
# End Group
# Begin Source File

SOURCE=.\NES\APU.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\APU.h
# End Source File
# Begin Source File

SOURCE=.\NES\Cheat.h
# End Source File
# Begin Source File

SOURCE=.\NES\Cpu.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\CPU.h
# End Source File
# Begin Source File

SOURCE=.\NES\MMU.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\MMU.h
# End Source File
# Begin Source File

SOURCE=.\NES\Nes.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\Nes.h
# End Source File
# Begin Source File

SOURCE=.\NES\PAD.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\PAD.h
# End Source File
# Begin Source File

SOURCE=.\NES\PPU.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\PPU.h
# End Source File
# Begin Source File

SOURCE=.\NES\ROM.cpp
# End Source File
# Begin Source File

SOURCE=.\NES\ROM.h
# End Source File
# Begin Source File

SOURCE=.\NES\ROM_Patch.cpp

!IF  "$(CFG)" == "VirtuaNES - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Profile"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "VirtuaNES - Win32 Release_Debugout"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\NES\State.h
# End Source File
# End Group
# Begin Source File

SOURCE=.\FE2chn.nes
# End Source File
# Begin Source File

SOURCE=.\fe2jap.nes
# End Source File
# End Target
# End Project

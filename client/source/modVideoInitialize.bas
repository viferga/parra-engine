Attribute VB_Name = "modVideoInitialize"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Copyright (C) 1999-2001 Microsoft Corporation.  All Rights Reserved.
'
'  File:       D3DInit.bas
'  Content:    VB D3DFramework global initialization module
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
' DOC:  Use with
' DOC:        D3DAnimation.cls
' DOC:        D3DFrame.cls
' DOC:        D3DMesh.cls
' DOC:        D3DSelectDevice.frm (optional)
' DOC:
' DOC:  Short list of usefull functions
' DOC:        D3DUtil_Init                  first call to framework
' DOC:        D3DUtil_LoadFromFile          loads an x-file
' DOC:        D3DUtil_SetupDefaultScene     setup a camera lights and materials
' DOC:        D3DUtil_SetupCamera           point camera
' DOC:        D3DUtil_SetupMediaPath        set directory to load textures from
' DOC:        D3DUtil_PresentAll            show graphic on the screen
' DOC:        D3DUtil_ResizeWindowed        resize for windowed modes
' DOC:        D3DUtil_ResizeFullscreen      resize to fullscreen mode
' DOC:        D3DUtil_CreateTextureInPool   create a texture
Option Explicit

'****************************************************************************
'    Parra Engine is a MMORPG Isometric Game Engine.
'    Copyright (C) 2009 - 2013 Vicente Eduardo Ferrer Garcia
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as
'    published by the Free Software Foundation, either version 3 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'****************************************************************************



'------------------------------------------------------------------
'  User Defined types
'------------------------------------------------------------------

' DOC: Type that hold info about a display mode
Public Type D3DUTIL_MODEINFO
    lWidth As Long                                  'Screen width in this mode
    lHeight As Long                                 'Screen height in this mode
    Format As CONST_D3DFORMAT                       'Pixel format in this mode
    VertexBehavior As CONST_D3DCREATEFLAGS          'Whether this mode does SW or HW vertex processing
End Type

' DOC: Type that hold info about a particular back buffer format
Public Type D3DUTIL_FORMATINFO
    Format As CONST_D3DFORMAT
    usage As Long
    bCanDoWindowed As Boolean
    bCanDoFullscreen As Boolean
End Type

Public Type D3DUTIL_DEVTYPEINFO
    'List of compatible Modes for this device
    lNumModes As Long
    Modes() As D3DUTIL_MODEINFO
        
    'List of compatible formats for this device
    lNumFormats As Long
    FormatInfo() As D3DUTIL_FORMATINFO

    lCurrentMode As Long        'index into modes array of the current mode
    
End Type

' DOC: Type that holds info about adapters installed on the system
Public Type D3DUTIL_ADAPTERINFO
    'Device data
    DeviceType As CONST_D3DDEVTYPE                  'Reference, HAL
    d3dcaps As D3DCAPS8                             'Caps of this device
    sDesc As String                                 'Name of this device
    lCanDoWindowed As Long                          'Whether or not this device can work windowed
                
    DevTypeInfo(2) As D3DUTIL_DEVTYPEINFO           'format and display mode list for hal=1 for ref=2
    
    d3dai As D3DADAPTER_IDENTIFIER8
            
    DesktopMode As D3DDISPLAYMODE
    
    'CurrentState
    bWindowed As Boolean        'currently in windowed mode
    bReference As Boolean       'currently using reference rasterizer
End Type

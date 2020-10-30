<#
.SYNOPSIS
Window Monitor
指定したウィンドウのサムネイルを表示します。
.DESCRIPTION
指定したウィンドウのサムネイルを表示します。
ウィンドウの指定は、ウィンドウ内からドラッグを開始し、サムネイル表示したいウィンドウに合わせてリリースするとその時点で指し示すウィンドウが指定されます。
ウィンドウからフォーカスが離れると半透明になり、マウス操作は透過して背景に重なったウィンドウを操作できます。
再度操作したいときは、タスクバーやキー操作でアクティブにすると、フォーカスが離れるまで操作できます。
.NOTES
Window Monitor version 3.00

MIT License

Copyright (c) 2020 Isao Sato

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>


################################################################
# Window Monitor
# Version 3.00
# (C) Isao Sato
################################
# 2016/10/26 first release
################################################################

param([double] $Opacity_Focused = 1.0, [double] $Opacity_Unfocused = 0.4)

[psobject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::Add('marshal',[System.Runtime.InteropServices.Marshal])

function Unregist-Thumbnail()
{
    try
    {
        [marshal]::ThrowExceptionForHR([Win32.DWM]::UnregisterThumbnail($MainForm.hThumbnail))
    }
    catch [System.ArgumentException]
    {
        $MainForm.hThumbnail = [IntPtr]::Zero
    }
}

function Update-ThumbnailProperties([Win32.DWM+THUMBNAIL_PROPERTIES] $ThumbnailProperties)
{
    try
    {
        [marshal]::ThrowExceptionForHR([Win32.DWM]::UpdateThumbnailProperties($MainForm.hThumbnail, $ThumbnailProperties))
    }
    catch [System.ArgumentException]
    {
        $MainForm.hThumbnail = [IntPtr]::Zero
    }
}

function Query-ThumbnailSourceSize()
{
    try
    {
        $size = New-Object Win32.SIZE
        [marshal]::ThrowExceptionForHR([Win32.DWM]::QueryThumbnailSourceSize($MainForm.hThumbnail, [ref] $size))
        Write-Output $size
    }
    catch [System.ArgumentException]
    {
        $MainForm.hThumbnail = [IntPtr]::Zero
    }
}

function Set-Thumbnail($hwnd)
{
    if($MainForm.hThumbnail -ne [IntPtr]::Zero)
    {
        Unregist-Thumbnail
    }
    $hThumbnail = [IntPtr]::Zero
    [marshal]::ThrowExceptionForHR([Win32.DWM]::RegisterThumbnail($MainForm.Handle, $hwnd, [ref] $hThumbnail))
    $MainForm.hThumbnail = $hThumbnail
    $ThumbnailProperties = New-Object Win32.DWM+THUMBNAIL_PROPERTIES
    $ThumbnailProperties.dwFlags `
        =    [Win32.DWM+TNP]::VISIBLE `
        -bor [Win32.DWM+TNP]::OPACITY `
        -bor [Win32.DWM+TNP]::RECTDESTINATION `
        -bor [Win32.DWM+TNP]::SOURCECLIENTAREAONLY
    
    $ThumbnailProperties.opacity = 255
    $ThumbnailProperties.fVisible = $true
    $ThumbnailProperties.rcDestination = New-Object Win32.RECT -Property @{
        left = 0
        top = 0
        right = $MainForm.ClientRectangle.Right
        bottom = $MainForm.ClientRectangle.Bottom
    }
    $ThumbnailProperties.fSourceClientAreaOnly = $true
    
    Update-ThumbnailProperties $ThumbnailProperties
    
    $StringBuffer = New-Object System.Text.StringBuilder 256
    [Win32.User32]::GetWindowText($hwnd, $StringBuffer, $StringBuffer.Capacity+1)
    
    $MainForm.Text = ('縮小ﾓﾆﾀ ({0})' -f $StringBuffer)
}

function Set-ThumbnailSize()
{
    if($MainForm.hThumbnail -ne $null)
    {
        if($MainForm.hThumbnail -ne [IntPtr]::Zero)
        {
            $ThumbnailProperties = New-Object Win32.Dwm+THUMBNAIL_PROPERTIES
            $ThumbnailProperties.dwFlags = [Win32.Dwm+TNP]::RECTDESTINATION
            
            $ThumbnailProperties.rcDestination = New-Object Win32.RECT -Property @{
                left = 0;
                top = 0;
                right = $MainForm.ClientRectangle.Right;
                bottom = $MainForm.ClientRectangle.Bottom
            }
            
            Update-ThumbnailProperties $ThumbnailProperties
        }
    }
}

function Adjust-FormHeight
{
    $sourcesize = (Query-ThumbnailSourceSize)
    $clientsize = $MainForm.ClientSize
    
    $ratio_x = ([double] $clientsize.Width)  / ([double] $sourcesize.cx)
    
    $clientsize.Height = [int] (([double] $sourcesize.cy) *$ratio_x)
    $MainForm.ClientSize = $clientsize
}

function Adjust-FormWidth
{
    $sourcesize = (Query-ThumbnailSourceSize)
    $clientsize = $MainForm.ClientSize
    
    $ratio_y = ([double] $clientsize.Height) / ([double] $sourcesize.cy)
    
    $clientsize.Width = [int] (([double] $sourcesize.cx) *$ratio_y)
    $MainForm.ClientSize = $clientsize
}

function Adjust-FormSize
{
    if($MainForm.hThumbnail -ne [IntPtr]::Zero)
    {
        $sizediff = $MainForm.ClientSize -$MainForm.SizeBeforeResize
        
        if($sizediff.Height -eq 0)
        {
            Adjust-FormHeight
        }
        else
        {
            if($sizediff.Width -eq 0)
            {
                Adjust-FormWidth
            }
            else
            {
                Adjust-FormSize2
            }
        }
    }
}

function Adjust-FormSize2
{
    $sourcesize = (Query-ThumbnailSourceSize)
    $clientsize = $MainForm.ClientSize
    
    $ratio_x = ([double] $clientsize.Width)  / ([double] $sourcesize.cx)
    $ratio_y = ([double] $clientsize.Height) / ([double] $sourcesize.cy)
    
    if($ratio_x -gt $ratio_y)
    {
        $clientsize.Width = [int] (([double] $sourcesize.cx) *$ratio_y)
        $MainForm.ClientSize = $clientsize
    }
    if($ratio_x -lt $ratio_y)
    {
        $clientsize.Height = [int] (([double] $sourcesize.cy) *$ratio_x)
        $MainForm.ClientSize = $clientsize
    }
}

function Set-FormAlpha
{
    if($MainForm.Focused)
    {
        $MainForm.Opacity = $Opacity_Focused
        [Win32.User32]::SetWindowLong($MainForm.Handle, [Win32.User32+GWL]::EXSTYLE, (([Win32.User32]::GetWindowLong($MainForm.Handle, [Win32.User32+GWL]::EXSTYLE) -bor [Win32.User32+WS_EX]::TOPMOST) -band (-bnot [Win32.User32+WS_EX]::TRANSPARENT)))
    }
    else
    {
        $MainForm.Opacity = $Opacity_Unfocused
        [Win32.User32]::SetWindowLong($MainForm.Handle, [Win32.User32+GWL]::EXSTYLE, ([Win32.User32]::GetWindowLong($MainForm.Handle, [Win32.User32+GWL]::EXSTYLE) -bor [Win32.User32+WS_EX]::TRANSPARENT -bor [Win32.User32+WS_EX]::TOPMOST))
    }
}

function Handle-MouseDown([object] $src, [System.Windows.Forms.MouseEventArgs] $e)
{
    $MainForm.Capture = $true
}

function Handle-MouseUp([object] $src, [System.Windows.Forms.MouseEventArgs] $e)
{
    if($MainForm.Capture)
    {
        $MainForm.Capture = $false
        Adjust-FormSize
    }
}

function Handle-MouseMove([object] $src, [System.Windows.Forms.MouseEventArgs] $e)
{
    if($MainForm.Capture)
    {
        $p = $MainForm.PointToScreen($e.Location)
        $hwnd = [Win32.User32]::GetAncestor([Win32.User32]::WindowFromPoint($p), [Win32.User32+GA]::ROOT)
        if(($MainForm.Handle -ne $hwnd) -and ($MainForm.hwndSource -ne $hwnd))
        {
            $MainForm.hwndSource = $hwnd
            Set-Thumbnail $hwnd
        }
    }
}

function Handle-MouseDoubleClick([object] $src, [System.Windows.Forms.MouseEventArgs] $e)
{
}

function Handle-Resize([object] $src, [System.EventArgs] $e)
{
    Set-ThumbnailSize
}

function Handle-ResizeBegin([object] $src, [System.EventArgs] $e)
{
    $MainForm.SizeBeforeResize = $MainForm.ClientSize
}

function Handle-ResizeEnd([object] $src, [System.EventArgs] $e)
{
    Adjust-FormSize
}

function Handle-HandleCreated([object] $src, [System.EventArgs] $e)
{
    Set-FormAlpha
}

function Handle-HandleDestroyed([object] $src, [System.EventArgs] $e)
{
}

function Handle-GotFocus([object] $src, [System.EventArgs] $e)
{
    Set-FormAlpha
}

function Handle-LostFocus([object] $src, [System.EventArgs] $e)
{
    Set-FormAlpha
}

function Handle-Load([object] $src, [System.EventArgs] $e)
{
    [System.Windows.Forms.Form] $f = $src -as [System.Windows.Forms.Form]
    if($f -ne $null)
    {
        $f.Close()
    }
}

function Define-WindowsAPI
{
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    $DefiningTypes = @{}
    $typestack = New-Object System.Collections.Generic.Stack[System.Collections.Hashtable]
    
    $appdomain = [AppDomain]::CurrentDomain
    $asmbuilder = $appdomain.DefineDynamicAssembly((New-Object Reflection.AssemblyName 'Win32'), [Reflection.Emit.AssemblyBuilderAccess]::Run)
    $modbuilder = $asmbuilder.DefineDynamicModule('Win32.dll')
    
    $modbuilder |% {
        $_.DefineType(
            'Win32.DWM',
            [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, Public, BeforeFieldInit',
            [System.Object]
            ) |% {
            $DefiningTypes['Win32.DWM'] = @{}
            $DefiningTypes['Win32.DWM'].Builder = $_
            $_.DefineNestedType(
                'THUMBNAIL_PROPERTIES',
                [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, NestedPublic, SequentialLayout, BeforeFieldInit',
                [System.Object]
                ) |% {
                $DefiningTypes['Win32.DWM+THUMBNAIL_PROPERTIES'] = @{}
                $DefiningTypes['Win32.DWM+THUMBNAIL_PROPERTIES'].Builder = $_
            } | Out-Null
            $_.DefineNestedType(
                'TNP',
                [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, NestedPublic, Sealed',
                [System.Enum]
                ) |% {
                $DefiningTypes['Win32.DWM+TNP'] = @{}
                $DefiningTypes['Win32.DWM+TNP'].Builder = $_
            } | Out-Null
        } | Out-Null
        $_.DefineType(
            'Win32.User32',
            [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, Public, BeforeFieldInit',
            [System.Object]
            ) |% {
            $DefiningTypes['Win32.User32'] = @{}
            $DefiningTypes['Win32.User32'].Builder = $_
            $_.DefineNestedType(
                'GWL',
                [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, NestedPublic, Sealed',
                [System.Enum]
                ) |% {
                $DefiningTypes['Win32.User32+GWL'] = @{}
                $DefiningTypes['Win32.User32+GWL'].Builder = $_
            } | Out-Null
            $_.DefineNestedType(
                'WS_EX',
                [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, NestedPublic, Sealed',
                [System.Enum]
                ) |% {
                $DefiningTypes['Win32.User32+WS_EX'] = @{}
                $DefiningTypes['Win32.User32+WS_EX'].Builder = $_
            } | Out-Null
            $_.DefineNestedType(
                'GA',
                [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, NestedPublic, Sealed',
                [System.Enum]
                ) |% {
                $DefiningTypes['Win32.User32+GA'] = @{}
                $DefiningTypes['Win32.User32+GA'].Builder = $_
            } | Out-Null
        } | Out-Null
        $_.DefineType(
            'Win32.RECT',
            [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, Public, SequentialLayout, Sealed, BeforeFieldInit',
            [System.ValueType]
            ) |% {
            $DefiningTypes['Win32.RECT'] = @{}
            $DefiningTypes['Win32.RECT'].Builder = $_
        } | Out-Null
        $_.DefineType(
            'Win32.SIZE',
            [System.Reflection.TypeAttributes] 'AutoLayout, AnsiClass, Class, Public, SequentialLayout, Sealed, BeforeFieldInit',
            [System.ValueType]
            ) |% {
            $DefiningTypes['Win32.SIZE'] = @{}
            $DefiningTypes['Win32.SIZE'].Builder = $_
        } | Out-Null
    }

    function Create-CustomAttributeBuilder([Reflection.ConstructorInfo] $constructor, [object[]] $arguments, [System.Collections.Hashtable] $attributes)
    {
        $AttributeFields = New-Object Collections.Generic.List[Reflection.FieldInfo]
        $AttributeValues = New-Object Collections.Generic.List[Object]
    
        $attributes.GetEnumerator() |% {
            $AttributeFields.Add($constructor.ReflectedType.GetField($_.Key)) | Out-Null
            $AttributeValues.Add($_.Value) | Out-Null
        }
    
        New-Object Reflection.Emit.CustomAttributeBuilder ($constructor, $arguments, $AttributeFields.ToArray(), $AttributeValues.ToArray())
    }

    function DllImport([string] $libraryname, [System.Collections.Hashtable] $attributes = @{})
    {
        Create-CustomAttributeBuilder ([Runtime.InteropServices.DllImportAttribute].GetConstructor(@([string]))) @($libraryname) $attributes
    }

    function MarshalAs([Runtime.InteropServices.UnmanagedType] $type, [System.Collections.Hashtable] $attributes = @{})
    {
        Create-CustomAttributeBuilder ([Runtime.InteropServices.MarshalAsAttribute].GetConstructor(@([Runtime.InteropServices.UnmanagedType]))) @($type) $attributes
    }

    function Out([System.Collections.Hashtable] $attributes = @{})
    {
        Create-CustomAttributeBuilder ([System.Runtime.InteropServices.OutAttribute].GetConstructor(@([Type]::EmptyTypes))) @() $attributes
    }

    function PreserveSig([System.Collections.Hashtable] $attributes = @{})
    {
        Create-CustomAttributeBuilder ([System.Runtime.InteropServices.PreserveSigAttribute].GetConstructor(@([Type]::EmptyTypes))) @() $attributes
    }

    function Flags([System.Collections.Hashtable] $attributes = @{})
    {
        Create-CustomAttributeBuilder ([System.FlagsAttribute].GetConstructor(@())) @() $attributes
    }

    $DefiningTypes['Win32.DWM'].Builder |% {
        $_.DefineMethod(
            'RegisterThumbnail',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.Int32],
            @([System.IntPtr], [System.IntPtr], [System.IntPtr].MakeByRefType())
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'dwmapi.dll'  @{
                        EntryPoint = 'DwmRegisterThumbnail'
                        PreserveSig = $true
                        CallingConvention = [System.Runtime.InteropServices.CallingConvention]::Winapi
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hwndDestination'
                ) | Out-Null
            $_.DefineParameter(
                2,
                [System.Reflection.ParameterAttributes] 'None',
                'hwndSource'
                ) | Out-Null
            $_.DefineParameter(
                3,
                [System.Reflection.ParameterAttributes] 'Out',
                'phThumbnailId'
                ) |% {
                $_.SetCustomAttribute((Out)) | Out-Null
            } | Out-Null
        } | Out-Null
        $_.DefineMethod(
            'UnregisterThumbnail',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.Int32],
            @([System.IntPtr])
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'dwmapi.dll' @{
                        EntryPoint = 'DwmUnregisterThumbnail' # [System.String]
                        CharSet = 1 # [System.Runtime.InteropServices.CharSet]
                        PreserveSig = $true # [System.Boolean]
                        CallingConvention = 1 # [System.Runtime.InteropServices.CallingConvention]
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hThumbnailId'
                ) | Out-Null
        } | Out-Null
        $_.DefineMethod(
            'UpdateThumbnailProperties',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.Int32],
            @([System.IntPtr], $DefiningTypes['Win32.DWM+THUMBNAIL_PROPERTIES'].Builder)
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'dwmapi.dll' @{
                        EntryPoint = 'DwmUpdateThumbnailProperties'
                        PreserveSig = $true
                        CallingConvention = [System.Runtime.InteropServices.CallingConvention]::Winapi
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hThumbnailId'
                ) | Out-Null
            $_.DefineParameter(
                2,
                [System.Reflection.ParameterAttributes] 'None',
                'ptnProperties'
                ) | Out-Null
        } | Out-Null
        $_.DefineMethod(
            'QueryThumbnailSourceSize',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.Int32],
            @([System.IntPtr], $DefiningTypes['Win32.SIZE'].Builder.MakeByRefType())
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'dwmapi.dll' @{
                        EntryPoint = 'DwmQueryThumbnailSourceSize' # [System.String]
                        CharSet = 1 # [System.Runtime.InteropServices.CharSet]
                        PreserveSig = $true # [System.Boolean]
                        CallingConvention = 1 # [System.Runtime.InteropServices.CallingConvention]
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hThumbnail'
                ) | Out-Null
            $_.DefineParameter(
                2,
                [System.Reflection.ParameterAttributes] 'Out',
                'pSize'
                ) |% {
                $_.SetCustomAttribute((Out)) | Out-Null
            } | Out-Null
        } | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.DWM+THUMBNAIL_PROPERTIES'].Builder |% {
        $_.DefineField(
            'dwFlags',
            $DefiningTypes['Win32.DWM+TNP'].Builder,
            [System.Reflection.FieldAttributes] 'Public, HasFieldMarshal'
            ) |% {
            $_.SetCustomAttribute((MarshalAs ([System.Runtime.InteropServices.UnmanagedType]::U4))) | Out-Null
        } | Out-Null
        $_.DefineField(
            'rcDestination',
            $DefiningTypes['Win32.RECT'].Builder,
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
        $_.DefineField(
            'rcSource',
            $DefiningTypes['Win32.RECT'].Builder,
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
        $_.DefineField(
            'opacity',
            [System.Byte],
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
        $_.DefineField(
            'fVisible',
            [System.Boolean],
            [System.Reflection.FieldAttributes] 'Public, HasFieldMarshal'
            ) |% {
            $_.SetCustomAttribute((MarshalAs ([System.Runtime.InteropServices.UnmanagedType]::Bool))) | Out-Null
        } | Out-Null
        $_.DefineField(
            'fSourceClientAreaOnly',
            [System.Boolean],
            [System.Reflection.FieldAttributes] 'Public, HasFieldMarshal'
            ) |% {
            $_.SetCustomAttribute((MarshalAs ([System.Runtime.InteropServices.UnmanagedType]::Bool))) | Out-Null
        } | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.DWM+TNP'].Builder |% {
        $_.SetCustomAttribute((Flags)) | Out-Null
        $_.DefineField(
            'value__',
            [System.UInt32],
            [System.Reflection.FieldAttributes] 'Public, SpecialName, RTSpecialName'
            ) | Out-Null
        $_.DefineField(
            'RECTDESTINATION',
            $DefiningTypes['Win32.DWM+TNP'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000001'))
        } | Out-Null
        $_.DefineField(
            'RECTSOURCE',
            $DefiningTypes['Win32.DWM+TNP'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000002'))
        } | Out-Null
        $_.DefineField(
            'OPACITY',
            $DefiningTypes['Win32.DWM+TNP'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000004'))
        } | Out-Null
        $_.DefineField(
            'VISIBLE',
            $DefiningTypes['Win32.DWM+TNP'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000008'))
        } | Out-Null
        $_.DefineField(
            'SOURCECLIENTAREAONLY',
            $DefiningTypes['Win32.DWM+TNP'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000010'))
        } | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.User32'].Builder |% {
        $_.DefineMethod(
            'WindowFromPoint',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.IntPtr],
            @([System.Drawing.Point])
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'user32.dll' @{
                        EntryPoint = 'WindowFromPoint'
                        PreserveSig = $true
                        CallingConvention = [System.Runtime.InteropServices.CallingConvention]::Winapi
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'p'
                ) | Out-Null
        } | Out-Null
        $_.DefineMethod(
            'GetWindowText',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.Int32],
            @([System.IntPtr], [System.Text.StringBuilder], [System.Int32])
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'user32.dll' @{
                        EntryPoint = 'GetWindowText'
                        CharSet = [System.Runtime.InteropServices.CharSet]::Auto
                        SetLastError = $true
                        PreserveSig = $true
                        CallingConvention = [System.Runtime.InteropServices.CallingConvention]::Winapi
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hWnd'
                ) | Out-Null
            $_.DefineParameter(
                2,
                [System.Reflection.ParameterAttributes] 'HasFieldMarshal',
                'lpString'
                ) |% {
                $_.SetCustomAttribute((MarshalAs ([System.Runtime.InteropServices.UnmanagedType]::LPTStr))) | Out-Null
            } | Out-Null
            $_.DefineParameter(
                3,
                [System.Reflection.ParameterAttributes] 'None',
                'nMaxCount'
                ) | Out-Null
        } | Out-Null
        $_.DefineMethod(
            'SetWindowLong',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.Int32],
            @([System.IntPtr], $DefiningTypes['Win32.User32+GWL'].Builder, [System.UInt32])
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'user32.dll' @{
                        EntryPoint = 'SetWindowLong'
                        SetLastError = $true
                        PreserveSig = $true
                        CallingConvention = [System.Runtime.InteropServices.CallingConvention]::Winapi
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hWnd'
                ) | Out-Null
            $_.DefineParameter(
                2,
                [System.Reflection.ParameterAttributes] 'HasFieldMarshal',
                'nIndex'
                ) |% {
                $_.SetCustomAttribute((MarshalAs ([System.Runtime.InteropServices.UnmanagedType]::I4))) | Out-Null
            } | Out-Null
            $_.DefineParameter(
                3,
                [System.Reflection.ParameterAttributes] 'None',
                'dwNewLong'
                ) | Out-Null
        } | Out-Null
        $_.DefineMethod(
            'GetWindowLong',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.Int32],
            @([System.IntPtr], $DefiningTypes['Win32.User32+GWL'].Builder)
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'user32.dll' @{
                        EntryPoint = 'GetWindowLong'
                        SetLastError = $true
                        PreserveSig = $true
                        CallingConvention = [System.Runtime.InteropServices.CallingConvention]::Winapi
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hWnd'
                ) | Out-Null
            $_.DefineParameter(
                2,
                [System.Reflection.ParameterAttributes] 'HasFieldMarshal',
                'nIndex'
                ) |% {
                $_.SetCustomAttribute((MarshalAs ([System.Runtime.InteropServices.UnmanagedType]::I4))) | Out-Null
            } | Out-Null
        } | Out-Null
        $_.DefineMethod(
            'GetAncestor',
            [System.Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl',
            [System.IntPtr],
            @([System.IntPtr], $DefiningTypes['Win32.User32+GA'].Builder)
            ) |% {
            $_.SetCustomAttribute(
                (DllImport 'user32.dll' @{
                        EntryPoint = 'GetAncestor'
                        SetLastError = $true
                        PreserveSig = $true
                        CallingConvention = [System.Runtime.InteropServices.CallingConvention]::Winapi
                    })
                ) | Out-Null
            $_.SetCustomAttribute((PreserveSig)) | Out-Null
            $_.DefineParameter(
                1,
                [System.Reflection.ParameterAttributes] 'None',
                'hWnd'
                ) | Out-Null
            $_.DefineParameter(
                2,
                [System.Reflection.ParameterAttributes] 'HasFieldMarshal',
                'flags'
                ) |% {
                $_.SetCustomAttribute((MarshalAs ([System.Runtime.InteropServices.UnmanagedType]::U4)))  | Out-Null
            } | Out-Null
        } | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.User32+GWL'].Builder |% {
        $_.DefineField(
            'value__',
            [System.Int32],
            [System.Reflection.FieldAttributes] 'Public, SpecialName, RTSpecialName'
            ) | Out-Null
        $_.DefineField(
            'WNDPROC',
            $DefiningTypes['Win32.User32+GWL'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(-4)
        } | Out-Null
        $_.DefineField(
            'HINSTANCE',
            $DefiningTypes['Win32.User32+GWL'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(-6)
        } | Out-Null
        $_.DefineField(
            'HWNDPARENT',
            $DefiningTypes['Win32.User32+GWL'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(-8)
        } | Out-Null
        $_.DefineField(
            'STYLE',
            $DefiningTypes['Win32.User32+GWL'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(-16)
        } | Out-Null
        $_.DefineField(
            'EXSTYLE',
            $DefiningTypes['Win32.User32+GWL'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(-20)
        } | Out-Null
        $_.DefineField(
            'USERDATA',
            $DefiningTypes['Win32.User32+GWL'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(-21)
        } | Out-Null
        $_.DefineField(
            'ID',
            $DefiningTypes['Win32.User32+GWL'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(-12)
        } | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.User32+WS_EX'].Builder |% {
        $_.SetCustomAttribute((Flags)) | Out-Null
        $_.DefineField(
            'value__',
            [System.UInt32],
            [System.Reflection.FieldAttributes] 'Public, SpecialName, RTSpecialName'
            ) | Out-Null
        $_.DefineField(
            'TOPMOST',
            $DefiningTypes['Win32.User32+WS_EX'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000008'))
        } | Out-Null
        $_.DefineField(
            'TRANSPARENT',
            $DefiningTypes['Win32.User32+WS_EX'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000020'))
        } | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.User32+GA'].Builder |% {
        $_.DefineField(
            'value__',
            [System.UInt32],
            [System.Reflection.FieldAttributes] 'Public, SpecialName, RTSpecialName'
            ) | Out-Null
        $_.DefineField(
            'ARENT',
            $DefiningTypes['Win32.User32+GA'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000001'))
        } | Out-Null
        $_.DefineField(
            'ROOT',
            $DefiningTypes['Win32.User32+GA'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000002'))
        } | Out-Null
        $_.DefineField(
            'ROOTOWNER',
            $DefiningTypes['Win32.User32+GA'].Builder,
            [System.Reflection.FieldAttributes] 'Public, Static, Literal, HasDefault'
            ) |% {
            $_.SetConstant(([uint32] '0x00000003'))
        } | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.RECT'].Builder |% {
        $_.DefineField(
            'left',
            [System.Int32],
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
        $_.DefineField(
            'top',
            [System.Int32],
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
        $_.DefineField(
            'right',
            [System.Int32],
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
        $_.DefineField(
            'bottom',
            [System.Int32],
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
    } | Out-Null
    $DefiningTypes['Win32.SIZE'].Builder |% {
        $_.DefineField(
            'cx',
            [System.Int32],
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
        $_.DefineField(
            'cy',
            [System.Int32],
            [System.Reflection.FieldAttributes] 'Public'
            ) | Out-Null
    } | Out-Null
    
    @(
        'Win32.DWM+TNP',
        'Win32.User32+GWL',
        'Win32.User32+WS_EX',
        'Win32.User32+GA',
        'Win32.RECT',
        'Win32.SIZE',
        'Win32.DWM',
        'Win32.DWM+THUMBNAIL_PROPERTIES',
        'Win32.User32'
    ) |% {
        [psobject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::Add(
            $_,
            $DefiningTypes[$_].Builder.CreateType()
            )
    }
}

Define-WindowsAPI

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Add_Load(${function:Handle-Load})
[System.Windows.Forms.Application]::Run($MainForm)
$MainForm.Dispose()

$MainForm = New-Object System.Windows.Forms.Form
Add-Member -InputObject $MainForm -Name hThumbnail       -MemberType NoteProperty -Value ([IntPtr]::Zero)
Add-Member -InputObject $MainForm -Name hwndSource       -MemberType NoteProperty -Value ([IntPtr]::Zero)
Add-Member -InputObject $MainForm -Name SizeBeforeResize -MemberType NoteProperty -Value ([System.Drawing.Size]::Empty)
$MainForm.Add_MouseDown(${function:Handle-MouseDown})
$MainForm.Add_MouseMove(${function:Handle-MouseMove})
$MainForm.Add_MouseUp(${function:Handle-MouseUp})
$MainForm.Add_MouseDoubleClick(${function:Handle-MouseDoubleClick})
$MainForm.Add_Resize(${function:Handle-Resize})
$MainForm.Add_ResizeBegin(${function:Handle-ResizeBegin})
$MainForm.Add_ResizeEnd(${function:Handle-ResizeEnd})
$MainForm.Add_HandleCreated(${function:Handle-HandleCreated})
$MainForm.Add_HandleDestroyed(${function:Handle-HandleDestroyed})
$MainForm.Add_GotFocus(${function:Handle-GotFocus})
$MainForm.Add_LostFocus(${function:Handle-LostFocus})
$MainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::SizableToolWindow
$MainForm.TopMost = $true
[System.Windows.Forms.Application]::Run($MainForm)

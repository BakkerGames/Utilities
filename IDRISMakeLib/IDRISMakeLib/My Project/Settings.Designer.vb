﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.7.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(sender As Global.System.Object, e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Arena_AppSettings.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property CallUpgrade() As Boolean
            Get
                Return CType(Me("CallUpgrade"),Boolean)
            End Get
            Set
                Me("CallUpgrade") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Y:\IDRIS\LOCAL\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathLocal() As String
            Get
                Return CType(Me("INIPathLocal"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\SVRTEST2\IDRIS\TEST\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathTest() As String
            Get
                Return CType(Me("INIPathTest"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\SVRTEST2\IDRIS\ACCEPT\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathAccept() As String
            Get
                Return CType(Me("INIPathAccept"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\SVRGIT\VSS\PRODUCTION\IDRIS\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathProd() As String
            Get
                Return CType(Me("INIPathProd"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\SVRREPORT\IDRIS\NEWFIS\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathFIS2() As String
            Get
                Return CType(Me("INIPathFIS2"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\DHFILE\H_*\\IDRIS\LOCAL\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathLocalAlt() As String
            Get
                Return CType(Me("INIPathLocalAlt"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Arena_AppSettings.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("C:\Temp")>  _
        Public Property LocalCompilePath() As String
            Get
                Return CType(Me("LocalCompilePath"),String)
            End Get
            Set
                Me("LocalCompilePath") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Arena_AppSettings.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property LastEnv() As String
            Get
                Return CType(Me("LastEnv"),String)
            End Get
            Set
                Me("LastEnv") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Arena_AppSettings.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property LastVolume() As String
            Get
                Return CType(Me("LastVolume"),String)
            End Get
            Set
                Me("LastVolume") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Arena_AppSettings.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property LastLibrary() As String
            Get
                Return CType(Me("LastLibrary"),String)
            End Get
            Set
                Me("LastLibrary") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\SVRTEST2\IDRIS\EOY\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathEOY() As String
            Get
                Return CType(Me("INIPathEOY"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property INIPathLocalAlt2() As String
            Get
                Return CType(Me("INIPathLocalAlt2"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\SVRREPORTTEST\IDRIS\NEWFIS\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathFISTest() As String
            Get
                Return CType(Me("INIPathFISTest"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("C:\IDRIS\LOCAL\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathPC() As String
            Get
                Return CType(Me("INIPathPC"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("D:\IDRIS\LOCAL\PROGRAMS\BIN\CONNECT.INI")>  _
        Public ReadOnly Property INIPathPCD() As String
            Get
                Return CType(Me("INIPathPCD"),String)
            End Get
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.IDRISMakeLib.My.MySettings
            Get
                Return Global.IDRISMakeLib.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace

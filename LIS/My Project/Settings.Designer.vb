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
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "14.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs)
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
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property Server() As String
            Get
                Return CType(Me("Server"),String)
            End Get
            Set
                Me("Server") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property UID() As String
            Get
                Return CType(Me("UID"),String)
            End Get
            Set
                Me("UID") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property PWD() As String
            Get
                Return CType(Me("PWD"),String)
            End Get
            Set
                Me("PWD") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property DBName() As String
            Get
                Return CType(Me("DBName"),String)
            End Get
            Set
                Me("DBName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property MyConnectionString() As String
            Get
                Return CType(Me("MyConnectionString"),String)
            End Get
            Set
                Me("MyConnectionString") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property LicenseKey() As String
            Get
                Return CType(Me("LicenseKey"),String)
            End Get
            Set
                Me("LicenseKey") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property ActivationStatus() As String
            Get
                Return CType(Me("ActivationStatus"),String)
            End Get
            Set
                Me("ActivationStatus") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property DefaultPrinter() As String
            Get
                Return CType(Me("DefaultPrinter"),String)
            End Get
            Set
                Me("DefaultPrinter") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public Property MedTech() As Integer
            Get
                Return CType(Me("MedTech"),Integer)
            End Get
            Set
                Me("MedTech") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property HL7Location() As String
            Get
                Return CType(Me("HL7Location"),String)
            End Get
            Set
                Me("HL7Location") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property HL7Destination() As String
            Get
                Return CType(Me("HL7Destination"),String)
            End Get
            Set
                Me("HL7Destination") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property HL7Read() As Boolean
            Get
                Return CType(Me("HL7Read"),Boolean)
            End Get
            Set
                Me("HL7Read") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property HL7Write() As Boolean
            Get
                Return CType(Me("HL7Write"),Boolean)
            End Get
            Set
                Me("HL7Write") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property PrintBarcode() As Boolean
            Get
                Return CType(Me("PrintBarcode"),Boolean)
            End Get
            Set
                Me("PrintBarcode") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property PDFLocation() As String
            Get
                Return CType(Me("PDFLocation"),String)
            End Get
            Set
                Me("PDFLocation") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property SaveAsPDF() As Boolean
            Get
                Return CType(Me("SaveAsPDF"),Boolean)
            End Get
            Set
                Me("SaveAsPDF") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property BCPrinterName() As String
            Get
                Return CType(Me("BCPrinterName"),String)
            End Get
            Set
                Me("BCPrinterName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property LISType() As String
            Get
                Return CType(Me("LISType"),String)
            End Get
            Set
                Me("LISType") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("server=localhost;User Id=root;PWD=;database=lis_db")>  _
        Public ReadOnly Property db_sbsi_lis_universalConnectionString() As String
            Get
                Return CType(Me("db_sbsi_lis_universalConnectionString"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("server=127.0.0.1;user id=root;password=;database=lis_db")>  _
        Public ReadOnly Property db_lisConnectionString() As String
            Get
                Return CType(Me("db_lisConnectionString"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property AuthenticateRelease() As Boolean
            Get
                Return CType(Me("AuthenticateRelease"),Boolean)
            End Get
            Set
                Me("AuthenticateRelease") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("3306")>  _
        Public Property TCPPort() As Integer
            Get
                Return CType(Me("TCPPort"),Integer)
            End Get
            Set
                Me("TCPPort") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("200")>  _
        Public Property BCWidth() As Integer
            Get
                Return CType(Me("BCWidth"),Integer)
            End Get
            Set
                Me("BCWidth") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("100")>  _
        Public Property BCHeight() As Integer
            Get
                Return CType(Me("BCHeight"),Integer)
            End Get
            Set
                Me("BCHeight") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("8.5")>  _
        Public Property PaperWidth() As Double
            Get
                Return CType(Me("PaperWidth"),Double)
            End Get
            Set
                Me("PaperWidth") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("11")>  _
        Public Property PaperHeight() As Double
            Get
                Return CType(Me("PaperHeight"),Double)
            End Get
            Set
                Me("PaperHeight") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property MyConnectionStringSQL() As String
            Get
                Return CType(Me("MyConnectionStringSQL"),String)
            End Get
            Set
                Me("MyConnectionStringSQL") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property SQLServer() As String
            Get
                Return CType(Me("SQLServer"),String)
            End Get
            Set
                Me("SQLServer") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property SQLUID() As String
            Get
                Return CType(Me("SQLUID"),String)
            End Get
            Set
                Me("SQLUID") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property SQLPWD() As String
            Get
                Return CType(Me("SQLPWD"),String)
            End Get
            Set
                Me("SQLPWD") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property SQLDBName() As String
            Get
                Return CType(Me("SQLDBName"),String)
            End Get
            Set
                Me("SQLDBName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property SQLConnection() As Boolean
            Get
                Return CType(Me("SQLConnection"),Boolean)
            End Get
            Set
                Me("SQLConnection") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property AutoEmail() As Boolean
            Get
                Return CType(Me("AutoEmail"),Boolean)
            End Get
            Set
                Me("AutoEmail") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property XMLDestination() As String
            Get
                Return CType(Me("XMLDestination"),String)
            End Get
            Set
                Me("XMLDestination") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Configuration.SettingsProviderAttribute(GetType(Bluegrams.Application.PortableSettingsProvider)),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property XMLWrite() As Boolean
            Get
                Return CType(Me("XMLWrite"),Boolean)
            End Get
            Set
                Me("XMLWrite") = value
            End Set
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.LIS.My.MySettings
            Get
                Return Global.LIS.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
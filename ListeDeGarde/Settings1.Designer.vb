﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.18408
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On



<Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
 Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0")>  _
Partial Public NotInheritable Class Settings1
    Inherits Global.System.Configuration.ApplicationSettingsBase
    
    Private Shared defaultInstance As Settings1 = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New Settings1()),Settings1)
    
    Public Shared ReadOnly Property [Default]() As Settings1
        Get
            Return defaultInstance
        End Get
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("""C:\Users\Martin\Documents\Scheduling Mira\ListesDeGarde.accdb"""),  _
     Global.System.Configuration.SettingsManageabilityAttribute(Global.System.Configuration.SettingsManageability.Roaming)>  _
    Public Property DataBaseLocation() As String
        Get
            Return CType(Me("DataBaseLocation"),String)
        End Get
        Set
            Me("DataBaseLocation") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("""ayayaya""")>  _
    Public Property Setting2() As String
        Get
            Return CType(Me("Setting2"),String)
        End Get
        Set
            Me("Setting2") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
    Public Property Setting4() As Long
        Get
            Return CType(Me("Setting4"),Long)
        End Get
        Set
            Me("Setting4") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("<?xml version=""1.0"" encoding=""utf-16""?>"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"<ArrayOfString xmlns:xsi=""http://www.w3."& _ 
        "org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  <s"& _ 
        "tring>""C:\Users\Martin\Documents\Scheduling Mira\ListesDeGarde.accdb""</string>"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)& _ 
        "</ArrayOfString>")>  _
    Public Property Setting5() As Global.System.Collections.Specialized.StringCollection
        Get
            Return CType(Me("Setting5"),Global.System.Collections.Specialized.StringCollection)
        End Get
        Set
            Me("Setting5") = value
        End Set
    End Property
End Class

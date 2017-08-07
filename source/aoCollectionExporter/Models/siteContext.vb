
Option Explicit On
Option Strict On

Imports Contensive.BaseClasses

Namespace Contensive.Addons.aoCollectionExporter
    ''' <summary>
    ''' object injected with common site properties 
    ''' </summary>
    Public Class siteContextClass
        Public Property cp As CPBaseClass = Nothing
        Public Property isVersion5 As Boolean = False
        Public Sub New(cp As CPBaseClass)
            Me.cp = cp
        End Sub
    End Class
End Namespace

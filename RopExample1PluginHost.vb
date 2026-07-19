Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Topomatic.ApplicationPlatform
Namespace RopExample1
    Public Class RopExample1PluginHost
        Inherits Plugins.PluginHostInitializator
        Protected Overrides Function GetTypes() As Type()
            Return New Type() {GetType(RopExample1Module)}
        End Function

    End Class
End Namespace


Imports System.ComponentModel
Imports CivilEnginStructures.BeamI
Imports CivilEnginStructures.Bridges
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.Operation.Buffer
Imports Newtonsoft.Json.Linq
Imports Topomatic.Alg
Imports Topomatic.Alg.Road.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime
Public Class BoundaryStructure
    ' Поля класса
    Private _name As String
    ' Конструктор по умолчанию
    Public Sub New()
        _name = ""
    End Sub
End Class

''' <summary>
''' This class contains a collection of the units used in CIS/2.  The collection is keyed by the name of the unit.  The values are to be multiplied
''' into the measure so that the resulting value is in SP3D native measure (e.g. SP3D measures in metres natively, so it converts everything to metres)
''' </summary>
''' <remarks></remarks>
Public Class measures
    Public unitMeasures As New Collection

    Public Sub New()
        unitMeasures.Add(0.001, "millimetre")
        unitMeasures.Add(0.01, "centimetre")
        unitMeasures.Add(1, "metre")
        unitMeasures.Add(0.0254, "inch")
        unitMeasures.Add(0.3048, "foot")
        unitMeasures.Add(0.9144, "yard")
    End Sub
End Class

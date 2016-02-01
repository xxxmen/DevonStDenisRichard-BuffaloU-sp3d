Imports Ingr.SP3D.Common.Middle
Imports System.Security.Cryptography
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Ingr.SP3D.Common.Middle.Services

Module SP3DUtilities

    ''' <summary>
    ''' Rotates an SP3D object via its 4x4 matrix
    ''' </summary>
    ''' <param name="objectToRotate"></param>
    ''' <param name="RadianAngleToRotate"></param>
    ''' <param name="axisToRotateAbout"></param>
    ''' <remarks></remarks>
    Public Sub rotateObject(ByRef objectToRotate As Object, ByVal RadianAngleToRotate As Double, ByVal axisToRotateAbout As String)
        Dim oT1 As New Matrix4X4
        Dim oT2 As New Matrix4X4
        Dim oT3 As New Matrix4X4
        Dim oT As New Matrix4X4

        Select Case UCase(axisToRotateAbout)
            Case "X"
                oT1.Rotate(RadianAngleToRotate, New Vector(1, 0, 0))
                oT2.Translate(New Vector(0, 0, 0))
                oT3.Translate(New Vector(0, 0, 0))
            Case "Y"
                oT1.Translate(New Vector(0, 0, 0))
                oT2.Rotate(RadianAngleToRotate, New Vector(0, 1, 0))
                oT3.Translate(New Vector(0, 0, 0))
            Case "Z"
                oT1.Translate(New Vector(0, 0, 0))
                oT2.Translate(New Vector(0, 0, 0))
                oT3.Rotate(RadianAngleToRotate, New Vector(0, 0, 1))
        End Select

        oT.MultiplyMatrix(oT3)
        oT.MultiplyMatrix(oT2)
        oT.MultiplyMatrix(oT1)

        Dim oo As ITransform = objectToRotate

        oo.Transform(oT)
    End Sub

    ''' <summary>
    ''' Moves an SP3D object via its 4x4 matrix
    ''' </summary>
    ''' <param name="objectToMove"></param>
    ''' <param name="moveToPosition"></param>
    ''' <remarks></remarks>
    Public Sub moveObject(ByRef objectToMove As Object, ByVal moveToPosition As Position)
        Dim oT1 As New Matrix4X4
        Dim oT2 As New Matrix4X4
        Dim oT3 As New Matrix4X4
        Dim oT As New Matrix4X4

        oT1.Translate(New Vector(moveToPosition.X, 0, 0))
        oT2.Translate(New Vector(0, moveToPosition.Y, 0))
        oT3.Translate(New Vector(0, 0, moveToPosition.Z))

        oT.MultiplyMatrix(oT3)
        oT.MultiplyMatrix(oT2)
        oT.MultiplyMatrix(oT1)

        Dim oo As ITransform = objectToMove

        oo.Transform(oT)
    End Sub

End Module

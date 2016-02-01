Imports Ingr.SP3D.Common.Client

Public Class importBuild

    Inherits BaseModalCommand
    Private m_ofrmAttsLab As importBuildForm

    Public Overrides Sub OnStart(ByVal commandID As Integer, ByVal argument As Object)
        MyBase.OnStart(commandID, argument)
        m_ofrmAttsLab = New importBuildForm
        m_ofrmAttsLab.ShowDialog()
    End Sub

    Public Overrides Sub OnStop()
        MyBase.OnStop()
        m_ofrmAttsLab.Close()
        m_ofrmAttsLab = Nothing
    End Sub
End Class

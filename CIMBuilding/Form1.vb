Imports System.Windows.Forms
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Client.Services

Public Class Form1
    Dim result As Boolean

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim collectionCollection As Collection
        Dim directives As Collection            'Existant in the end
        Dim topLevelDirectives As Collection    'Existant in the end
        Dim itemAssignment As Collection
        Dim finalStandards As Collection        'Existant in the end
        Dim globalContexts As Collection
        Dim globalUnit As Double

        collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\Sample_CIS2_Files\Sample_CIS2_Files\U04_Test_STR.stp")
        'collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\Sample_CIS2_Files\Sample_CIS2_Files\Bentley1.stp")
        'collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\Sample_CIS2_Files\Sample_CIS2_Files\SmartPlant3D.stp")
        'collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\Sample_CIS2_Files\Sample_CIS2_Files\ETABS.stp")
        'collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\Sample_CIS2_Files\Sample_CIS2_Files\STAAD1.stp")
        'collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\Sample_CIS2_Files\Sample_CIS2_Files\AISC_Sculpture_J.stp")
        'collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\Sample_CIS2_Files\Sample_CIS2_Files\Tekla1.stp")
        'collectionCollection = loadDirectives("C:\Documents and Settings\stded2\Desktop\CIMBuilding\special export3.stp")


        'Process the collections into individual collections.
        directives = collectionCollection("directives")
        topLevelDirectives = collectionCollection("topLevelDirectives")
        itemAssignment = collectionCollection("itemAssignment")
        globalContexts = collectionCollection("globalContexts")

        'Create the standards collections
        'finalStandards = makeStandards2(itemAssignment,  directives)
        globalUnit = getGlobalContext(globalContexts, directives)

        directives.Add(finalStandards, "FinalStandards")            ' Append the standards to the end of the directive collection -- keeps them together
        collectionCollection = Nothing
        itemAssignment = Nothing
        globalContexts = Nothing

        '---------------------------------------- VISUAL TREE BUILDER --------------------------------
        Dim c As Integer
        TextBox1.Text = 1
        For c = 1 To directives.Count - 1
            If InStr(directives(c), "#") = 0 Then
                DataGridView1.Rows.Add(directives(c))
            Else
                Dim ff As Object

                ff = TreeView1.Nodes.Add("#" & c, directives(c))

                buildHeirarchyTree(directives, directives(c), ff, 1)
            End If
        Next c

        '---------------------------------------- TESTING --------------------------------------------

    End Sub

    '************************* TREE STUFF -- For visual representation of the data
    Private Sub buildHeirarchyTree(ByVal d As Collection, ByVal branch As String, ByRef branchMem As Object, ByVal h As Integer)
        Dim tc As Collection = getSubDirectives(branch)
        Dim c2 As Integer
        On Error Resume Next
        If CInt(h) > CInt(TextBox1.Text) Then
            TextBox1.Text = h
        End If

        For c2 = 1 To tc.Count
            Dim newMem As Object
            newMem = branchMem.nodes.add(tc(c2) & "--" & d("#" & tc(c2)))

            If InStr(d("#" & tc(c2)), "#") Then
                buildHeirarchyTree(d, d("#" & tc(c2)), newMem, h + 1)
            End If
        Next c2

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        TreeView1.ExpandAll()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim c As Integer
        Dim temp As String
        Dim stext As String = TextBox2.Text.ToUpper
        Dim n As TreeNode

        result = False

        For c = 0 To TreeView1.Nodes.Count - 1
            temp = UCase(TreeView1.Nodes(c).Text)

            If InStr(temp.ToUpper, stext) > 0 Then

                TextBox3.Text = c

                n = TreeView1.Nodes(c)
                TreeView1.SelectedNode = n
                TreeView1.Focus()
                Exit For
            Else
                checkChilluns(TreeView1.Nodes(c), stext)

                If result = True Then
                    TextBox3.Text = c

                    n = TreeView1.Nodes(c)
                    TreeView1.Nodes(c).ExpandAll()
                    TreeView1.SelectedNode = n
                    TreeView1.Focus()

                    Exit For
                End If
            End If
        Next c

        If c = TreeView1.Nodes.Count Then
            MsgBox("Not Found")
        End If
    End Sub

    Public Function checkChilluns(ByVal nodeToCheck As TreeNode, ByVal stext As String) As Boolean
        Dim c As Integer
        Dim n As TreeNode
        Dim temp As String

        For c = 0 To nodeToCheck.Nodes.Count - 1
            temp = UCase(nodeToCheck.Nodes(c).Text)

            If InStr(temp.ToUpper, stext) > 0 Then
                result = True
                Exit For
            Else
                If nodeToCheck.Nodes(c).Nodes.Count > 0 Then
                    checkChilluns(nodeToCheck.Nodes(c), stext)
                End If
            End If
        Next

        checkChilluns = result
    End Function

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        If TextBox3.Text = "" Then
            Button3_Click(sender, e)
        Else
            Dim c As Integer
            Dim temp As String
            Dim stext As String = TextBox2.Text.ToUpper
            Dim n As TreeNode

            result = False

            For c = CInt(TextBox3.Text) + 1 To TreeView1.Nodes.Count - 1
                temp = UCase(TreeView1.Nodes(c).Text)

                If InStr(temp.ToUpper, stext) > 0 Then

                    TextBox3.Text = c

                    n = TreeView1.Nodes(c)
                    TreeView1.SelectedNode = n
                    TreeView1.Focus()
                    Exit For
                Else
                    checkChilluns(TreeView1.Nodes(c), stext)

                    If result = True Then
                        TextBox3.Text = c

                        n = TreeView1.Nodes(c)
                        TreeView1.Nodes(c).ExpandAll()
                        TreeView1.SelectedNode = n
                        TreeView1.Focus()

                        Exit For
                    End If
                End If

                If c = TreeView1.Nodes.Count - 1 Then
                    c = 0
                End If

                If c = TextBox3.Text Then
                    MsgBox("Not Found")

                    Exit For
                End If
            Next c

        End If
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click


        'MsgBox(analyze(TextBox4.Text).GetType.FullName)    ''''''' Reference line
        'Dim c As Integer
        'Dim s As String
        'Dim f() As String = getbracketElements("POSITIVE_LENGTH_MEASURE_WITH_UNIT(POSITIVE_LENGTH_MEASURE(7620.),#762")

        'For c = 0 To f.Count - 1
        '    s = s & f(c) & "||"
        'Next c

        's = s & f.Count

        'MsgBox(s)

        'MsgBox(advanceToNextWord("GLOBAL_UNIT_ASSIGNED_CONTEXT((#758))MATERIAL_PROPERTY_CONTEXT()MATERIAL_PROPERTY_CONTEXT_DIMENSIONAL(0.,9999999.)"))
        'MsgBox(advanceToNextWord("NAMED_UNIT(*)PRESSURE_UNIT()SI_UNIT(.KILO.,.PASCAL.))"))

    End Sub
End Class


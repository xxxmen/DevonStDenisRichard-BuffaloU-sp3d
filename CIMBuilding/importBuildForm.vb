Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Windows.Forms
Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Space.Middle
Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.Support.Middle.Services
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Route.Middle.PathFeatureObjectTypes
Imports Ingr.SP3D.Route.Middle.PathFeatureFunctions
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.Structure.Exceptions
Imports Ingr.SP3D.Grids.Middle
Imports Ingr.SP3D.Grids.Middle.GridElevationPlane
Imports Ingr.SP3D.ReferenceData.Middle.Services.CatalogStructHelper


Public Class importBuildForm
    Private m_oSelectSet As SelectSet
    Private m_oObj As BusinessObject
    Private m_oIntfInfo As InterfaceInformation
    Private m_oPropInfo As PropertyInformation

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim fileLocation As New OpenFileDialog

        fileLocation.DefaultExt = "CIM/2 Step files (*.stp)|*.stp"
        fileLocation.Filter = "CIM/2 Step files (*.stp)|*.stp|All files (*.*)|*.*"

        If fileLocation.ShowDialog() = 1 Then
            TextBox1.Text = fileLocation.FileName

            Dim objReader As New System.IO.StreamReader(fileLocation.FileName)
            Dim start As Boolean = False
            Dim tempString As String

            While start = False And Not objReader.EndOfStream
                tempString = objReader.ReadLine

                If InStr(tempString, "originating_system") > 0 Then
                    Exit While
                End If

                If tempString = "DATA;" Then
                    tempString = objReader.ReadLine
                    Exit While
                End If
            End While

            If InStr(LCase(tempString), "tekla") Then
                ComboBox1.SelectedItem = "Tekla"
            ElseIf InStr(LCase(tempString), "staad") Then
                ComboBox1.SelectedItem = "STAAD"
            ElseIf InStr(LCase(tempString), "smartplant") Then
                ComboBox1.SelectedItem = "SmartPlant 3D"
            ElseIf InStr(LCase(tempString), "triforma") Then
                ComboBox1.SelectedItem = "TriForma"
            Else
                ComboBox1.SelectedItem = Nothing
            End If
        End If
    End Sub

    Private Sub Import_Click(sender As System.Object, e As System.EventArgs) Handles Import.Click


        If ComboBox1.SelectedItem.ToString = "" Then
            MsgBox("Please Select the Program that Created the file")
        ElseIf TextBox1.Text <> "" And System.IO.File.Exists(TextBox1.Text) Then
            Dim collectionCollection As Collection
            Dim directives As Collection            'Existant in the end
            Dim topLevelDirectives As Collection    'Existant in the end
            Dim itemAssignment As Collection
            Dim finalStandards As Collection        'Existant in the end
            Dim globalContexts As Collection
            Dim sectionProfiles As Collection
            Dim globalUnit As Double
            Dim SP3DColumnObjects As New Collection             'This will hold the created member objects
            Dim SP3DBeamObjects As New Collection             'This will hold the created member objects
            Dim SP3DBraceObjects As New Collection             'This will hold the created member objects
            Dim c As Integer
            Dim oPos1 As New Position(0, 0, 0)
            Dim oModel As Model = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel
            Dim oRootSys As BusinessObject = oModel.RootSystem
            Dim oCatStruHlpr As New CatalogStructHelper(oModel.PlantCatalog)

            If CheckBox3.Checked = True Then
                oPos1.X = oPos1.X + CDbl(TextBox3.Text)
                oPos1.Y = oPos1.Y + CDbl(TextBox4.Text)
                oPos1.Z = oPos1.Z + CDbl(TextBox5.Text)
            End If

            collectionCollection = loadDirectives(TextBox1.Text)

            'Process the collections into individual collections.
            directives = collectionCollection("directives")
            topLevelDirectives = collectionCollection("topLevelDirectives")
            itemAssignment = collectionCollection("itemAssignment")
            globalContexts = collectionCollection("globalContexts")
            sectionProfiles = collectionCollection("sectionProfiles")

            'If itemAssignment.Count = 0 Then
            '    finalStandards = makeImproperStandards(sectionProfiles, directives)
            'Else
            'finalStandards = makeStandards(itemAssignment, directives)
            'End If

            'MsgBox(itemAssignment(1) & vbCrLf & sectionProfiles(1))

            If collectionCount(itemAssignment) > 0 Then         ' Properly specced steel standard
                finalStandards = makeStandards2(itemAssignment, directives, ComboBox1.SelectedItem.ToString)
            Else
                finalStandards = makeStandards2(sectionProfiles, directives, ComboBox1.SelectedItem.ToString, 2)
            End If

            'Create the standards collections

            globalUnit = getGlobalContext(globalContexts, directives)   ' Global unit processor -- used for coordinate systems

            directives.Add(finalStandards, "FinalStandards")            ' Append the standards to the end of the directive collection -- keeps them together
            collectionCollection = Nothing
            itemAssignment = Nothing
            globalContexts = Nothing
            sectionProfiles = Nothing

            '--------------------------- File/Directives Loaded -------------------------------------------
            Try
                Dim structuralObjects As New Collection

                For c = 1 To topLevelDirectives.Count
                    Dim temp As Object

                    temp = analyze(directives(topLevelDirectives(c)), directives)

                    If temp IsNot Nothing Then
                        structuralObjects.Add(temp)
                    End If
                Next c

                '--------------------------- Objects Created -------------------------------------------

                ' Create a system and sub systems
                Dim oGenericSys As New GenericSystem(oRootSys)
                Dim sysch As New StructuralSystem(oGenericSys)
                Dim beams As New StructuralSystem(sysch)
                Dim columns As New StructuralSystem(sysch)
                Dim braces As New StructuralSystem(sysch)
                Dim coordSystem As New StructuralSystem(sysch)
                Dim slabsSystem As New StructuralSystem(sysch)

                oGenericSys.SetUserDefinedName("Imported Structure")
                sysch.SetUserDefinedName("Structural")
                columns.SetUserDefinedName("Columns")
                beams.SetUserDefinedName("Beams")
                braces.SetUserDefinedName("Braces")
                coordSystem.SetUserDefinedName("Coordinate Systems")
                slabsSystem.SetUserDefinedName("Slabs")

                ClientServiceProvider.TransactionMgr.Commit("ImportedStructure")

                '-------------------------- Base System Created ----------------------------------------
                'Dim temp222 As design_part = structuralObjects(4)
                'MsgBox(temp222.partDefinition.ToString())
                For c = 1 To structuralObjects.Count
                    Dim temp As design_part = structuralObjects(c)
                    Dim argc() As String = Split(temp.partDefinition.ToString, ".")

                    Dim p As New part '= temp.partDefinition

                    If LCase(argc(1)) = "part_prismatic_simple" Then
                        p.partTypeClass = temp.partDefinition
                    Else
                        p = temp.partDefinition
                    End If

                    Dim parent As Collection = temp.parentAssembly
                    Dim location As Collection = temp.location
                    Dim label As String = temp.label
                    Dim count As Integer

                    ' Design part processor ----- Will try to house all the parts within this class (e.g. "managed_data_item", etc.)
                    If p.partTypeClass IsNot Nothing Then
                        Dim argv() As String = Split(p.partTypeClass.ToString, ".")

                        Select Case LCase(argv(1))
                            Case "part_prismatic_simple"
                                Dim p1 As part_prismatic_simple = p.partTypeClass

                                If p1.sectionToUse.ToString <> "slab" Then
                                    Dim sec As item_reference_standard = p1.sectionToUse

                                    If sec.standardName <> "Section Not In Catalog" And sec.standardName <> "Unknown Section" Then
                                        For count = 1 To parent.Count
                                            Dim loc As coord_system_cartesian_3d = location(count)
                                            Dim steelMember As New steelObject

                                            steelMember.label = label
                                            steelMember.parentSystemLabel = parent(count)
                                            steelMember.lTypeCategory = beamOrColumn(loc)
                                            steelMember.bMir = False
                                            steelMember.lCard = CLng(sec.cardinalPoint)

                                            steelMember.oMaterial = oCatStruHlpr.GetMaterial("Steel - Carbon", "A36")
                                            steelMember.oStruct = oCatStruHlpr.GetCrossSection(sec.standardName, sec.sectionType, sec.sectionName)
                                            'steelMember.oStruct = oCatStruHlpr.GetCrossSection("AISC-SHAPES-3.1", "W", "W8x24")
                                            steelMember.startPoint = oPos1.Offset(New Vector(loc.axisSystem.cartPoint.X * globalUnit, loc.axisSystem.cartPoint.Y * globalUnit, loc.axisSystem.cartPoint.Z * globalUnit))
                                            steelMember.endPoint = oPos1.Offset(New Vector(loc.axisSystem.cartPoint.X * globalUnit + (p1.length * loc.axisSystem.direction2.X), loc.axisSystem.cartPoint.Y * globalUnit + (p1.length * loc.axisSystem.direction2.Y), loc.axisSystem.cartPoint.Z * globalUnit + (p1.length * loc.axisSystem.direction2.Z)))

                                            If steelMember.lTypeCategory = 1 Then
                                                steelMember.lType = CLng(MemberType.Beam)
                                                steelMember.dRotAngle = 0
                                                steelMember.structureObject = beams
                                                SP3DBeamObjects.Add(steelMember)
                                            ElseIf steelMember.lTypeCategory = 2 Then
                                                steelMember.lType = CLng(MemberType.Column)
                                                steelMember.dRotAngle = Math.Atan(loc.axisSystem.direction1.Y / loc.axisSystem.direction1.X)
                                                steelMember.structureObject = columns
                                                SP3DColumnObjects.Add(steelMember)
                                            ElseIf steelMember.lTypeCategory = 3 Then
                                                steelMember.lType = CLng(MemberType.VerticalBrace)
                                                steelMember.dRotAngle = 0
                                                steelMember.structureObject = braces
                                                SP3DBraceObjects.Add(steelMember)
                                            End If
                                        Next count

                                    Else
                                        Dim FILE_NAME As String = "C:\Documents and Settings\stded2\Desktop\errorLog.txt"
                                        Dim errorLine As String

                                        errorLine = sec.standardName & " -- Labeled: " & label

                                        Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)
                                        objWriter.WriteLine(errorLine)
                                        objWriter.Close()
                                    End If
                                Else                    'Slabs

                                End If
                        End Select
                    End If
                Next c
                'MsgBox(SP3DColumnObjects.Count.ToString & "||" & SP3DBeamObjects.Count.ToString & "||" & SP3DBraceObjects.Count.ToString)
                '-------------------------- Structure Objects Created ----------------------------------------
                reassessBraces(SP3DColumnObjects, SP3DBeamObjects, SP3DBraceObjects)

                buildSteelStructure(SP3DColumnObjects)
                buildSteelStructure(SP3DBeamObjects)
                buildSteelStructure(SP3DBraceObjects, True)
                '-------------------------- Structure Created and Placed ----------------------------------------

                Dim gridlineCollection As Collection

                If CheckBox2.Checked = True Then
                    gridlineCollection = makeGrids(SP3DColumnObjects, SP3DBeamObjects, CDbl(TextBox6.Text) / 1000, coordSystem)
                End If

                If CheckBox1.Checked = True Then
                    If Not IsNumeric(TextBox2.Text) Then
                        TextBox2.Text = 0
                    End If

                    makeRelations2(SP3DColumnObjects, SP3DBeamObjects, SP3DBraceObjects, gridlineCollection, CDbl(TextBox2.Text) / 1000)
                End If

                'gridlineCollection = Nothing
                'directives = Nothing
                'topLevelDirectives = Nothing
                'finalStandards = Nothing
                'createdColumnObjects = Nothing
                'createdBeamObjects = Nothing
                'createdBraceObjects = Nothing
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            MsgBox("Imported")
        Else
            MsgBox("File does not exist")
        End If


    End Sub



    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Label3.Enabled = True
            TextBox2.Enabled = True
            Label4.Enabled = True
        Else
            Label3.Enabled = False
            TextBox2.Enabled = False
            Label4.Enabled = False
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Label6.Enabled = True
            Label7.Enabled = True
            Label8.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
        Else
            Label6.Enabled = False
            Label7.Enabled = False
            Label8.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox6.Enabled = True
            Label11.Enabled = True
            Label12.Enabled = True
        Else
            TextBox6.Enabled = False
            Label11.Enabled = False
            Label12.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
       
        getDataTable()
        'oModel.PlantCatalog.Server.ToString        -- Server
        'oModel.PlantCatalog.Name.ToString          -- DB Catalog
        'oModel.PlantCatalog.ParentPlant.ToString   -- Plant Name

        'Dim m_oObj As BusinessObject
        'Dim m_oSelectSet As SelectSet
        'Dim aa As PipeEndFeature
        'Dim xx As PipeStraightFeature

        'm_oSelectSet = ClientServiceProvider.SelectSet
        'm_oObj = m_oSelectSet.SelectedObjects.Item(0)
        ''aa = m_oObj
        'xx = m_oObj

        'MsgBox(xx.Name)

        'Dim ff As PropertyValueDouble

        'ff = aa.GetPropertyValue("IJRtePipePathFeat", "NPD")
        'MsgBox(ff.PropValue)
        'xx = aa.SystemParent
        'MsgBox(xx.Name)


        'MsgBox(aa.Location.X.ToString & "||" & aa.Location.Y.ToString & "||" & aa.Location.Z.ToString)


        'MsgBox(aa.SystemParent.SystemChildren(0).ToString)
        'Dim c As Integer
        'Dim fred As String = ""

        'For c = 0 To xx.GetAllProperties.Count - 1
        '    'fred = fred & aa.GetAllProperties.Item(c).PropertyInfo.InterfaceInfo.Name & " -- " & aa.GetAllProperties.Item(c).PropertyInfo.Name & vbCrLf
        '    fred = fred & xx.GetAllProperties.Item(c).PropertyInfo.InterfaceInfo.Name & " -- " & xx.GetAllProperties.Item(c).PropertyInfo.Name & vbCrLf
        'Next

        'MsgBox(fred)



    End Sub

    Private Sub importBuildForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ComboBox1.SelectedIndex = 3
    End Sub
End Class

Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Systems.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Grids.Middle.GridElevationPlane
Imports Ingr.SP3D.Grids.Middle

Module structureUtilities
    Const piOverTwo = Math.PI / 2       ' 90 Degree rotation in Radials

    ''' <summary>
    ''' Re-labels a steel member object and parent system.
    ''' </summary>
    ''' <param name="memberSystemObject"></param>
    ''' <param name="parentName"></param>
    ''' <param name="childName"></param>
    ''' <remarks></remarks>
    Public Sub relabelSteel(ByRef memberSystemObject As MemberSystem, ByVal parentName As String, ByVal childName As String)
        If parentName <> "" Then                         ' Add labels -- Must be done after part creation
            memberSystemObject.SetPropertyValue(parentName, "IJNamedItem", "Name")
        End If

        Dim omem As MemberPart

        'these can be created in any order "end, member, start" (all steel is made of these 3 parts),
        'thus the control structure for labelling the correct piece.

        For c2 = 0 To memberSystemObject.SystemChildren.Count - 1
            If Mid(memberSystemObject.SystemChildren(c2).ToString, 1, 6) = "Member" Then
                omem = memberSystemObject.SystemChildren(c2)

                If Trim(childName) <> "" Then
                    omem.SetUserDefinedName(childName)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Attempts to place the steel members in the collection and add them to the steelObject.
    ''' </summary>
    ''' <param name="memberCollection">A collection of steelObject Objects.</param>
    ''' <remarks>The optional value thingy will need to be removed in the final version.  This is here to handle L-Steel which is unside down in SP3D</remarks>
    Public Sub buildSteelStructure(ByRef memberCollection As Collection, Optional ByVal aa As Boolean = False) 'As Collection
        Dim c As Integer
        Dim newSteelCollection As New Collection

        For c = 1 To memberCollection.Count
            Dim oModel As Model = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel
            Dim oRootSys As BusinessObject = oModel.RootSystem
            Dim oCatStruHlpr As New CatalogStructHelper(oModel.PlantCatalog)

            Dim temp As steelObject = memberCollection(c)

            If aa = True Then
                temp.dRotAngle = Math.PI
            End If

            Dim oMemSys1 As New MemberSystem(temp.structureObject, temp.startPoint, temp.endPoint, temp.oStruct, temp.oMaterial, temp.lType, temp.lTypeCategory, temp.lCard, temp.dRotAngle, temp.bMir)
            'Dim oMemSys1 As New MemberSystem(temp.structureObject, temp.startPoint, temp.endPoint, temp.oStruct, temp.oMaterial, temp.lType, temp.lCard, temp.dRotAngle, temp.bMir)

            relabelSteel(oMemSys1, temp.parentSystemLabel, temp.label)
            temp.createdObject = oMemSys1

            newSteelCollection.Add(temp)
        Next
       
        ClientServiceProvider.TransactionMgr.Commit("steel")
      

        memberCollection = newSteelCollection
    End Sub

    ''' <summary>
    ''' Determines if a beam is a column or a beam based on its vector normals. 1=beam, 2=column
    ''' </summary>
    ''' <param name="cartesianObject"></param>
    ''' <returns>e</returns>
    ''' <remarks>d</remarks>
    Public Function beamOrColumn(ByVal cartesianObject As coord_system_cartesian_3d) As Integer
        Dim temp As axis2_placement_3d = cartesianObject.axisSystem
        Dim ref1 As direction = temp.direction1
        Dim ref2 As direction = temp.direction2
        Dim result As Integer = 1

        If ref2.Z = 1 And ref2.X = 0 And ref2.Y = 0 Then
            result = 2
        End If

        If (ref2.Z <> 1 And ref2.Z <> 0 And ((ref2.X <> 0 And ref2.X <> 1) Or (ref2.Y <> 0 And ref2.Y <> 1))) Then
            result = 3
        End If

        beamOrColumn = result
    End Function

    ''' <summary>
    ''' Simple method to make false positive braces into beams.
    ''' </summary>
    ''' <param name="columnCollection"></param>
    ''' <param name="beamCollection"></param>
    ''' <param name="braceCollection"></param>
    ''' <remarks></remarks>
    Public Sub reassessBraces(ByRef columnCollection As Collection, ByRef beamCollection As Collection, braceCollection As Collection)
        Dim c As Integer
        Dim columnHighPoint As Double = -10000
        Dim killBraces As New Collection
        Dim beamStructureSystem As StructuralSystem = beamCollection(1).structureObject

        For c = 1 To columnCollection.Count
            Dim temp As steelObject = columnCollection(c)

            If temp.endPoint.Z > columnHighPoint Then
                columnHighPoint = temp.endPoint.Z
            End If

            If temp.startPoint.Z > columnHighPoint Then
                columnHighPoint = temp.startPoint.Z
            End If
        Next c

        For c = 1 To braceCollection.Count
            Dim temp As steelObject = braceCollection(c)

            If temp.endPoint.Z >= columnHighPoint Or temp.startPoint.Z >= columnHighPoint Then
                temp.structureObject = beamStructureSystem
                temp.lTypeCategory = MemberTypeCategory.Beam
                temp.lType = MemberType.Beam
                beamCollection.Add(temp)
                killBraces.Add(c)
            End If
        Next c

        For c = killBraces.Count To 1 Step -1
            braceCollection.Remove(killBraces(c))
        Next c
    End Sub

    Public Function getIntersectingGridLines(ByVal queryPoint As Position, ByRef gridlineCollection As Collection, ByVal tollerance As Double) As Collection
        Dim c As Integer
        Dim gridItemsIntersection As New Collection
        Dim highX#, lowX#, highY#, lowY#

        For c = 1 To gridlineCollection.Count
            Dim temp As GridLine = gridlineCollection(c)

            If temp.StartPoint.X > temp.EndPoint.X Then
                highX = temp.StartPoint.X
                lowX = temp.EndPoint.X
            Else
                lowX = temp.StartPoint.X
                highX = temp.EndPoint.X
            End If

            If temp.StartPoint.Y > temp.EndPoint.Y Then
                highY = temp.StartPoint.Y
                lowY = temp.EndPoint.Y
            Else
                lowY = temp.StartPoint.Y
                highY = temp.EndPoint.Y
            End If

            If compareNumbers(temp.StartPoint.Z, queryPoint.Z, tollerance) = True Then
                If isPointOnALine(queryPoint, temp.StartPoint, temp.EndPoint, tollerance) = True Then
                    ' Added check for items on a shared point and have two planes in the same direction
                    If ((queryPoint.X <= highX And queryPoint.X >= lowX) Or (Math.Abs(queryPoint.X - temp.StartPoint.X) <= tollerance) Or (Math.Abs(queryPoint.X - temp.EndPoint.X) <= tollerance)) And ((queryPoint.Y <= highY And queryPoint.Y >= lowY) Or (Math.Abs(queryPoint.Y - temp.StartPoint.Y) <= tollerance) Or (Math.Abs(queryPoint.Y - temp.EndPoint.Y) <= tollerance)) Then
                        gridItemsIntersection.Add(c)
                    End If
                End If

                If gridItemsIntersection.Count >= 2 Then
                    Exit For
                End If
            End If
        Next c

        getIntersectingGridLines = gridItemsIntersection
    End Function

    Public Sub makeRelations2(ByRef SP3DColumnObjects As Collection, ByRef SP3DBeamObjects As Collection, ByRef SP3DBraceObjects As Collection, ByVal gridlineCollection As Collection, ByVal tollerance As Double)
        Dim c%, c2%
        Dim usedBeams As New Collection

        '------------------------------------------------------- Create associations between columns and gridlines
        For c = 1 To SP3DColumnObjects.Count
            Dim tempSteel As steelObject = SP3DColumnObjects(c)
            Dim startConnections As Collection = getIntersectingGridLines(tempSteel.startPoint, gridlineCollection, tollerance)
            Dim endConnections As Collection = getIntersectingGridLines(tempSteel.endPoint, gridlineCollection, tollerance)

            If startConnections.Count > 1 Then
                Dim oRelatedMemSys As New Collection(Of BusinessObject)

                oRelatedMemSys.Add(gridlineCollection(startConnections(1)))
                oRelatedMemSys.Add(gridlineCollection(startConnections(2)))

                tempSteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
            End If

            If endConnections.Count > 1 Then
                Dim oRelatedMemSys2 As New Collection(Of BusinessObject)

                oRelatedMemSys2.Add(gridlineCollection(endConnections(1)))
                oRelatedMemSys2.Add(gridlineCollection(endConnections(2)))

                tempSteel.createdObject.SetRelatedObjects(oRelatedMemSys2, MemberAxisEnd.End)
            End If
        Next c

        ClientServiceProvider.TransactionMgr.Commit("Column Grid Associations")

        '------------------------------------------------------- Create associations between columns and gridlines

        For c = 1 To SP3DBeamObjects.Count
            Dim tempSteel As steelObject = SP3DBeamObjects(c)
            Dim startColumn As Integer = getMy3dColumnConnection(SP3DColumnObjects, tempSteel.startPoint, 0.0001)
            Dim endColumn As Integer = getMy3dColumnConnection(SP3DColumnObjects, tempSteel.endPoint, 0.0001)
            Dim startConnections As Collection = getIntersectingGridLines(tempSteel.startPoint, gridlineCollection, tollerance)
            Dim endConnections As Collection = getIntersectingGridLines(tempSteel.endPoint, gridlineCollection, tollerance)

            If startColumn > 0 And startConnections.Count > 1 Then
                Dim oRelatedMemSys As New Collection(Of BusinessObject)

                If isPointOnALine(tempSteel.startPoint, gridlineCollection(startConnections(1)).StartPoint, gridlineCollection(startConnections(1)).EndPoint, 0.0001) = True And isPointOnALine(tempSteel.endPoint, gridlineCollection(startConnections(1)).StartPoint, gridlineCollection(startConnections(1)).EndPoint, 0.0001) = True Then
                    c2 = startConnections(2)
                Else
                    c2 = startConnections(1)
                End If

                oRelatedMemSys.Add(gridlineCollection(c2))
                oRelatedMemSys.Add(SP3DColumnObjects(startColumn).createdObject)

                tempSteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)

                usedBeams.Add(c, "s" & c)
            End If

            If endColumn > 0 And endConnections.Count > 1 Then
                Dim oRelatedMemSys As New Collection(Of BusinessObject)

                If isPointOnALine(tempSteel.startPoint, gridlineCollection(startConnections(1)).StartPoint, gridlineCollection(startConnections(1)).EndPoint, 0.0001) = True And isPointOnALine(tempSteel.endPoint, gridlineCollection(startConnections(1)).StartPoint, gridlineCollection(startConnections(1)).EndPoint, 0.0001) = True Then
                    c2 = endConnections(2)
                Else
                    c2 = endConnections(1)
                End If

                oRelatedMemSys.Add(gridlineCollection(c2))
                oRelatedMemSys.Add(SP3DColumnObjects(endColumn).createdObject)

                tempSteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.End)

                usedBeams.Add(c, "e" & c)
            End If
        Next c

        ClientServiceProvider.TransactionMgr.Commit("Column Grid Associations")

        '------------------------------------------------------- Handle the left over beams

        For c = 1 To SP3DBeamObjects.Count
            Dim tempSteel As steelObject = SP3DBeamObjects(c)

            If existsInCollection(usedBeams, "s" & c) = "DNE" Then
                Dim beamEnds As Collection = getMyBeamConnections(SP3DBeamObjects, tempSteel.startPoint, usedBeams)

                If beamEnds.Count = 2 Then
                    Dim oRelatedMemSys As New Collection(Of BusinessObject)

                    oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(2)).createdObject)
                    oRelatedMemSys.Add(New Point3d(tempSteel.startPoint))

                    tempSteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
                ElseIf beamEnds.Count > 2 Then
                    Dim answer As Collection = getParallelsForBeams(beamEnds)
                    Dim b1% = 1, b2% = 2

                    If answer.Count > 0 Then
                        b1 = answer(1)
                        b2 = answer(2)
                    End If

                    Dim baseSteelConnecter As steelObject = SP3DBeamObjects(beamEnds(b1).beamInCollection)
                    Dim oRelatedMemSys As New Collection(Of BusinessObject)

                    oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(b2).beamInCollection).createdObject)
                    oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(b1).beamInCollection).createdObject)

                    If beamEnds(b1).where = "s" Then
                        baseSteelConnecter.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
                    Else
                        baseSteelConnecter.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.End)
                    End If

                    For c2 = 1 To beamEnds.Count
                        If c2 <> b1 And c2 <> b2 Then
                            Dim baseSteelConnecter2 As steelObject = SP3DBeamObjects(beamEnds(c2).beamInCollection)
                            Dim oRelatedMemSys2 As New Collection(Of BusinessObject)


                            If beamEnds(b1).where = "s" Then
                                oRelatedMemSys2.Add(baseSteelConnecter.createdObject.FrameConnection(MemberAxisEnd.Start))
                            Else
                                oRelatedMemSys2.Add(baseSteelConnecter.createdObject.FrameConnection(MemberAxisEnd.End))
                            End If

                            If beamEnds(c2).where = "s" Then
                                baseSteelConnecter2.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
                            Else
                                baseSteelConnecter2.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.End)
                            End If
                        End If
                    Next c2
                End If

                For c2 = 1 To beamEnds.Count
                    addToCollection(usedBeams, beamEnds(c2).beamInCollection, beamEnds(c2).where & beamEnds(c2).beamInCollection)
                Next c2
            End If

            If existsInCollection(usedBeams, "e" & c) = "DNE" Then
                Dim beamEnds As Collection = getMyBeamConnections(SP3DBeamObjects, tempSteel.endPoint, usedBeams)

                If beamEnds.Count = 2 Then
                    Dim oRelatedMemSys As New Collection(Of BusinessObject)

                    oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(2)).createdObject)
                    oRelatedMemSys.Add(New Point3d(tempSteel.endPoint))

                    tempSteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
                ElseIf beamEnds.Count > 2 Then
                    Dim answer As Collection = getParallelsForBeams(beamEnds)
                    Dim b1% = 1, b2% = 2

                    If answer.Count > 0 Then
                        b1 = answer(1)
                        b2 = answer(2)
                    End If

                    Dim baseSteelConnecter As steelObject = SP3DBeamObjects(beamEnds(b1).beamInCollection)
                    Dim oRelatedMemSys As New Collection(Of BusinessObject)

                    oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(b2).beamInCollection).createdObject)
                    oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(b1).beamInCollection).createdObject)

                    If beamEnds(b1).where = "s" Then
                        baseSteelConnecter.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
                    Else
                        baseSteelConnecter.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.End)
                    End If

                    For c2 = 1 To beamEnds.Count
                        If c2 <> b1 And c2 <> b2 Then
                            Dim baseSteelConnecter2 As steelObject = SP3DBeamObjects(beamEnds(c2).beamInCollection)
                            Dim oRelatedMemSys2 As New Collection(Of BusinessObject)


                            If beamEnds(b1).where = "s" Then
                                oRelatedMemSys2.Add(baseSteelConnecter.createdObject.FrameConnection(MemberAxisEnd.Start))
                            Else
                                oRelatedMemSys2.Add(baseSteelConnecter.createdObject.FrameConnection(MemberAxisEnd.End))
                            End If

                            If beamEnds(c2).where = "s" Then
                                baseSteelConnecter2.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
                            Else
                                baseSteelConnecter2.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.End)
                            End If
                        End If
                    Next c2
                End If

                For c2 = 1 To beamEnds.Count
                    addToCollection(usedBeams, beamEnds(c2).beamInCollection, beamEnds(c2).where & beamEnds(c2).beamInCollection)
                Next c2
            End If

        Next c

        ClientServiceProvider.TransactionMgr.Commit("Column Grid Associations")

        '------------------------------------------------------- Connect the braces

        For c = 1 To SP3DBraceObjects.Count
            Dim fake As New Collection

            Dim tempsteel As steelObject = SP3DBraceObjects(c)
            Dim tempSteelAngle As Double = getRoundedNormalizedAngle(tempsteel.startPoint, tempsteel.endPoint)
            Dim beamEnds As Collection = getMyBeamConnections(SP3DBeamObjects, tempsteel.endPoint, fake, 0.3)
            Dim beamStarts As Collection = getMyBeamConnections(SP3DBeamObjects, tempsteel.startPoint, fake, 0.3)

            If beamEnds.Count > 0 Then
                For c2 = 1 To beamEnds.Count
                    If tempSteelAngle = beamEnds(c2).angle Or Math.Abs(tempSteelAngle - beamEnds(c2).angle) = 180 Then
                        Exit For
                    End If
                Next

                If c2 <= beamEnds.Count Then
                    Dim oRelatedMemSys As New Collection(Of BusinessObject)

                    If beamEnds(c2).where = "s" Then
                        oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(c2).beamInCollection).createdObject.FrameConnection(MemberAxisEnd.Start))
                    Else
                        oRelatedMemSys.Add(SP3DBeamObjects(beamEnds(c2).beamInCollection).createdObject.FrameConnection(MemberAxisEnd.End))
                    End If

                    tempsteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.End)
                End If
            Else
                Dim oRelatedMemSys As New Collection(Of BusinessObject)

                oRelatedMemSys.Add(SP3DColumnObjects(getMy3dColumnConnection(SP3DColumnObjects, tempsteel.endPoint, 0.0001)).createdObject)
                oRelatedMemSys.Add(New Point3d(tempsteel.startPoint))

                tempsteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.End)
            End If

            If beamStarts.Count > 0 Then
                For c2 = 1 To beamStarts.Count
                    If tempSteelAngle = beamStarts(c2).angle Or Math.Abs(tempSteelAngle - beamStarts(c2).angle) = 180 Then
                        Exit For
                    End If
                Next

                If c2 <= beamStarts.Count Then
                    Dim oRelatedMemSys As New Collection(Of BusinessObject)

                    If beamStarts(c2).where = "s" Then
                        oRelatedMemSys.Add(SP3DBeamObjects(beamStarts(c2).beamInCollection).createdObject.FrameConnection(MemberAxisEnd.Start))
                    Else
                        oRelatedMemSys.Add(SP3DBeamObjects(beamStarts(c2).beamInCollection).createdObject.FrameConnection(MemberAxisEnd.End))
                    End If

                    tempsteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
                End If
            Else
                Dim oRelatedMemSys As New Collection(Of BusinessObject)

                oRelatedMemSys.Add(SP3DColumnObjects(getMy3dColumnConnection(SP3DColumnObjects, tempsteel.startPoint, 0.0001)).createdObject)
                oRelatedMemSys.Add(New Point3d(tempsteel.endPoint))

                tempsteel.createdObject.SetRelatedObjects(oRelatedMemSys, MemberAxisEnd.Start)
            End If
        Next c

        ClientServiceProvider.TransactionMgr.Commit("Column Grid Associations")
    End Sub

    Public Function makeGrids(ByRef SP3DColumnObjects As Collection, ByRef SP3DBeamObjects As Collection, ByVal tolerance As Double, ByRef coordSystem As StructuralSystem) As Collection
        Dim c As Integer
        Dim c2 As Integer
        Dim c3 As Integer
        Dim topOfSteel#, bottomOfSteel#
        Dim rotational As Double = normalizeTo90(SP3DColumnObjects(1).dRotAngle)
        Dim columnAngles As New Collection
        Dim oTransactionMgr As TransactionManager = MiddleServiceProvider.TransactionMgr
        Dim oSiteMgr As SiteManager = MiddleServiceProvider.SiteMgr
        Dim oPlant As Plant = oSiteMgr.ActiveSite.ActivePlant
        Dim oModel As Model = oPlant.PlantModel
        Dim oRootObj As BusinessObject = DirectCast(oModel.RootSystem, BusinessObject)
        Dim oModelConn As SP3DConnection = oRootObj.DBConnection
        Dim gridPlanes As New Collection
        Dim previous As Integer
        ' Makes a collection of unique column angles -- This will give us the number of coordinate systems that will need to be made.
        ' This is basic right now.  We may want to explore a smarter algorithm to create radial grids.
        '
        ' Basically -- If the count is 1, then it is a square grid.  If the count raises to 2, it is likely a building with an angle.  3 and above 
        ' likely has a radial (part of a circle) portion.
        For c = 1 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)

            If existsInCollection(columnAngles, "a" & temp.dRotAngle) = "DNE" And existsInCollection(columnAngles, "a" & (temp.dRotAngle + piOverTwo)) = "DNE" Then
                columnAngles.Add(temp.dRotAngle, "a" & temp.dRotAngle)
            End If
        Next

        '---------------------------------------------- Making Grid-Planes

        For c = 1 To columnAngles.Count
            Dim tempSteelCollection As New Collection
            Dim subCollections As New Collection

            Dim complete As Integer
            Dim currentCount As Integer = 0

            rotational = columnAngles(c)

            For c2 = 1 To SP3DColumnObjects.Count                   ' Creates a sub collection base on angles
                Dim steel As steelObject = SP3DColumnObjects(c2)

                If steel.dRotAngle = rotational Then
                    tempSteelCollection.Add(steel)
                End If
            Next c2

            complete = tempSteelCollection.Count
            previous = tempSteelCollection.Count

            While currentCount < complete
                Dim subCollections2 As Collection = makeGridSubCollections(tempSteelCollection, SP3DBeamObjects, tolerance)

                currentCount = currentCount - (subCollections2.Count - 1)

                For c3 = 1 To subCollections2.Count
                    currentCount = currentCount + subCollections2(c3).Count
                    subCollections.Add(subCollections2(c3))

                    If currentCount < complete Then
                        removeSteelObjectsFromCollection(tempSteelCollection, subCollections2(c3))

                        If collectionCount(tempSteelCollection) < 1 Then
                            Exit While
                        End If
                    Else
                        Exit While
                    End If

                    If collectionCount(tempSteelCollection) = previous Then
                        Exit While
                    End If

                    previous = collectionCount(tempSteelCollection)
                Next
            End While

            For c3 = 1 To subCollections.Count
                Dim elevations As Collection = getElevations(subCollections(c3), SP3DBeamObjects)

                topOfSteel = elevations.Item("topOfSteel")
                bottomOfSteel = elevations.Item("bottomOfSteel")
                elevations.Remove("topOfSteel")
                elevations.Remove("bottomOfSteel")

                Dim oCS As New CoordinateSystem(oModelConn, CoordinateSystem.CoordinateSystemType.Grids)
                'Dim gridPlanes As Collection = createGridPlanes(oCS, subCollections(c3), coordSystem, elevations, bottomOfSteel, tolerance)
                Dim tempGridPlanes As Collection = createGridPlanes(oCS, subCollections(c3), coordSystem, elevations, bottomOfSteel, tolerance)

                For c2 = 1 To tempGridPlanes.Count
                    gridPlanes.Add(tempGridPlanes(c2))
                Next c2

                If rotational <> 0 Then
                    rotateObject(oCS, columnAngles(c), "Z")
                End If

                moveObject(oCS, subCollections(c3)(1).startPoint)
            Next c3
        Next c

        oTransactionMgr.Commit("Grids")

        makeGrids = gridPlanes
    End Function

    Public Function makeGridSubCollections(ByRef SP3DColumnObjects As Collection, ByRef SP3DBeamObjects As Collection, ByVal tolerance As Double) As Collection
        Dim c As Integer
        Dim c2 As Integer = 1
        Dim topRight As Integer = 1
        Dim connectionResult As Integer
        Dim current As Integer = 0
        Dim complete As Boolean = False
        Dim perimeterCollection As New Collection
        Dim repeats As New Collection
        Dim result As New Collection
        Dim beamsused As New Collection

        If SP3DColumnObjects(1).dRotAngle <> 0 Then
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DColumnObjects, -SP3DColumnObjects(1).dRotAngle)
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DBeamObjects, -SP3DColumnObjects(1).dRotAngle)
        End If

        ' Get the top right Column, emphasizing the top over the right
        For c = 2 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)

            If (temp.startPoint.X > SP3DColumnObjects(topRight).startPoint.X And temp.startPoint.Y >= SP3DColumnObjects(topRight).startPoint.Y) Or temp.startPoint.Y > SP3DColumnObjects(topRight).startPoint.Y Then
                topRight = c
            End If
        Next c

        ' Add the start point
        perimeterCollection.Add(topRight, "c" & SP3DColumnObjects(topRight).startPoint.X & "/" & SP3DColumnObjects(topRight).startPoint.Y)
        connectionResult = topRight

        ' Traverse the columns to get a perimeter
        While complete = False
            connectionResult = getLogicalConnection(SP3DColumnObjects(connectionResult), SP3DColumnObjects, SP3DBeamObjects, beamsused)

            If existsInCollection(perimeterCollection, "c" & connectionResult) = "DNE" And connectionResult <> topRight And connectionResult <> 0 Then
                perimeterCollection.Add((connectionResult), "c" & connectionResult)
            ElseIf connectionResult <> 0 Then
                perimeterCollection.Add(connectionResult)

                If connectionResult <> topRight Then
                    repeats.Add(connectionResult, "c" & connectionResult)
                End If
            End If

            If connectionResult = topRight Or connectionResult = 0 Then
                complete = True
                Exit While
            End If
        End While

        If repeats.Count = 0 Then
            Dim BB As twoDBoundingBox = get2DBoundingBox(SP3DColumnObjects, perimeterCollection)
            Dim finalCollection As Collection = getColumnsInBoundingBox(SP3DColumnObjects, BB)
            result.Add(finalCollection)
        Else
            For c = 1 To repeats.Count + 1
                Dim tempCollection As Collection = makePerimeter(perimeterCollection, repeats, c)
                Dim BB As twoDBoundingBox = get2DBoundingBox(SP3DColumnObjects, tempCollection)
                Dim finalCollection As Collection = getColumnsInBoundingBox(SP3DColumnObjects, BB)

                result.Add(finalCollection)
            Next c
        End If

        If SP3DColumnObjects(1).dRotAngle <> 0 Then
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DColumnObjects, SP3DColumnObjects(1).dRotAngle)
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DBeamObjects, SP3DColumnObjects(1).dRotAngle)
        End If

        makeGridSubCollections = result
    End Function

    ''' <summary>
    ''' Makes the grid sub-collections where there are crossovers.
    ''' </summary>
    ''' <param name="perimeterCollection"></param>
    ''' <param name="repeats"></param>
    ''' <param name="level"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function makePerimeter(ByVal perimeterCollection As Collection, ByVal repeats As Collection, ByVal level As Integer) As Collection
        Dim c As Integer
        Dim result As New Collection
        Dim whereAmI As Integer = 1
        Dim whoThrewMeOff As Integer

        If level = 1 Then
            whoThrewMeOff = 0
        Else
            whoThrewMeOff = repeats(level - 1)
        End If

        For c = 1 To perimeterCollection.Count
            Dim ans As String = existsInCollection(repeats, "c" & perimeterCollection(c))

            If ans <> "DNE" Then
                If whoThrewMeOff = 0 Then
                    whoThrewMeOff = CInt(ans)
                ElseIf whoThrewMeOff = CInt(ans) Then
                    whoThrewMeOff = 0
                End If
            End If

            If whoThrewMeOff = 0 Then
                result.Add(perimeterCollection(c))
            End If
        Next

        makePerimeter = result
    End Function

    ''' <summary>
    ''' Gets the next logical beam to travel to
    ''' </summary>
    ''' <param name="SP3DColumnObjects">Column to travel from</param>
    ''' <param name="SP3DColumnObjects2">Collection of columns</param>
    ''' <param name="SP3DBeamObjects">Beam collection</param>
    ''' <param name="lastDir">Last direction we came from, so we do not backtrack</param>
    ''' <param name="beamsUsed">collection of used beams</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLogicalConnection(ByVal SP3DColumnObjects As steelObject, ByVal SP3DColumnObjects2 As Collection, ByVal SP3DBeamObjects As Collection, ByRef beamsUsed As Collection) As Integer
        Dim c As Integer
        Dim result As Integer
        Dim beamUsed As Integer = 0
        Dim high#, low#
        Dim miniTraverse As New mt

        result = 0

        For c = 1 To SP3DBeamObjects.Count
            Dim tempSteel As steelObject = SP3DBeamObjects(c)
            Dim temp As Position = Nothing

            If SP3DColumnObjects.startPoint.Z > SP3DColumnObjects.endPoint.Z Then   ' Make sure the beam is within the bounds of the column
                high = SP3DColumnObjects.startPoint.Z
                low = SP3DColumnObjects.endPoint.Z
            Else
                low = SP3DColumnObjects.startPoint.Z
                high = SP3DColumnObjects.endPoint.Z
            End If

            If existsInCollection(beamsUsed, "c" & c.ToString) = "DNE" Then
                If compareNumbers(tempSteel.endPoint.X, SP3DColumnObjects.startPoint.X, 0.005) = True And compareNumbers(tempSteel.endPoint.Y, SP3DColumnObjects.startPoint.Y, 0.005) = True And (tempSteel.endPoint.Z <= high And tempSteel.endPoint.Z >= low And tempSteel.startPoint.Z <= high And tempSteel.startPoint.Z >= low) Then
                    temp = tempSteel.startPoint
                ElseIf compareNumbers(tempSteel.startPoint.X, SP3DColumnObjects.startPoint.X, 0.005) = True And compareNumbers(tempSteel.startPoint.Y, SP3DColumnObjects.startPoint.Y, 0.005) = True And (tempSteel.endPoint.Z <= high And tempSteel.endPoint.Z >= low And tempSteel.startPoint.Z <= high And tempSteel.startPoint.Z >= low) Then
                    temp = tempSteel.endPoint
                End If
            End If

            If temp IsNot Nothing Then          ' Get the Candidate beams.
                If compareLessThan(SP3DColumnObjects.startPoint.X, temp.X, 0.05) = True Then
                    miniTraverse.East = getMyColumnConnection(SP3DColumnObjects2, temp)
                    miniTraverse.EastBeam = c
                ElseIf compareNumbers(temp.X, SP3DColumnObjects.startPoint.X, 0.05) = True And compareLessThan(SP3DColumnObjects.startPoint.Y, temp.Y, 0.05) = True Then
                    miniTraverse.North = getMyColumnConnection(SP3DColumnObjects2, temp)
                    miniTraverse.NorthBeam = c
                ElseIf compareLessThan(temp.X, SP3DColumnObjects.startPoint.X, 0.05) = True Then
                    miniTraverse.West = getMyColumnConnection(SP3DColumnObjects2, temp)
                    miniTraverse.WestBeam = c
                ElseIf compareNumbers(temp.X, SP3DColumnObjects.startPoint.X, 0.05) = True And compareLessThan(temp.Y, SP3DColumnObjects.startPoint.Y, 0.05) = True Then
                    miniTraverse.South = getMyColumnConnection(SP3DColumnObjects2, temp)
                    miniTraverse.SouthBeam = c
                End If
            End If
        Next c

        If miniTraverse.West <> 0 And miniTraverse.North <> 0 And miniTraverse.East <> 0 Then
            addToCollection(beamsUsed, miniTraverse.EastBeam, "c" & miniTraverse.EastBeam)
            result = miniTraverse.East
        ElseIf miniTraverse.West <> 0 And miniTraverse.North <> 0 Then
            addToCollection(beamsUsed, miniTraverse.NorthBeam, "c" & miniTraverse.NorthBeam)
            result = miniTraverse.North
        ElseIf miniTraverse.West <> 0 Then
            addToCollection(beamsUsed, miniTraverse.WestBeam, "c" & miniTraverse.WestBeam)
            result = miniTraverse.West
        ElseIf miniTraverse.South <> 0 Then
            addToCollection(beamsUsed, miniTraverse.SouthBeam, "c" & miniTraverse.SouthBeam)
            result = miniTraverse.South
        ElseIf miniTraverse.East <> 0 Then
            addToCollection(beamsUsed, miniTraverse.EastBeam, "c" & miniTraverse.EastBeam)
            result = miniTraverse.East
        ElseIf miniTraverse.North <> 0 And miniTraverse.East <> 0 Then
            addToCollection(beamsUsed, miniTraverse.EastBeam, "c" & miniTraverse.EastBeam)
            result = miniTraverse.East
        ElseIf miniTraverse.North <> 0 Then
            addToCollection(beamsUsed, miniTraverse.NorthBeam, "c" & miniTraverse.NorthBeam)
            result = miniTraverse.North
        End If

        getLogicalConnection = result
    End Function

    Public Function getMy3dColumnConnection(ByVal SP3DColumnObjects As Collection, ByVal steel As Position, Optional ByVal precision As Double = 0) As Integer
        Dim c As Integer
        Dim result As Integer = 0
        Dim top#, bottom#

        For c = 1 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)

            If temp.startPoint.Z > temp.endPoint.Z Then
                top = temp.startPoint.Z
                bottom = temp.endPoint.Z
            Else
                bottom = temp.startPoint.Z
                top = temp.endPoint.Z
            End If

            If compareNumbers(temp.startPoint.X, steel.X, 0.005) = True And compareNumbers(temp.startPoint.Y, steel.Y, 0.005) = True Then
                If (steel.Z <= top And steel.Z >= bottom) Or (Math.Abs(steel.Z - top) <= precision) Or (Math.Abs(steel.Z - bottom) <= precision) Then
                    result = c

                    Exit For
                End If
            End If
        Next c

        getMy3dColumnConnection = result
    End Function

    Public Function getMyColumnConnection(ByVal SP3DColumnObjects As Collection, ByVal steel As Position, Optional ByVal review As Boolean = False) As Integer
       Dim c As Integer
        Dim result As Integer = 0

        For c = 1 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)

            If review = True Then
                MsgBox(temp.startPoint.X & "||" & steel.X & "//" & temp.startPoint.Y & "||" & steel.Y & "||" & c)
            End If

            If compareNumbers(temp.startPoint.X, steel.X, 0.005) = True And compareNumbers(temp.startPoint.Y, steel.Y, 0.005) = True Then
                result = c

                Exit For
            End If
        Next c

        getMyColumnConnection = result
    End Function

    ''' <summary>
    ''' Gets beams/beam ends located at a point.  Excludes beams in the "usedBeams" collection (integer collection). 
    ''' </summary>
    ''' <param name="SP3DBeamObjects"></param>
    ''' <param name="steel"></param>
    ''' <param name="usedBeams"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getMyBeamConnections(ByVal SP3DBeamObjects As Collection, ByVal steel As Position, ByVal usedBeams As Collection, Optional ByVal tolerance As Double = 0.0001) As Collection
        Dim c As Integer
        Dim result As New Collection

        For c = 1 To SP3DBeamObjects.Count
            Dim temp As steelObject = SP3DBeamObjects(c)
            Dim bounds As New cartesianBounds(temp.startPoint, temp.endPoint)
            Dim connection As New beamPositionConnections

            connection.beamInCollection = c
            connection.angle = getRoundedNormalizedAngle(temp.startPoint, temp.endPoint)

            If compareNumbers(temp.startPoint.X, steel.X, tolerance) = True And compareNumbers(temp.startPoint.Y, steel.Y, tolerance) = True And compareNumbers(temp.startPoint.Z, steel.Z, tolerance) = True Then
                If existsInCollection(usedBeams, "s" & c) = "DNE" Then
                    connection.where = "s"
                End If
            ElseIf compareNumbers(temp.endPoint.X, steel.X, tolerance) = True And compareNumbers(temp.endPoint.Y, steel.Y, tolerance) = True And compareNumbers(temp.endPoint.Z, steel.Z, tolerance) = True Then
                If existsInCollection(usedBeams, "e" & c) = "DNE" Then
                    connection.where = "e"
                End If
            ElseIf bounds.inBounds(steel) = True And isPointOnALine3D(steel, temp.startPoint, temp.endPoint, tolerance) = True Then
                connection.where = "m"
            End If

            If connection.where <> "" Then
                result.Add(connection)
            End If
        Next c

        getMyBeamConnections = result
    End Function

    Public Function getParallelsForBeams(ByVal checkCollection As Collection) As Collection
        Dim c%, c2%
        Dim temp As New Collection

        For c = 1 To checkCollection.Count - 1
            For c2 = c + 1 To checkCollection.Count
                If checkCollection(c).angle = checkCollection(c2).angle Then
                    temp.Add(c)
                    temp.Add(c2)

                    c = checkCollection.Count
                    Exit For
                End If
            Next c2
        Next c

        getParallelsForBeams = temp
    End Function

    ''' <summary>
    ''' Rotates a steel collection about the Z-Axis, with respect to the primary point
    ''' </summary>
    ''' <param name="primePoint">Point to rotate about</param>
    ''' <param name="steelCollection">A collection of steel</param>
    ''' <param name="rotationAngle">Angle to rotate to</param>
    ''' <remarks></remarks>
    Public Sub rotateStructure(ByVal primePoint As Position, ByRef steelCollection As Collection, ByVal rotationAngle As Double, Optional ByVal setAngleTo0 As Boolean = False)
        Dim x#, y#
        Dim xp#, yp#
        Dim xpp#, ypp#
        Dim cosRot# = Math.Cos(rotationAngle)
        Dim sinRot# = Math.Sin(rotationAngle)
        Dim c As Integer

        x = primePoint.X
        y = primePoint.Y

        For c = 1 To steelCollection.Count
            Dim steel As steelObject = steelCollection(c)

            xp = steel.startPoint.X - x
            yp = steel.startPoint.Y - y

            xpp = xp * cosRot - yp * sinRot
            ypp = xp * sinRot + yp * cosRot

            steel.startPoint.X = xpp + x
            steel.startPoint.Y = ypp + y

            xp = steel.endPoint.X - x
            yp = steel.endPoint.Y - y

            xpp = xp * cosRot - yp * sinRot
            ypp = xp * sinRot + yp * cosRot

            steel.endPoint.X = xpp + x
            steel.endPoint.Y = ypp + y

            If setAngleTo0 = True Then
                steel.dRotAngle = 0
            End If
        Next
    End Sub

    ''' <summary>
    ''' Creates the grid-planes and grid-lines from the collections provided.
    ''' </summary>
    ''' <param name="oCS">the coordinate system to use</param>
    ''' <param name="SP3DColumnObjects">A collection of the columns -- Datatype, not the actual created columns</param>
    ''' <param name="coordSystem">Structural System</param>
    ''' <param name="elevations">A collection of doubles of the various Z-plane heights to create</param>
    ''' <param name="bottomOfSteel">The bottom of structure -- To scale the building position</param>
    ''' <param name="tolerance">The amount a column can be off the plane to be in the collection</param>
    ''' <returns>The created line objects</returns>
    ''' <remarks></remarks>
    Public Function createGridPlanes(ByRef oCS As CoordinateSystem, ByVal SP3DColumnObjects As Collection, ByVal coordSystem As StructuralSystem, ByVal elevations As Collection, ByVal bottomOfSteel As Double, ByVal tolerance As Double) As Collection
        'Dim XCollection As New Collection
        'Dim YCollection As New Collection
        'Dim strPlane As String
        'Dim c As Integer

        'For c = 1 To SP3DColumnObjects.Count
        '    Dim tempSteel As steelObject = SP3DColumnObjects(c)

        '    If existsInCollection(XCollection, "c" & tempSteel.startPoint.X) = "DNE" Then
        '        XCollection.Add(c, "c" & tempSteel.startPoint.X)
        '    End If

        '    If existsInCollection(YCollection, "c" & tempSteel.startPoint.Y) = "DNE" Then
        '        YCollection.Add(c, "c" & tempSteel.startPoint.Y)
        '    End If
        'Next c

        'Dim oSystem As ISystem = New SystemHelper(coordSystem)

        'oSystem.AddSystemChild(oCS)

        'Dim oGridAxis As New GridAxis(oCS, AxisType.Z)
        'Dim gridPlaneCollection As New Collection
        'Dim ZPlanes As New Collection

        'Dim oElevation As GridElevationPlane = Nothing

        'For c = 1 To elevations.Count
        '    oElevation = New GridElevationPlane(elevations(c) - bottomOfSteel, oGridAxis)
        '    strPlane = "Elevation_" & c
        '    oElevation.SetPropertyValue(strPlane, "IJNamedItem", "Name")
        '    ZPlanes.Add(oElevation)
        'Next c

        'oGridAxis = New GridAxis(oCS, AxisType.X)

        'Dim oGridXPlane As GridPlane = Nothing

        'For c = 1 To XCollection.Count
        '    oGridXPlane = New GridPlane(getHypotenuseComponent(SP3DColumnObjects(1).startPoint.X, SP3DColumnObjects(XCollection(c)).startPoint.X, SP3DColumnObjects(1).startPoint.Y, SP3DColumnObjects(XCollection(c)).startPoint.Y, "X"), oGridAxis)
        '    strPlane = "GridX_" & c.ToString()
        '    oGridXPlane.SetPropertyValue(strPlane, "IJNamedItem", "Name")

        '    gridPlaneCollection.Add(oGridXPlane)
        'Next

        'oGridAxis = New GridAxis(oCS, AxisType.Y)

        'Dim oGridYPlane As GridPlane = Nothing

        'For c = 1 To YCollection.Count
        '    oGridYPlane = New GridPlane(getHypotenuseComponent(SP3DColumnObjects(1).startPoint.X, SP3DColumnObjects(YCollection(c)).startPoint.X, SP3DColumnObjects(1).startPoint.Y, SP3DColumnObjects(YCollection(c)).startPoint.Y, "Y"), oGridAxis)
        '    strPlane = "GridY_" & c.ToString()
        '    oGridYPlane.SetPropertyValue(strPlane, "IJNamedItem", "Name")

        '    gridPlaneCollection.Add(oGridYPlane)
        'Next

        ''---------------------------------------------- Actually create the lines
        'Dim result As New Collection

        'For c2 = 1 To ZPlanes.Count
        '    Dim temp As GridElevationPlane = ZPlanes(c2)

        '    For c = 1 To gridPlaneCollection.Count
        '        Dim xx As GridPlane = gridPlaneCollection(c)
        '        Dim oGridLine2 As GridLine = xx.CreateGridLine(temp)
        '        result.Add(oGridLine2)
        '    Next c
        'Next c2

        'createGridPlanes = result
        '----------------------------------REWORK ABOVE

        Dim steelSection As steelObject = SP3DColumnObjects(1)
        Dim d1B As New Collection
        Dim d2B As New Collection
        Dim makeCollectiond1B As New Collection                 ' Beams on the Primary beams Y-Plane -- These make the X-planes
        Dim makeCollectiond2B As New Collection                 ' Beams on the Primary beams X-Plane -- These make the Y-planes

        makeCollectiond1B.Add(1)
        makeCollectiond2B.Add(1)

        For c = 2 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)
            Dim res As Integer

            res = isPointOnDualPlane(steelSection.endPoint, temp.endPoint, steelSection.dRotAngle, tolerance)

            If res = 1 And existsInCollection(d1B, "n" & c) = "DNE" And Math.Abs(steelSection.endPoint.Z - temp.endPoint.Z) <= tolerance Then
                d1B.Add(c, "n" & c)
                makeCollectiond1B.Add(c)
            ElseIf res = 2 And existsInCollection(d2B, "n" & c) = "DNE" And Math.Abs(steelSection.endPoint.Z - temp.endPoint.Z) <= tolerance Then
                d2B.Add(c, "n" & c)
                makeCollectiond2B.Add(c)
            End If
        Next c

        Dim result As New Collection
        Dim strPlane As String
        Dim xbase# = SP3DColumnObjects(1).endPoint.X, ybase# = SP3DColumnObjects(1).endPoint.Y
        Dim oSystem As ISystem = New SystemHelper(coordSystem)

        oSystem.AddSystemChild(oCS)

        Dim oGridAxis As New GridAxis(oCS, AxisType.Z)
        Dim gridPlaneCollection As New Collection
        Dim ZPlanes As New Collection

        Dim oElevation As GridElevationPlane = Nothing

        For c = 1 To elevations.Count
            oElevation = New GridElevationPlane(elevations(c) - bottomOfSteel, oGridAxis)
            strPlane = "Elevation_" & c
            oElevation.SetPropertyValue(strPlane, "IJNamedItem", "Name")
            ZPlanes.Add(oElevation)
        Next c

        oGridAxis = New GridAxis(oCS, AxisType.X)

        Dim oGridXPlane As GridPlane = Nothing

        For c = 1 To makeCollectiond1B.Count
            Dim temp As steelObject = SP3DColumnObjects(makeCollectiond1B(c))
            Dim temp2 As Double = getHypotenuseComponent(xbase, temp.startPoint.X, ybase, temp.startPoint.Y, "X")

            oGridXPlane = New GridPlane((temp2), oGridAxis)
            strPlane = "GridX_" & c.ToString()
            oGridXPlane.SetPropertyValue(strPlane, "IJNamedItem", "Name")

            gridPlaneCollection.Add(oGridXPlane)
        Next

        oGridAxis = New GridAxis(oCS, AxisType.Y)

        Dim oGridYPlane As GridPlane = Nothing

        For c = 1 To makeCollectiond2B.Count
            Dim temp As steelObject = SP3DColumnObjects(makeCollectiond2B(c))
            Dim temp2 As Double = getHypotenuseComponent(xbase, temp.startPoint.X, ybase, temp.startPoint.Y, "Y")

            oGridYPlane = New GridPlane((temp2), oGridAxis)
            strPlane = "GridY_" & c.ToString()
            oGridYPlane.SetPropertyValue(strPlane, "IJNamedItem", "Name")

            gridPlaneCollection.Add(oGridYPlane)
        Next

        '---------------------------------------------- Actually create the lines

        For c2 = 1 To ZPlanes.Count
            Dim temp As GridElevationPlane = ZPlanes(c2)

            For c = 1 To gridPlaneCollection.Count
                Dim xx As GridPlane = gridPlaneCollection(c)
                Dim oGridLine2 As GridLine = xx.CreateGridLine(temp)
                result.Add(oGridLine2)
            Next c
        Next c2

        createGridPlanes = result
    End Function

    ''' <summary>
    ''' gets the elevations of floors and steel.  Returns a collection of these elevations and the top and bottom maximal heights.
    ''' </summary>
    ''' <param name="SP3DColumnObjects"></param>
    ''' <param name="SP3DBeamObjects"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getElevations(ByVal SP3DColumnObjects As Collection, ByVal SP3DBeamObjects As Collection) As Collection
        Dim topOfSteel# = -10000, bottomOfSteel# = 10000
        Dim elevations As New Collection
        Dim BBXT# = -200000, BBXB# = 200000, BBYT# = -200000, BBYB# = 200000                      ' To hold bounding box coordinates

        ' Important: modifications to collections will propogate to the parent if there is a "ByRef" directive anywhere in the chain of children.
        ' That said, the collections need to be rotated back after the function is complete/at its end.
        If SP3DColumnObjects(1).dRotAngle <> 0 Then
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DColumnObjects, -SP3DColumnObjects(1).dRotAngle)
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DBeamObjects, -SP3DColumnObjects(1).dRotAngle)
        End If

        For c = 1 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)

            If temp.startPoint.X > BBXT Then
                BBXT = temp.startPoint.X
            ElseIf temp.startPoint.X < BBXB Then
                BBXB = temp.startPoint.X
            End If

            If temp.startPoint.Y > BBYT Then
                BBYT = temp.startPoint.Y
            ElseIf temp.startPoint.Y < BBYB Then
                BBYB = temp.startPoint.Y
            End If
        Next

        For c = 1 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)

            If temp.endPoint.Z > topOfSteel Then
                topOfSteel = temp.endPoint.Z
            ElseIf temp.endPoint.Z < bottomOfSteel Then
                bottomOfSteel = temp.endPoint.Z
            End If

            If temp.startPoint.Z > topOfSteel Then
                topOfSteel = temp.startPoint.Z
            ElseIf temp.startPoint.Z < bottomOfSteel Then
                bottomOfSteel = temp.startPoint.Z
            End If

            If existsInCollection(elevations, "e" & temp.startPoint.Z) = "DNE" Then
                elevations.Add(temp.startPoint.Z, "e" & temp.startPoint.Z)
            End If

            If existsInCollection(elevations, "e" & temp.endPoint.Z) = "DNE" Then
                elevations.Add(temp.endPoint.Z, "e" & temp.endPoint.Z)
            End If
        Next c

        For c = 1 To SP3DBeamObjects.Count
            Dim temp As steelObject = SP3DBeamObjects(c)

            If (temp.startPoint.X >= BBXB And temp.startPoint.X <= BBXT And temp.startPoint.Y <= BBYT And temp.startPoint.Y >= BBYB) Or (temp.endPoint.X >= BBXB And temp.endPoint.X <= BBXT And temp.endPoint.Y <= BBYT And temp.endPoint.Y >= BBYB) Then
                If existsInCollection(elevations, "e" & temp.startPoint.Z) = "DNE" And temp.startPoint.Z < topOfSteel And temp.startPoint.Z > bottomOfSteel Then
                    elevations.Add(temp.startPoint.Z, "e" & temp.startPoint.Z)
                End If

                If existsInCollection(elevations, "e" & temp.endPoint.Z) = "DNE" And temp.endPoint.Z < topOfSteel And temp.endPoint.Z > bottomOfSteel Then
                    elevations.Add(temp.endPoint.Z, "e" & temp.endPoint.Z)
                End If
            End If
        Next c

        If SP3DColumnObjects(1).dRotAngle <> 0 Then
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DColumnObjects, SP3DColumnObjects(1).dRotAngle)
            rotateStructure(SP3DColumnObjects(1).startPoint, SP3DBeamObjects, SP3DColumnObjects(1).dRotAngle)
        End If

        elevations.Add(topOfSteel, "topOfSteel")
        elevations.Add(bottomOfSteel, "bottomOfSteel")

        getElevations = elevations
    End Function

    ''' <summary>
    ''' Removes items form a steel collection, based on their start and end points
    ''' </summary>
    ''' <param name="totalSteelCollection"></param>
    ''' <param name="itemsToRemove"></param>
    ''' <remarks></remarks>
    Public Sub removeSteelObjectsFromCollection(ByRef totalSteelCollection As Collection, ByVal itemsToRemove As Collection)
        Dim c As Integer
        Dim c2 As Integer

        For c = 1 To itemsToRemove.Count
            For c2 = 1 To totalSteelCollection.Count
                If totalSteelCollection(c2).startPoint.X = itemsToRemove(c).startPoint.X And totalSteelCollection(c2).startPoint.Y = itemsToRemove(c).startPoint.Y And totalSteelCollection(c2).startPoint.Z = itemsToRemove(c).startPoint.Z And (totalSteelCollection(c2).endPoint.X = itemsToRemove(c).endPoint.X And totalSteelCollection(c2).endPoint.Y = itemsToRemove(c).endPoint.Y And totalSteelCollection(c2).endPoint.Z = itemsToRemove(c).endPoint.Z) Then
                    totalSteelCollection.Remove(c2)

                    Exit For
                End If
            Next c2
        Next c
    End Sub

    ''' <summary>
    ''' Returns the, looking down, the column in the top right corner with an emphasis on the top 
    ''' </summary>
    ''' <param name="SP3DColumnObjects"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function get2DBoundingBox(ByVal colCollection As Collection, ByVal SP3DColumnObjects As Collection) As twoDBoundingBox
        Dim result As New twoDBoundingBox

        For c = 1 To SP3DColumnObjects.Count
            Dim temp As steelObject = colCollection(SP3DColumnObjects(c))

            If temp.startPoint.X > result.topRight.X Then
                result.topRight.X = temp.startPoint.X
            ElseIf temp.startPoint.X < result.bottomLeft.X Then
                result.bottomLeft.X = temp.startPoint.X
            End If

            If temp.startPoint.Y > result.topRight.Y Then
                result.topRight.Y = temp.startPoint.Y
            ElseIf temp.startPoint.Y < result.bottomLeft.Y Then
                result.bottomLeft.Y = temp.startPoint.Y
            End If
        Next c

        get2DBoundingBox = result
    End Function

    Public Function getColumnsInBoundingBox(ByVal SP3DColumnObjects As Collection, ByVal BoundingBox As twoDBoundingBox) As Collection
        Dim c As Integer
        Dim result As New Collection

        For c = 1 To SP3DColumnObjects.Count
            Dim temp As steelObject = SP3DColumnObjects(c)

            If temp.startPoint.X <= BoundingBox.topRight.X And temp.startPoint.X >= BoundingBox.bottomLeft.X And temp.startPoint.Y <= BoundingBox.topRight.Y And temp.startPoint.Y >= BoundingBox.bottomLeft.Y Then
                result.Add(temp)
            End If
        Next c

        getColumnsInBoundingBox = result
    End Function
End Module



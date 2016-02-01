Imports System.Text.RegularExpressions
Imports Ingr.SP3D.Common.Middle

Module CIMUtilities
    ''' <summary>
    ''' Reads the specified file and returns a collection of collections keyed by their directive #ID -- directives, topLevelDirectives, itemStandards, itemAssignment
    ''' </summary>
    ''' <param name="fileName">Path to the file to read</param>
    ''' <returns>A Collection of collections</returns>
    ''' <remarks>Designed for CIM/2 step files.
    ''' The 'If tempString = "design_part".....' line defines the top level directives -- Designed Assembly was removed as this indicated a connection point
    ''' </remarks>
    Public Function loadDirectives(ByVal fileName As String) As Collection
        Dim directives As New Collection
        Dim topLevelDirectives As New Collection
        Dim itemStandards As New Collection
        Dim itemAssignment As New Collection
        Dim sectionProfiles As New Collection
        Dim globalContexts As New Collection

        Dim objReader As New System.IO.StreamReader(fileName)
        Dim start As Boolean = False
        Dim tempString As String
        Dim finalString As String
        Dim directive As String
        Dim pos As Integer

        While start = False And Not objReader.EndOfStream
            tempString = objReader.ReadLine

            If tempString = "DATA;" Then
                start = True
            End If
        End While

        start = False

        Try
            While Not objReader.EndOfStream
                tempString = objReader.ReadLine

                If tempString = "ENDSEC;" Then
                    Exit While
                End If

                If start = False Then
                    directive = tempString.Substring(0, InStr(tempString, "=") - 1)
                    pos = InStr(tempString, "=") + 1

                    If Mid(tempString, tempString.Count, 1) = ";" Then
                        finalString = tempString.Substring(pos - 1, tempString.Count - 1 - pos)
                    Else
                        start = True
                        finalString = Mid(tempString, pos)
                    End If
                Else
                    If Mid(tempString, tempString.Count, 1) = ";" Then
                        start = False
                        finalString = finalString & tempString.Substring(0, tempString.Count - 1)
                    Else
                        finalString = finalString & tempString
                    End If
                End If

                If start = False Then
                    directives.Add(LTrim(finalString.Replace(vbCrLf, "")), directive)

                    '******* Create speed up collections --- This will save a lot of time in the long run.
                    tempString = LCase(getWord(directives(directive)))

                    If tempString = "design_part" Or tempString = "located_part" Or tempString = "managed_data_item" Then
                        topLevelDirectives.Add(directive)
                    End If

                    If tempString = "item_reference_assigned" Then
                        itemAssignment.Add(directive)
                    End If

                    If tempString = "section_profile" Then      ' Making this incase the item reference standards do not exist
                        sectionProfiles.Add(directive)
                    End If

                    If tempString = "global_unit_assigned_context" Or (tempString = "geometric_representation_context") Then ' And InStr(tempString, "global_unit_assigned_context") > 0) Then
                        globalContexts.Add(directive)
                    End If
                End If

            End While
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Dim collectionCollection As New Collection

        collectionCollection.Add(directives, "directives")
        collectionCollection.Add(topLevelDirectives, "topLevelDirectives")
        collectionCollection.Add(itemAssignment, "itemAssignment")
        collectionCollection.Add(globalContexts, "globalContexts")
        collectionCollection.Add(sectionProfiles, "sectionProfiles")

        loadDirectives = collectionCollection
    End Function

    ''' <summary>
    ''' Gets the unit measure (from the measures class) for global units such as coordinate systems
    ''' </summary>
    ''' <param name="contextCollection"></param>
    ''' <param name="directives"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getGlobalContext(ByVal contextCollection As Collection, ByRef directives As Collection) As Double
        Dim c As Integer
        Dim c2 As Integer
        Dim currentString As String
        Dim result As Double = -1
        Dim word As String = ""
        Dim meas As New measures

        For c = 1 To contextCollection.Count
            Dim tempString() As String
            Dim tempDirs() As String

            currentString = LCase(directives(contextCollection(c)))

            While getWord(currentString) <> "global_unit_assigned_context" And currentString <> "-1"
                currentString = advanceToNextWord(currentString)
            End While

            If currentString <> "-1" Then

                tempString = getbracketElements(currentString)
                tempString(0) = Replace(tempString(0), "(", "")
                tempString(0) = Replace(tempString(0), ")", "")
                tempDirs = Split(tempString(0), ",")

                For c2 = 0 To tempDirs.Count - 1
                    If Trim(tempDirs(c2)) <> "" Then
                        currentString = LCase(directives(tempDirs(c2)))

                        If getWord(currentString) = "context_dependent_unit" Then

                            Dim xx() As String = getbracketElements(currentString)

                            If xx(0) = "inch" Or xx(0) = "foot" Or xx(0) = "centimetre" Or xx(0) = "millimetre" Or xx(0) = "metre" Or xx(0) = "yard" Then
                                result = meas.unitMeasures(xx(0))

                                Exit For
                            End If
                        Else
                            currentString = Mid(currentString, InStr(currentString, "si_unit"))
                            currentString = getBracketContent(currentString)

                            currentString = currentString.Replace(",", "")
                            currentString = currentString.Replace(".", "")

                            If currentString = "inch" Or currentString = "foot" Or currentString = "centimetre" Or currentString = "millimetre" Or currentString = "metre" Or currentString = "yard" Then

                                result = meas.unitMeasures(currentString)

                                Exit For
                            End If
                        End If

                    End If

                Next c2

                If result <> -1 Then
                    Exit For
                End If

            End If
        Next c

        getGlobalContext = result
    End Function

    ''' <summary>
    ''' Retrieves the directive numbers in a string and returns them as a collection
    ''' </summary>
    ''' <param name="lineToParse"></param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function getSubDirectives(ByVal lineToParse As String) As Collection
        Dim c%, c2%
        Dim result As New Collection
        Dim temp() As String

        temp = Split(lineToParse, "#")

        For c = 1 To temp.Count - 1
            For c2 = 1 To temp(c).Count
                If Not IsNumeric(Mid(temp(c), c2, 1)) Then
                    Exit For
                End If
            Next c2

            result.Add(Mid(temp(c), 1, c2 - 1))
        Next c

        getSubDirectives = result
    End Function

    ''' <summary>
    ''' Advances to the next key word token in a line of bracket seperated directives/tokens.  Returns -1 if there are no more words to advance to.
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function advanceToNextWord(ByVal line As String) As String
        Dim x As Integer = 0
        Dim c As Integer

        While Mid(line, 1, 1) = "("
            line = Mid(line, 2)
        End While

        c = InStr(line, "(")

        If c = 0 Then
            line = "-1"
        Else
            While x >= 0 And c < line.Count
                If line(c) = "(" Then
                    x = x + 1
                ElseIf line(c) = ")" Then
                    x = x - 1
                End If
                c = c + 1
            End While
        End If

        If c >= line.Count Then
            line = "-1"
        Else
            line = Mid(line, c + 1)
        End If

        advanceToNextWord = line
    End Function

    ''' <summary>
    ''' Makes the steel standards and compares them to the existing catalog.
    ''' </summary>
    ''' <param name="itemAssignment"></param>
    ''' <param name="sectionProfiles"></param>
    ''' <param name="directives"></param>
    ''' <param name="programName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function makeStandards2(ByVal itemAssignment As Collection, ByVal directives As Collection, ByVal programName As String, Optional ByVal mode As Short = 1) As Collection
        Dim finalCollection As New Collection
        Dim c As Integer
        Dim query As String

        For c = 1 To itemAssignment.Count
            Dim result As Collection
            Dim temp As New item_reference_standard
            Dim itemRefStd() As String
            Dim itemRefSourceStd() As String
            Dim card() As String
            Dim itemTag As String

            If mode = 1 Then
                Dim tempCollection As Collection
                tempCollection = getSubDirectives(directives(itemAssignment(c)))
                itemRefStd = getbracketElements(directives("#" & tempCollection(1)))
                itemRefSourceStd = getbracketElements(directives(itemRefStd(1)))
                card = getbracketElements(directives("#" & tempCollection(2)))

                temp.standardOrgName = itemRefSourceStd(0)
                temp.standardName = itemRefSourceStd(1)
                temp.year = itemRefSourceStd(2)
                temp.version = itemRefSourceStd(3)
                temp.sectionName = itemRefStd(0)

                If card.Count > 4 Then
                    temp.cardinalPoint = card(4)
                End If

                itemTag = "#" & tempCollection(2)
            ElseIf mode = 2 Then        ' No proper steel definition, just a section name.
                Dim tempCollection() As String
                tempCollection = getbracketElements(directives(itemAssignment(c)))
                temp.sectionName = tempCollection(1)
                temp.cardinalPoint = 2
                itemTag = itemAssignment(c)
            End If

            If programName.ToUpper = "TEKLA" Then
                query = "SELECT * FROM SP3DSteelSections, SteelTable_STAAD WHERE SteelTable_STAAD.SectionName LIKE '" & temp.sectionName & "' AND SteelTable_STAAD.SPID = SP3DSteelSections.SPID"
                result = retrieveFromDBtoCollection(query)
            ElseIf programName.ToUpper = "STAAD" Then
                query = "SELECT * FROM SP3DSteelSections, SteelTable_Tekla WHERE SteelTable_Tekla.SectionName LIKE '" & temp.sectionName & "' AND SteelTable_Tekla.SPID = SP3DSteelSections.SPID"
                result = retrieveFromDBtoCollection(query)
            End If

            If collectionCount(result) < 1 Then
                If temp.standardName <> "" Then
                    query = "SELECT * FROM SP3DSteelSections WHERE SP3DSchedule LIKE '" & temp.standardName & "' AND SP3DSectionName LIKE '" & temp.sectionName & "'"
                    result = retrieveFromDBtoCollection(query)
                End If

                If collectionCount(result) < 1 Then
                    query = "SELECT * FROM SP3DSteelSections WHERE SP3DSectionName LIKE '" & temp.sectionName & "'"
                    result = retrieveFromDBtoCollection(query)
                End If

                If collectionCount(result) < 1 Then
                    query = "SELECT * FROM SP3DSteelSections WHERE EDIName LIKE '" & temp.sectionName & "'"
                    result = retrieveFromDBtoCollection(query)
                End If

                If collectionCount(result) < 1 Then
                    query = "SELECT * FROM SP3DSteelSections, SteelTable_STAAD WHERE SteelTable_STAAD.SectionName LIKE '" & temp.sectionName & "' AND SteelTable_STAAD.SPID = SP3DSteelSections.SPID"
                    result = retrieveFromDBtoCollection(query)
                End If

                If collectionCount(result) < 1 Then
                    query = "SELECT * FROM SP3DSteelSections, SteelTable_Tekla WHERE SteelTable_Tekla.SectionName LIKE '" & temp.sectionName & "' AND SteelTable_Tekla.SPID = SP3DSteelSections.SPID"
                    result = retrieveFromDBtoCollection(query)
                End If
            End If

            If collectionCount(result) < 1 Then
                temp.standardName = "Unknown Section"
            Else
                Dim tempSplit() As String

                tempSplit = Split(result(1), "||")

                temp.standardName = Trim(tempSplit(4))
                temp.sectionName = Trim(tempSplit(2))
                temp.sectionType = Trim(tempSplit(3))

                'If checkSteelInDB(temp.standardName, temp.sectionType, temp.sectionName) = False Then
                '    temp.standardName = "Section Not In Catalog"
                'End If
            End If

            finalCollection.Add(temp, itemTag)

            result = Nothing
        Next

        makeStandards2 = finalCollection
    End Function

    ''' <summary>
    ''' Seperates out items located within the first set of brackets, seperated by commas.  Sub-Bracketed items will be returned as an element of the array.
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns>Array of strings</returns>
    ''' <remarks></remarks>
    Public Function getbracketElements(ByVal line As String) As String()
        Dim x As Integer = 0
        Dim c As Integer
        Dim c2 As Integer = 0
        Dim contents(0 To 0) As String

        While Mid(line, 1, 1) = "("
            line = Mid(line, 2)
        End While

        line = Mid(line, InStr(line, "(") + 1).Replace("'", "")

        For c = 0 To line.Count - 1
            If line(c) = "(" Then
                x = x + 1
            ElseIf line(c) = ")" Then
                x = x - 1

                If x = -1 Then
                    line = Mid(line, 1, c)

                    Exit For
                End If
            End If

            If line(c) = "," And x = 0 Then
                c2 = c2 + 1

                ReDim Preserve contents(0 To c2)
            End If

            If (line(c) <> "," And x = 0) Or (x > 0) Then
                contents(c2) = contents(c2) & line(c)
            End If
        Next c

        getbracketElements = contents
    End Function

    ''' <summary>
    ''' Returns a string with the complete contents of the first set of brackets encountered.
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getBracketContent(ByVal line As String) As String
        Dim temp As String

        temp = Mid(line, InStr(line, "(") + 1, (InStr(line, ")")) - (InStr(line, "(") + 1))

        getBracketContent = temp
    End Function

    ''' <summary>
    ''' Parses out and returns the directive key word for the line
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns>The directive word only</returns>
    ''' <remarks></remarks>
    Public Function getWord(ByVal line As String) As String
        Dim result As String = "-1"

        While Mid(line, 1, 1) = "("
            line = Mid(line, 2)
        End While

        If InStr(line, "(") > 0 Then
            result = Mid(line, 1, InStr(line, "(") - 1)
        End If

        getWord = result
    End Function
End Module

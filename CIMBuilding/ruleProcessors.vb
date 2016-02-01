''' <summary>
''' A Module file containing the processors for the individual directives -- Symantec and Syntactic processors
''' </summary>
''' <remarks>ISSUES: - Need to parse past the Part Prismatics to get the steel types for labelling purposes
'''                  - Part Prismatic also has sub-types.....have not seen them used yet though
'''                  - Context Dependent Units -- Possibly need to upgrade this
'''                  - Need to add Partsheet for Parts definition -- Check if the object "isnot nothing" for now -- This is for slabs
'''</remarks>

Module ruleProcessors

    '******************************** TERMINATORS
    ''' <summary>
    ''' Cartesian_Point
    ''' </summary>
    ''' <param name="lineString">The whole string for the line</param>
    ''' <returns>an object of type "Cartesian_Point"</returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC003(ByVal lineString As String) As cartesian_point
        Dim result As New cartesian_point
        Dim argv() As String = getbracketElements(lineString)

        result.label = argv(0)

        If argv.Count < 4 Then
            Dim argv2() = getbracketElements(argv(1))

            result.X = CDbl(argv2(0))
            result.Y = CDbl(argv2(1))
            result.Z = CDbl(argv2(2))
        Else
            result.X = CDbl(argv(1))
            result.Y = CDbl(argv(2))
            result.Z = CDbl(argv(3))
        End If

        LPM5_CC003 = result
    End Function

    ''' <summary>
    ''' Direction
    ''' </summary>
    ''' <param name="lineString">The whole string for the line</param>
    ''' <returns>an object of type "Direction"</returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC032(ByVal lineString As String) As direction
        Dim result As New direction
        Dim argv() As String = getbracketElements(lineString)

        result.label = argv(0)

        If argv.Count < 4 Then
            Dim argv2() = getbracketElements(argv(1))

            result.X = CDbl(argv2(0))
            result.Y = CDbl(argv2(1))
            result.Z = CDbl(argv2(2))
        Else
            result.X = CDbl(argv(1))
            result.Y = CDbl(argv(2))
            result.Z = CDbl(argv(3))
        End If

        LPM5_CC032 = result
    End Function

    ''' <summary>
    ''' Functional_Role
    ''' </summary>
    ''' <param name="lineString">The whole string for the line</param>
    ''' <returns>an object of type "Functional_Role"</returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC122(ByVal lineString As String) As functional_role
        Dim result As New functional_role
        Dim temp As String
        Dim temp2() As String

        temp = Mid(lineString, InStr(lineString, "(")).Replace("(", "")
        temp2 = Split(temp.Replace(")", ""), ",")

        result.name = temp2(0).Replace("'", "")
        result.description = temp2(1).Replace("'", "")

        LPM5_CC122 = result
    End Function

    ''' <summary>
    ''' SI_Unit -- Returns a Double to multiply with a measure to put it in SP3D native units
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC034(ByVal lineString As String) As Double
        Dim temp As String = getBracketContent(lineString)
        Dim x As New measures

        temp = temp.Replace(".", "")
        temp = temp.Replace(",", "")
        temp = temp.Replace("$", "")

        LPM5_CC034 = x.unitMeasures(LCase(temp))
    End Function

    ''' <summary>
    ''' Positive_Length_Measure
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <returns>A double</returns>
    ''' <remarks>Since this should only ever have a float value, we can assume the element content return</remarks>
    Public Function LPM5_CC014X1(ByVal lineString As String) As Double
        Dim result() As String = getbracketElements(lineString)

        LPM5_CC014X1 = CDbl(result(0))
    End Function

    ''' <summary>
    ''' Context_Depnedent_Unit -- Possibly needs to be updated to handle the dimensional exponents
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks>Need to upgrade this later so it can be used for labelling purposes -- Dimensional exponents seem to be a lookup table, ignoring at the moment.</remarks>
    Public Function LPM5_CC030(ByVal lineString As String, ByRef directiveCollection As Collection) As Object
        Dim temp As New measures
        Dim unit() As String

        unit = getbracketElements(lineString)

        If unit(0) = "" Then
            unit(0) = "inch"
        End If

        LPM5_CC030 = temp.unitMeasures(LCase(unit(0)))
    End Function

    ''' <summary>
    ''' Assembly_Design_Structural_Member_Linear
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC110(ByVal lineString As String) As String
        Dim result() As String

        result = getbracketElements(lineString)

        LPM5_CC110 = result(1)
    End Function

    '******************************** NON-TERMINATORS
    ''' <summary>
    ''' Coord_System_Cartesian_3D
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC310(ByVal lineString As String, ByRef directiveCollection As Collection) As coord_system_cartesian_3d
        Dim argv() As String = getbracketElements(lineString)
        Dim result As New coord_system_cartesian_3d

        result.coordSystemName = argv(0)
        result.coordSystemUse = argv(1)
        result.sign = argv(2)
        result.dimensions = CInt(argv(3))
        result.axisSystem = analyze(directiveCollection(argv(4)), directiveCollection)

        LPM5_CC310 = result
    End Function

    ''' <summary>
    ''' Axis2_Placement_3D
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC024(ByVal lineString As String, ByRef directiveCollection As Collection) As axis2_placement_3d
        Dim argv() As String = getbracketElements(lineString)
        Dim result As New axis2_placement_3d

        result.axisName = argv(0)
        result.cartPoint = analyze(directiveCollection(argv(1)))
        result.direction1 = analyze(directiveCollection(argv(2)))
        result.direction2 = analyze(directiveCollection(argv(3)))

        LPM5_CC024 = result
    End Function

    ''' <summary>
    ''' Coord_System
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC310X(ByVal lineString As String, ByRef directiveCollection As Collection) As coord_system
        Dim result As New coord_system
        Dim argv() As String = getbracketElements(lineString)

        result.coordSystemName = argv(0)
        result.coordSystemUse = argv(1)
        result.sign = argv(2)
        result.dimensions = CInt(argv(3))

        lineString = advanceToNextWord(lineString)

        result.parentSystem = analyze(directiveCollection(getBracketContent(lineString)), directiveCollection)

        lineString = advanceToNextWord(lineString)

        If lineString <> "-1" Then
            result.childSystem = analyze(directiveCollection(getBracketContent(lineString)), directiveCollection)
        End If

        LPM5_CC310X = result
    End Function

    ''' <summary>
    ''' Named_Unit
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <returns></returns>
    ''' <remarks>Have to check for the "*" as this can extend the class to a different type --> different return types, so we use an object</remarks>
    Public Function LPM5_CC034X(ByVal lineString As String, ByRef directiveCollection As Collection) As Object
        Dim temp As String

        temp = getBracketContent(lineString)

        If temp = "*" Then
            lineString = advanceToNextWord(lineString)

            Return analyze(lineString)
        Else
            Return analyze(directiveCollection(temp))
        End If
    End Function

    ''' <summary>
    ''' length_unit
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks>This is a place holder (in the specs)</remarks>
    Public Function LPM5_CC014X2(ByVal lineString As String, ByRef directiveCollection As Collection) As Object
        lineString = advanceToNextWord(lineString)

        Return analyze(lineString, directiveCollection)
    End Function

    ''' <summary>
    ''' positive_length_measure_with_unit
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks>Place holder to focus the parent</remarks>
    Public Function LPM5_CC014(ByVal lineString As String, ByRef directiveCollection As Collection) As Double
        Dim argv() As String
        Dim result As Double

        argv = getbracketElements(lineString)

        If InStr(argv(0), "LENGTH_MEASURE") > 0 Then            ' Majority case, but the specification says it can be in reverse order (the arguments).
            result = analyze(argv(0), directiveCollection)
            result = result * CDbl(analyze(directiveCollection(argv(1)), directiveCollection))
        Else
            result = analyze(argv(1), directiveCollection)
            result = result * CDbl(analyze(directiveCollection(argv(0)), directiveCollection))
        End If

        LPM5_CC014 = result
    End Function


    ''' <summary>
    ''' Part
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC248(ByVal lineString As String, ByRef directiveCollection As Collection) As part
        Dim result As New part
        Dim IDs() As String

        'While lineString(0) = "("               '' in case of leading brackets, remove them....they can cause issues and are unnessecary
        '    lineString = Mid(lineString, 2)
        'End While

        IDs = getbracketElements(lineString)

        result.fabrication_method = IDs(0).Replace(".", "")
        result.manufacturers_ref = IDs(1).Replace(".", "")
        result.manufacturers_ref = result.manufacturers_ref.Replace("$", "")

        lineString = advanceToNextWord(lineString)

        result.partTypeClass = analyze(lineString, directiveCollection)

        LPM5_CC248 = result
    End Function

    ''' <summary>
    ''' Part_Prismatic -- A focusing agent; Advances to the next word
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC249X(ByVal lineString As String, ByRef directiveCollection As Collection) As Object
        lineString = advanceToNextWord(lineString)

        LPM5_CC249X = analyze(lineString, directiveCollection)
    End Function

    ''' <summary>
    ''' Part_Prismatic_Simple -- Needs Upgrade, See Remarks
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks>There are more unaccounted derivatives to account for.  Also, has a curved part possability</remarks>
    Public Function LPM5_CC249(ByVal lineString As String, ByRef directiveCollection As Collection) As Object
        On Error GoTo er

        Dim prisPart As New part_prismatic_simple
        Dim argv() As String = getbracketElements(lineString)
        Dim standardCollection As Collection = directiveCollection("FinalStandards")
        Dim itemRef As item_reference_standard

        If argv.Count > 4 Then      ' Rare case of a scoped prismatic, but we need to allow for it.

            prisPart.strength = argv(1)
            prisPart.madeFrom = argv(4).Replace(".", "")
            itemRef = standardCollection(argv(6))
            prisPart.sectionToUse = itemRef

            prisPart.length = analyze(directiveCollection(argv(7)), directiveCollection)
        Else
            'MsgBox(standardCollection.Count)
            If LCase(getWord(directiveCollection(argv(0)))) = "section_profile" Then
                'MsgBox(41)
                itemRef = standardCollection(argv(0))
                'MsgBox(42)
                prisPart.sectionToUse = itemRef
                'MsgBox(43)
                prisPart.length = analyze(directiveCollection(argv(1)), directiveCollection)
                'MsgBox(44)
            Else
                prisPart.sectionToUse = "slab"
            End If


        End If

        LPM5_CC249 = prisPart
er:
        Dim FILE_NAME As String = "C:\Documents and Settings\stded2\Desktop\errorLog.txt"
        Dim errorLine As String

        errorLine = "Rule LPM5_CC249 -- " & lineString

        Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)
        objWriter.WriteLine(errorLine)
        objWriter.Close()
    End Function


    ''' <summary>
    ''' located_assembly
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC312(ByVal lineString As String, ByRef directiveCollection As Collection) As located_assembly
        Dim argv() As String = getbracketElements(lineString.Replace("$", ""))
        Dim result As New located_assembly

        result.itemNumber = argv(0)
        result.itemName = argv(1)
        result.itemDescription = argv(2)
        result.location = analyze(directiveCollection(argv(3)), directiveCollection)

        If Trim(argv(4)) <> "" Then
            Dim temp As New Collection
            Dim dirs() As String
            Dim c As Integer

            argv(4) = Replace(argv(4), "(", "")
            argv(4) = Trim(Replace(argv(4), ")", ""))

            dirs = Split(argv(4), ",")

            For c = 1 To dirs.Count
                If dirs(c) <> "" Then
                    temp.Add(analyze(directiveCollection(dirs(c)), directiveCollection))
                End If
            Next c

            result.locationOnGrid = temp
        End If

        result.descriptiveAssembly = analyze(directiveCollection(argv(5)), directiveCollection)
        result.parentStructure = analyze(directiveCollection(argv(6)), directiveCollection)

        LPM5_CC312 = result
    End Function

    ''' <summary>
    ''' Located Structure
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC318(ByVal lineString As String, ByRef directiveCollection As Collection) As located_assembly
        Dim argv() As String = getbracketElements(lineString.Replace("$", ""))
        Dim result As New located_assembly

        result.itemNumber = argv(0)
        result.itemName = argv(1)
        result.itemDescription = argv(2)
        result.location = analyze(directiveCollection(argv(3)), directiveCollection)
        result.location = analyze(directiveCollection(argv(4)), directiveCollection)
        result.location = analyze(directiveCollection(argv(5)), directiveCollection)

        LPM5_CC318 = result
    End Function

    ''' <summary>
    ''' Assembly Manufacturing
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC112(ByVal lineString As String) As assembly_manufacturing
        Dim argv() As String = getbracketElements(lineString.Replace("$", ""))
        Dim result As New assembly_manufacturing

        result.itemNumber = argv(0)
        result.itemName = argv(1)
        result.itemDescription = argv(2)
        result.lifeCycleStage = argv(3)
        result.assemblySequenceNumber = argv(4)
        result.complexity = argv(5)
        result.surfaceTreatment = argv(6)
        result.assemblySequence = argv(7)
        result.assemblyUse = argv(8)
        result.placeOfAssembly = argv(9)

        LPM5_CC112 = result
    End Function

    ''' <summary>
    ''' Structure
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC129(ByVal lineString As String) As structureDef
        Dim argv() As String = getbracketElements(lineString.Replace("$", ""))
        Dim result As New structureDef

        result.itemNumber = argv(0)
        result.itemName = argv(1)
        result.itemDescription = argv(2)

        LPM5_CC129 = result
    End Function
    '******************************** STARTERS

    ''' <summary>
    ''' Design_Part
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC118(ByVal lineString As String, ByRef directiveCollection As Collection) As design_part
        Dim argv() As String = getbracketElements(lineString)
        Dim result As New design_part
        Dim parents() As String
        Dim locations() As String
        Dim c As Integer
        Dim tp As New Collection
        Dim tl As New Collection

        result.label = argv(0)
        result.partDefinition = analyze(directiveCollection(argv(1)), directiveCollection)

        argv(2) = argv(2).Replace("(", "")
        argv(2) = argv(2).Replace(")", "")
        argv(3) = argv(3).Replace("(", "")
        argv(3) = argv(3).Replace(")", "")

        parents = Split(argv(2), ",")
        locations = Split(argv(3), ",")

        For c = 0 To parents.Count - 1
            If parents(c) IsNot Nothing Then
                tl.Add(analyze(directiveCollection(locations(c)), directiveCollection))
                tp.Add(analyze(directiveCollection(parents(c))))
            End If
        Next

        result.location = tl
        result.parentAssembly = tp

        LPM5_CC118 = result
    End Function

    ''' <summary>
    ''' located_part
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC179(ByVal lineString As String, ByRef directiveCollection As Collection) As located_part
        Dim argv() As String = getbracketElements(lineString)
        Dim result As New located_part

        result.item_number = argv(0)
        result.item_name = argv(1)
        result.item_description = argv(2)
        result.location = analyze(directiveCollection(argv(3)), directiveCollection)
        result.partType = analyze(directiveCollection(argv(4)), directiveCollection)
        result.parent_assembly = analyze(directiveCollection(argv(5)), directiveCollection)

        LPM5_CC179 = result
    End Function









    '********************* NEED TO DEAL WITH THIS MONSTROSITY **************************************'
    ''' <summary>
    ''' managed_data_item
    ''' </summary>
    ''' <param name="lineString"></param>
    ''' <param name="directiveCollection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LPM5_CC100(ByVal lineString As String, ByRef directiveCollection As Collection) As design_part
        Dim argv() As String = getbracketElements(lineString)


    End Function
End Module

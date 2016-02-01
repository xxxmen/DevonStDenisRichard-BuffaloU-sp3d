Public Module lexicalAnalyzer

    ' The base lexical analyzer -- While it is bad practice to pass the directive collection as a pointer, in this case it saves a ton of
    '                              resources and time (from not copying the collection).
    ' Implement ASSEMBLY_DESIGN_STRUCTURAL_MEMBER_LINEAR_COLUMN


    Public Function analyze(ByVal line As String, Optional ByRef directiveCollection As Collection = Nothing) As Object
        Dim returnValue As Object

        While line(0) = "("               '' in case of leading brackets, remove them....they can cause issues and are unnessecary
            line = Mid(line, 2)
        End While
        'MsgBox(line)
        Select Case LCase(getWord(line))            '******************* START TERMINATORS
            Case "cartesian_point"
                returnValue = LPM5_CC003(line)
            Case "direction"
                returnValue = LPM5_CC032(line)
            Case "functional_role"
                returnValue = LPM5_CC122(line)
            Case "si_unit"
                returnValue = LPM5_CC034(line)
            Case "assembly_design_structural_member_linear"
                returnValue = LPM5_CC110(line)

            Case "coord_system_cartesian_3d"        '******************* START MID-LEVEL
                returnValue = LPM5_CC310(line, directiveCollection)
            Case "axis2_placement_3d"
                returnValue = LPM5_CC024(line, directiveCollection)
            Case "coord_system"
                returnValue = LPM5_CC310X(line, directiveCollection)

                'START Parts
            Case "part"
                returnValue = LPM5_CC248(line, directiveCollection)
            Case "part_prismatic"                   ' Does not exist in the schema entity specs, so it gets an X class
                returnValue = LPM5_CC249X(line, directiveCollection)
            Case "part_prismatic_simple"
                returnValue = LPM5_CC249(line, directiveCollection)
                'END Parts

                'START length units
            Case "positive_length_measure_with_unit"
                returnValue = LPM5_CC014(line, directiveCollection)
            Case "positive_length_measure"
                returnValue = LPM5_CC014X1(line)
            Case "length_unit"
                returnValue = LPM5_CC014X2(line, directiveCollection)
            Case "named_unit"
                returnValue = LPM5_CC034X(line, directiveCollection)
            Case "context_dependent_unit"
                returnValue = LPM5_CC030(line, directiveCollection)
                'END length units

            Case "located_assembly"
                returnValue = LPM5_CC312(line, directiveCollection)
            Case "assembly_manufacturing"
                returnValue = LPM5_CC112(line)
            Case "structure"
                returnValue = LPM5_CC129(line)
            Case "located_structure"
                returnValue = LPM5_CC318(line, directiveCollection)


            Case "design_part"                      '******************** TOP LEVEL
                returnValue = LPM5_CC118(line, directiveCollection)         ' SP3D & Bently Method
            Case "located_part"
                returnValue = LPM5_CC179(line, directiveCollection)         ' Tekla Method
            Case "managed_data_item"
                returnValue = LPM5_CC100(line, directiveCollection)         ' STAAD Method
            Case Else
                'MsgBox("Unknown Directive: " & LCase(getWord(line)))
        End Select

        analyze = returnValue
    End Function


End Module

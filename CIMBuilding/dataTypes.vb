Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Structure.Middle

'*****************************************************************************************************************************************
'* Data Class Types Library v1.0
'*
'* Assumptions: - Since we are working in SP3D, only 3D systems will be valid.  This implies that only "axis2_placement_3d" will
'*                normally be used.  If more types are supported, there will need to be an encapsulater to ID the various coord
'*                systems.
'*
'*
'*****************************************************************************************************************************************


'************************************* BOTTOM LEVEL
Public Class cartesian_point
    Public label As String
    Public X As Double
    Public Y As Double
    Public Z As Double
End Class
Public Class direction
    Public label As String
    Public X As Double
    Public Y As Double
    Public Z As Double
End Class
Public Class functional_role
    Public name As String
    Public description As String
End Class

'************************************* MID LEVEL

' Part class objects
Public Class part
    Public fabrication_method As String
    Public manufacturers_ref As String
    Public partTypeClass As Object          ' Can be derived from the ".GetType" method
End Class

Public Class part_prismatic_simple
    Public strength As String
    Public madeFrom As String
    Public sectionToUse As Object 'item_reference_standard -- There may also be a curve
    Public length As Double
End Class
Public Class part_prismatic_complex
    Public fabricationType As String
    Public referenceText As String

End Class   'incomplete here down
Public Class part_sheet_bonded
    Public fabricationType As String
    Public referenceText As String

End Class
Public Class part_sheet_profiled
    Public fabricationType As String
    Public referenceText As String

End Class
Public Class part_complex
    Public fabricationType As String
    Public referenceText As String
End Class
'end part classes

''' <summary>
''' This is an abstract that allows for a child coordinate system based off of its parent.  The child system is optional.
''' </summary>
''' <remarks></remarks>
Public Class coord_system
    Public coordSystemName As String
    Public coordSystemUse As String
    Public sign As String
    Public dimensions As Integer
    Public parentSystem As Object
    Public childSystem As Object
End Class

Public Class coord_system_cartesian_3d
    Public coordSystemName As String
    Public coordSystemUse As String
    Public sign As String
    Public dimensions As Integer
    Public axisSystem As axis2_placement_3d
End Class

Public Class axis2_placement_3d
    Public axisName As String
    Public cartPoint As cartesian_point
    Public direction1 As direction
    Public direction2 As direction
End Class

Public Class located_assembly
    Public itemNumber As String                 ' Should be an integer value
    Public itemName As String
    Public itemDescription As String
    Public location As Object
    Public locationOnGrid As Collection
    Public descriptiveAssembly As Object        ' Should be either ASSEMBLY_DESIGN_STRUCTURAL_CONNECTION_INTERNAL or ASSEMBLY_MANUFACTURING in most cases
    Public parentStructure As Object            ' One of (structure (99%), located_structure, zone_of_structure, zone_of_building, located_assembly)
End Class

Public Class located_structure
    Public itemNumber As String
    Public itemName As String
    Public itemDescription As String
    Public location As Object
    Public descriptiveStructure As structureDef
    Public parentSite As Object
End Class

Public Class assembly_manufacturing
    Public itemNumber As String
    Public itemName As String
    Public itemDescription As String
    Public lifeCycleStage As String
    Public assemblySequenceNumber As String
    Public complexity As String
    Public surfaceTreatment As String
    Public assemblySequence As String
    Public assemblyUse As String
    Public placeOfAssembly As String    'One of shop_process, site_process, undefined
End Class

Public Class structureDef
    Public itemNumber As String                 ' Should be an integer value
    Public itemName As String
    Public itemDescription As String
End Class

'************************************* TOP LEVEL

''' <summary>
''' The Design_Part class will require datatype classes within it.
''' </summary>
''' <remarks>"parentAssemblies" and "locations" can be a list of 1 to n -- While this is not normal, we need to allow for this, thusly the collecitons</remarks>
Public Class design_part
    Public label As String
    Public partDefinition As Object             ' Can be "Part" or "Part Prismatic Simple" (TriForma)
    Public parentAssembly As Collection         ' Collection of strings -- Names of the parent systems
    Public location As Collection               ' coord_system_cartesian_3d object
End Class

Public Class located_part
    Public item_number As String        ' Should be an integer value
    Public item_name As String
    Public item_description As String
    Public location As Object
    Public partType As part
    Public parent_assembly As located_assembly
End Class


Public Class managed_data_item  ' <---- POS Method
    Public label As String
    Public originatingApplication As String
    Public selectedDataItem As Object
    Public history As Collection
    Public originalData As String
End Class
'************************************* Helpers and Custom Reworks
''' <summary>
''' This is a class to tie the reference standard and source standard together in a collection of associations
''' </summary>
''' <remarks></remarks>
Public Class item_reference_standard
    Public standardOrgName As String
    Public standardName As String
    Public year As String
    Public version As String
    Public sectionName As String
    Public cardinalPoint As String
    Public sectionType As String
End Class

''' <summary>
''' A class to hold the steel members to be constructed, in a generic class.
''' </summary>
''' <remarks></remarks>
Public Class steelObject
    Public label As String
    Public parentSystemLabel As String
    Public structureObject As StructuralSystem
    Public startPoint As Position
    Public endPoint As Position
    Public oStruct As CrossSection
    Public oMaterial As Material
    Public lType As Long
    Public lCard As Long
    Public dRotAngle As Double
    Public bMir As Boolean
    Public lTypeCategory As Long            'Catagory type Beam=1, Column=2, brace=3
    Public createdObject As MemberSystem
End Class

''' <summary>
''' Holds bounding box coordinates -- Intended for use with grid areas (originally)
''' </summary>
''' <remarks></remarks>
Public Class twoDBoundingBox
    Public topRight As New Position
    Public bottomLeft As New Position

    Public Sub New()
        topRight.X = -1000000
        topRight.Y = -1000000
        bottomLeft.X = 1000000
        bottomLeft.Y = 1000000
    End Sub
End Class

Public Class mt
    Public North As Integer = 0
    Public East As Integer = 0
    Public West As Integer = 0
    Public South As Integer = 0

    Public NorthBeam As Integer = 0
    Public EastBeam As Integer = 0
    Public WestBeam As Integer = 0
    Public SouthBeam As Integer = 0
End Class

Public Class cartesianBounds
    Public XMin As Double
    Public XMax As Double
    Public YMin As Double
    Public YMax As Double
    Public ZMin As Double
    Public ZMax As Double

    Public Sub New(ByVal startPoint As Position, ByVal endPoint As Position)
        If startPoint.X > endPoint.X Then
            XMax = startPoint.X
            XMin = endPoint.X
        Else
            XMin = startPoint.X
            XMax = endPoint.X
        End If

        If startPoint.Y > endPoint.Y Then
            YMax = startPoint.Y
            YMin = endPoint.Y
        Else
            YMin = startPoint.Y
            YMax = endPoint.Y
        End If

        If startPoint.Z > endPoint.Z Then
            ZMax = startPoint.Z
            ZMin = endPoint.Z
        Else
            ZMin = startPoint.Z
            ZMax = endPoint.Z
        End If
    End Sub

    Public Function inBounds(ByVal point1 As Position) As Boolean
        Dim result As Boolean = False

        If point1.X <= XMax And point1.X >= XMin And point1.Y <= YMax And point1.Y >= YMin And point1.Z <= ZMax And point1.Z >= ZMin Then
            result = True
        End If

        inBounds = result
    End Function
End Class

Public Class beamPositionConnections
    Public beamInCollection As Integer
    Public where As String
    Public angle As Double
End Class
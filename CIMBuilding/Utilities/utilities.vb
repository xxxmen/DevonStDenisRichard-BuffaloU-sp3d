Imports System.Text.RegularExpressions
Imports Ingr.SP3D.Common.Middle

Module utilities

    ''' <summary>
    ''' Returns the vector angle between two points and normalizes it to 180 degrees (i.e. no negatives) 
    ''' </summary>
    ''' <param name="position1"></param>
    ''' <param name="position2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRoundedNormalizedAngle(ByVal position1 As Position, ByVal position2 As Position) As Integer
        Dim result As Double

        result = Math.Atan2(position1.Y - position2.Y, position2.X - position1.X) * 180 / Math.PI

        If result < 0 Then
            result = result + 180
        End If

        result = Math.Round(result, 0)

        getRoundedNormalizedAngle = result
    End Function

    ''' <summary>
    ''' Checks if a point is on a line.
    ''' </summary>
    ''' <param name="queryPoint">point to check</param>
    ''' <param name="lineEndPoint1">one end of the line</param>
    ''' <param name="lineEndPoint2">other end of the line</param>
    ''' <param name="tolerance">Answer tolerance. Defaulted to 0 (normal case).  If changed, the solution will check the +/- range</param>
    ''' <returns></returns>
    ''' <remarks>The function results in about twice the distance from the line.</remarks>
    Public Function isPointOnALine(ByVal queryPoint As Position, ByVal lineEndPoint1 As Position, ByVal lineEndPoint2 As Position, Optional ByVal tolerance As Double = 0) As Boolean
        Dim result As Double
        Dim answer As Boolean = False

        result = (lineEndPoint2.X - lineEndPoint1.X) * (queryPoint.Y - lineEndPoint1.Y) - (lineEndPoint2.Y - lineEndPoint1.Y) * (queryPoint.X - lineEndPoint1.X)

        If (result = 0 And tolerance = 0) Or (tolerance <> 0 And Math.Abs(result) <= Math.Abs(tolerance)) Then
            answer = True
        End If
        
        isPointOnALine = answer
    End Function

    ''' <summary>
    ''' Checks if the "pointToCheck" is one of the two planes created by the planepoint and its rotational vector.  This is checked about the z vector
    ''' at 90 degrees.  Plane 1 is its normal plane (e.g. perpendicular to the web of a W-section).  Plane 2 is the rotated plane.
    ''' </summary>
    ''' <param name="planePoint"></param>
    ''' <param name="pointToCheck"></param>
    ''' <param name="tolerance"></param>
    ''' <returns>0 = Not on a plane; 1 on one plane; 2 on the other plane</returns>
    ''' <remarks></remarks>
    Public Function isPointOnDualPlane(ByVal planePoint As Position, ByVal pointToCheck As Position, ByVal rotation As Double, Optional ByVal tolerance As Double = 0) As Integer
        Dim pointp As New Position
        Dim result As Integer = 0

        pointp.X = planePoint.X + Math.Cos(rotation)
        pointp.Y = planePoint.Y + Math.Sin(rotation)

        If isPointOnALine(pointToCheck, planePoint, pointp, tolerance) = True Then
            result = 1
        End If

        pointp.X = planePoint.X - Math.Sin(rotation)
        pointp.Y = planePoint.Y + Math.Cos(rotation)

        If isPointOnALine(pointToCheck, planePoint, pointp, tolerance) = True Then
            result = 2
        End If

        isPointOnDualPlane = result
    End Function

    Public Function isPointOnALine3D(ByVal queryPoint As Position, ByVal lineEndPoint1 As Position, ByVal lineEndPoint2 As Position, Optional ByVal tolerance As Double = 0) As Boolean
        Dim answer As Boolean

        Dim qp As New Position
        Dim l1 As New Position
        Dim l2 As New Position

        answer = isPointOnALine(queryPoint, lineEndPoint1, lineEndPoint2, tolerance)

        If answer = True Then
            qp.X = queryPoint.X
            qp.Y = queryPoint.Z

            l1.X = lineEndPoint1.X
            l1.Y = lineEndPoint1.Z

            l2.X = lineEndPoint2.X
            l2.Y = lineEndPoint2.Z

            answer = isPointOnALine(qp, l1, l2, tolerance)

            If answer = True Then
                qp.X = queryPoint.Y
                qp.Y = queryPoint.Z

                l1.X = lineEndPoint1.Y
                l1.Y = lineEndPoint1.Z

                l2.X = lineEndPoint2.Y
                l2.Y = lineEndPoint2.Z

                answer = isPointOnALine(qp, l1, l2, tolerance)
            End If
        End If


        isPointOnALine3D = answer
    End Function

    ''' <summary>
    ''' Checks is a key value exists in a colleciton.  Returns "DNE" if the key is not found.
    ''' </summary>
    ''' <param name="collectionToSearch"></param>
    ''' <param name="keyValueToFind"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function existsInCollection(ByVal collectionToSearch As Collection, ByVal keyValueToFind As String) As String
        'On Error GoTo falseCase
        Dim result As String

        Try
            If collectionToSearch(keyValueToFind) IsNot Nothing Then
                result = collectionToSearch(keyValueToFind)
            End If
        Catch ex As Exception
            result = "DNE"
        End Try

        existsInCollection = result
    End Function

    ''' <summary>
    ''' Attempts to remove an element form a collection.  To be used if you do not know the element exists.
    ''' </summary>
    ''' <param name="collectionObject"></param>
    ''' <param name="keyValue"></param>
    ''' <remarks></remarks>
    Public Sub removeCollectionElement(ByRef collectionObject As Collection, ByVal keyValue As String)
        Try
            collectionObject.Remove(keyValue)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Returns the number of elements in a collection.  Returns 0 if the collection is empty or not initialized.
    ''' </summary>
    ''' <param name="collectionObject"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function collectionCount(ByVal collectionObject As Collection) As Integer
        Dim result As Integer = 0

        On Error Resume Next

        If collectionObject.Count > 0 Then
            result = collectionObject.Count
        End If

        collectionCount = result
    End Function

    ''' <summary>
    ''' Deals with duplicate keys in a collection -- for adding
    ''' </summary>
    ''' <param name="cCollection"></param>
    ''' <param name="data"></param>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Sub addToCollection(ByRef cCollection As Collection, ByVal data As Object, ByVal key As String)
        Try
            cCollection.Add(data, key)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Compares the equality of two numbers within a precision factor (e.g. 0.02 difference)
    ''' </summary>
    ''' <param name="firstNumber"></param>
    ''' <param name="SecondNumber"></param>
    ''' <param name="precision">How far apart the numbers can be</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function compareNumbers(ByVal firstNumber As Double, ByVal SecondNumber As Double, ByVal precision As Double) As Boolean
        Dim result As Boolean = False
        Dim mathBit As Double = Math.Abs(firstNumber - SecondNumber)

        If mathBit <= precision Then
            result = True
        End If

        compareNumbers = result
    End Function

    ''' <summary>
    ''' Checks if the first number is less than the second by a value greater than the 'precision'
    ''' </summary>
    ''' <param name="firstNumber"></param>
    ''' <param name="SecondNumber"></param>
    ''' <param name="precision"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function compareLessThan(ByVal firstNumber As Double, ByVal SecondNumber As Double, ByVal precision As Double) As Boolean
        Dim result As Boolean = False

        If compareNumbers(firstNumber, SecondNumber, precision) = False And firstNumber < SecondNumber Then
            result = True
        End If

        compareLessThan = result
    End Function

    Public Function degToRad(ByVal degreeValue As Double) As Double
        degToRad = degreeValue / 180 * Math.PI
    End Function

    Public Function radToDeg(ByVal radianValue As Double) As Double
        radToDeg = radianValue / 180 * Math.PI
    End Function

    Public Function normalizeTo90(ByVal radianAngle As Double) As Double
        While radianAngle > Math.PI / 2
            radianAngle = radianAngle - Math.PI / 2
        End While

        normalizeTo90 = radianAngle
    End Function

    Public Function getHypotenuseComponent(ByVal mainX As Double, ByVal pointX As Double, ByVal mainY As Double, ByVal pointY As Double, ByVal component As String) As Double
        Dim result As Double

        result = Math.Sqrt((pointX - mainX) ^ 2 + (pointY - mainY) ^ 2)

        If (pointY < mainY And UCase(component) = "Y") Or (pointX < mainX And UCase(component) = "X") Then
            result = -result
        End If

        getHypotenuseComponent = result
    End Function

    Public Function twoDriseOverRunNormalized(ByVal position1 As Position, ByVal position2 As Position) As Double
        Dim result As Double

        result = Math.Abs((position1.Y - position2.Y) / (position1.X - position2.X))

        twoDriseOverRunNormalized = result
    End Function

End Module

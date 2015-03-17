Imports System.Data
Public Class Location
    Public Property lat As Double
    Public Property lon As Double
End Class
Partial Class Default7
    Inherits System.Web.UI.Page

    Const pi = 3.1415926535897931
    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Id As Integer = ListBox1.SelectedIndex + 1
        Dim Id2 As Integer = ListBox2.SelectedIndex + 1
        Dim loc As New Location

        Dim lat1 As Double = GetCity(Id, loc).lat
        Dim lon1 As Double = loc.lon
        Dim lat2 As Double = GetCity(Id2, loc).lat
        Dim lon2 As Double = loc.lon
        Label1.Text = Distance(lat1, lon1, lat2, lon2, "M")
    End Sub

    Function GetCity(Id As Integer, loc As Location) As Location
        Dim Table1 As DataView = CType(AccessDataSource1.Select(DataSourceSelectArguments.Empty), DataView)
        Table1.RowFilter = "ID = '" & Id & "'"
        Dim cRow As DataRowView = Table1(0)
        loc.lat = cRow("cityLatitude")
        loc.lon = cRow("cityLongitude")
        Return loc
    End Function

    Function Distance(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double, ByVal unit As String) As Double
        Dim theta As Double, dist As Double
        theta = lon1 - lon2
        dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta))
        Label1.Text = "dist = " & dist & "<BR>"
        dist = acos(dist)
        dist = rad2deg(dist)
        dist = dist * 60 * 1.1515
        Select Case UCase(unit)
            Case "K"
                dist = dist * 1.609344
            Case "N"
                dist = dist * 0.8684
        End Select
        Return dist
    End Function

    Function acos(ByVal rad As Double) As Double
        If Math.Abs(rad) <> 1 Then
            acos = pi / 2 - Math.Atan(rad / Math.Sqrt(1 - rad * rad))
        ElseIf rad = -1 Then
            acos = pi
        End If
        Return acos
    End Function

    Function deg2rad(ByVal deg As Double) As Double
        deg2rad = CDbl(deg * pi / 180)
        Return deg2rad
    End Function

    Function rad2deg(ByVal Rad As Double) As Double
        rad2deg = CDbl(Rad * 180 / pi)
        Return rad2deg
    End Function

End Class

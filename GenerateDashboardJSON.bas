
Option Compare Database
Option Explicit

' Main function to generate dashboard.json
Public Sub GenerateJSON()

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Dim jsonDict As Object
    Dim headerDict As Object

    Dim statsArray As Collection
    Dim statDict As Object

    Dim miniChartDataArray As Collection

    Dim salesChartData As Object
    Dim donutChartData As Object
    Dim customerChartData As Object

    Dim recentOrdersArray As Collection
    Dim orderDict As Object

    Dim filePath As String
    Dim fileNum As Integer
    Dim jsonOutput As String

    Dim i As Integer

    Set db = CurrentDb

    ' ==================================================
    ' MAIN JSON OBJECT
    ' ==================================================
    Set jsonDict = CreateObject("Scripting.Dictionary")

    ' ==================================================
    ' HEADER
    ' ==================================================
    Set headerDict = CreateObject("Scripting.Dictionary")

    headerDict.Add "title", "Welcome back"
    headerDict.Add "subtitle", "Check your last activity today"

    jsonDict.Add "header", headerDict

    ' ==================================================
    ' STATIC LABELS
    ' ==================================================
    jsonDict.Add "mini_chart_labels", Array("1", "2", "3", "4", "5", "6")

    Dim salesLabels As Variant
    salesLabels = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                         "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    Dim donutLabels As Variant
    donutLabels = Array("Online sales", "Offline sales", "Returns")

    Dim donutColors As Variant
    donutColors = Array("#6366f1", "#fb923c", "#facc15")

    Dim customerLabels As Variant
    customerLabels = Array("Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' ==================================================
    ' STATS
    ' ==================================================
    Set statsArray = New Collection

    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblDashboardStats ORDER BY StatID", _
        dbOpenSnapshot)

    Do While Not rs.EOF

        Set statDict = CreateObject("Scripting.Dictionary")

        statDict.Add "label", Nz(rs!StatName, "")
        statDict.Add "value", Nz(rs!Value, "")
        statDict.Add "trend", Nz(rs!Trend, "")
        statDict.Add "trend_up", Nz(rs!TrendUp, False)

        ' Theme
        Select Case Nz(rs!StatName, "")

            Case "Total orders"
                statDict.Add "color_theme", "dark"
                statDict.Add "mini_chart_color", "#818cf8"

            Case "Total sales"
                statDict.Add "color_theme", "light"
                statDict.Add "mini_chart_color", "#f97316"

            Case "Online sessions"
                statDict.Add "color_theme", "light"
                statDict.Add "mini_chart_color", "#facc15"

            Case "Average order"
                statDict.Add "color_theme", "light"
                statDict.Add "mini_chart_color", "#ef4444"

            Case Else
                statDict.Add "color_theme", "light"
                statDict.Add "mini_chart_color", "#6366f1"

        End Select

        ' ==============================================
        ' MINI CHART DATA
        ' ==============================================
        Set miniChartDataArray = New Collection

        Dim rsMini As DAO.Recordset

        Set rsMini = db.OpenRecordset( _
            "SELECT DataValue " & _
            "FROM tblStatsMiniChartData " & _
            "WHERE StatID = " & rs!StatID & " " & _
            "ORDER BY [DataOrder]", _
            dbOpenSnapshot)

        Do While Not rsMini.EOF

            miniChartDataArray.Add Nz(rsMini!DataValue, 0)

            rsMini.MoveNext
        Loop

        rsMini.Close
        Set rsMini = Nothing

        Dim miniDataArray() As Variant

        If miniChartDataArray.Count > 0 Then

            ReDim miniDataArray(0 To miniChartDataArray.Count - 1)

            For i = 1 To miniChartDataArray.Count
                miniDataArray(i - 1) = miniChartDataArray(i)
            Next i

            statDict.Add "mini_chart_data", miniDataArray

        Else

            statDict.Add "mini_chart_data", Array()

        End If

        statsArray.Add statDict

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    ' Convert collection to array
    Dim statsArrayFinal() As Variant

    If statsArray.Count > 0 Then

        ReDim statsArrayFinal(0 To statsArray.Count - 1)

        For i = 1 To statsArray.Count
            Set statsArrayFinal(i - 1) = statsArray(i)
        Next i

        jsonDict.Add "stats", statsArrayFinal

    Else

        jsonDict.Add "stats", Array()

    End If

    ' ==================================================
    ' SALES CHART
    ' ==================================================
    Set salesChartData = CreateObject("Scripting.Dictionary")

    salesChartData.Add "labels", salesLabels

    Dim onlineValues(0 To 11) As Variant
    Dim offlineValues(0 To 11) As Variant

    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblSalesChartData ORDER BY MonthOrder", _
        dbOpenSnapshot)

    Do While Not rs.EOF

        onlineValues(rs!MonthOrder - 1) = Nz(rs!OnlineValue, 0)
        offlineValues(rs!MonthOrder - 1) = Nz(rs!OfflineValue, 0)

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    salesChartData.Add "online", onlineValues
    salesChartData.Add "offline", offlineValues

    jsonDict.Add "sales_chart", salesChartData

    ' ==================================================
    ' DONUT CHART
    ' ==================================================
    Set donutChartData = CreateObject("Scripting.Dictionary")

    donutChartData.Add "labels", donutLabels
    donutChartData.Add "colors", donutColors

    Dim donutValues(0 To 2) As Variant

    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblDonutChart ORDER BY SegmentID", _
        dbOpenSnapshot)

    Do While Not rs.EOF

        donutValues(rs!SegmentID - 1) = Nz(rs!DataValue, 0)

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    donutChartData.Add "data", donutValues

    jsonDict.Add "donut_chart", donutChartData

    ' ==================================================
    ' CUSTOMER CHART
    ' ==================================================
    Set customerChartData = CreateObject("Scripting.Dictionary")

    customerChartData.Add "labels", customerLabels

    Dim loyalValues(0 To 5) As Variant
    Dim newValues(0 To 5) As Variant

    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblCustomerChartData ORDER BY MonthOrder", _
        dbOpenSnapshot)

    Do While Not rs.EOF

        loyalValues(rs!MonthOrder - 1) = Nz(rs!LoyalValue, 0)
        newValues(rs!MonthOrder - 1) = Nz(rs!NewValue, 0)

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    customerChartData.Add "loyal", loyalValues
    customerChartData.Add "new", newValues

    jsonDict.Add "customer_chart", customerChartData

    ' ==================================================
    ' RECENT ORDERS
    ' ==================================================
    Set recentOrdersArray = New Collection

    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblRecentOrders ORDER BY OrderID", _
        dbOpenSnapshot)

    Do While Not rs.EOF

        Set orderDict = CreateObject("Scripting.Dictionary")

        orderDict.Add "product", Nz(rs!Product, "")
        orderDict.Add "date", Nz(rs!OrderDate, "")
        orderDict.Add "price", Nz(rs!Price, "")
        orderDict.Add "status", Nz(rs!Status, "")

        recentOrdersArray.Add orderDict

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    Dim ordersArrayFinal() As Variant

    If recentOrdersArray.Count > 0 Then

        ReDim ordersArrayFinal(0 To recentOrdersArray.Count - 1)

        For i = 1 To recentOrdersArray.Count
            Set ordersArrayFinal(i - 1) = recentOrdersArray(i)
        Next i

        jsonDict.Add "recent_orders", ordersArrayFinal

    Else

        jsonDict.Add "recent_orders", Array()

    End If

    ' ==================================================
    ' CONVERT TO JSON
    ' ==================================================
    jsonOutput = ConvertToJson(jsonDict)

    ' ==================================================
    ' SAVE FILE
    ' ==================================================
    filePath = CurrentProject.Path & "\dashboard.json"

    fileNum = FreeFile

    Open filePath For Output As #fileNum

    Print #fileNum, jsonOutput

    Close #fileNum

    MsgBox "dashboard.json generated successfully!" & vbCrLf & vbCrLf & _
           filePath, vbInformation

Cleanup:

    On Error Resume Next

    If Not rs Is Nothing Then rs.Close

    Set rs = Nothing
    Set db = Nothing

    Set jsonDict = Nothing
    Set statsArray = Nothing
    Set miniChartDataArray = Nothing
    Set salesChartData = Nothing
    Set donutChartData = Nothing
    Set customerChartData = Nothing
    Set recentOrdersArray = Nothing

    Exit Sub

ErrorHandler:

    MsgBox "Error : " & Err.Number & vbCrLf & Err.Description, vbCritical

    Resume Cleanup

End Sub




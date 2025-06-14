
'=========================
' UFT HTML Iteration Logger
'=========================
Dim g_reportFilePath, g_fileObj, g_testStartTime, g_testName, g_iterationCount

'=========================
' Initialize HTML Report
'=========================
Sub StartTestReport(testName)
    Dim fso, hostName, timestamp

    g_testStartTime = Now
    g_testName = testName
    g_iterationCount = 0

    hostName = Environment("LocalHostName")
    timestamp = Replace(Replace(Replace(Now, "/", "_"), ":", "_"), " ", "_")
    g_reportFilePath = "C:\UFT_TestReports\" & hostName & "_" & testName & "_" & timestamp & ".html"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set g_fileObj = fso.CreateTextFile(g_reportFilePath, True)

    g_fileObj.WriteLine "<html><head><title>UFT Test Report - " & testName & "</title></head><body>"
    g_fileObj.WriteLine "<style>body{font-family:Arial;} table{border-collapse:collapse;} th,td{border:1px solid #ccc;padding:6px;}</style>"
    g_fileObj.WriteLine "<h2>Test Report: " & testName & "</h2>"

    Set fso = Nothing
End Sub

'=========================
' Start a New Iteration Section
'=========================
Sub StartIteration(iterationNumber)
    g_iterationCount = g_iterationCount + 1
    g_fileObj.WriteLine "<h3>Iteration " & iterationNumber & "</h3>"
    g_fileObj.WriteLine "<table>"
    g_fileObj.WriteLine "<tr><th>Timestamp</th><th>Status</th><th>Step Name</th><th>Details</th><th>Screenshot</th></tr>"
End Sub

'=========================
' Log a Test Step in Iteration
'=========================
Sub LogTestStep(statusCode, stepName, takeScreenshot)
    Dim statusDict, colorDict, statusText, statusColor
    Dim screenshotPath, htmlRow, timeStamp, cleanStepName

    Set statusDict = CreateObject("Scripting.Dictionary")
    Set colorDict = CreateObject("Scripting.Dictionary")

    statusDict.Add 1, "PASSED"           : colorDict.Add 1, "green"
    statusDict.Add 2, "FAILED"           : colorDict.Add 2, "red"
    statusDict.Add 3, "WARNING"          : colorDict.Add 3, "orange"
    statusDict.Add 4, "DONE"             : colorDict.Add 4, "blue"
    statusDict.Add 5, "REPLAY"           : colorDict.Add 5, "teal"
    statusDict.Add 6, "SKIPPED"          : colorDict.Add 6, "gray"
    statusDict.Add 7, "NOT COMPLETED"    : colorDict.Add 7, "darkred"

    If statusDict.Exists(CInt(statusCode)) Then
        statusText = statusDict(CInt(statusCode))
        statusColor = colorDict(CInt(statusCode))
    Else
        statusText = "INVALID"
        statusColor = "black"
    End If

    timeStamp = Now
    cleanStepName = Replace(Replace(stepName, " ", "_"), ":", "")

    If takeScreenshot = True Then
        screenshotPath = "C:\UFT_TestReports\" & g_testName & "_" & cleanStepName & "_" & Replace(Replace(Replace(Now, "/", "_"), ":", "_"), " ", "_") & ".png"
        Desktop.CaptureBitmap screenshotPath
    Else
        screenshotPath = ""
    End If

    htmlRow = "<tr><td>" & timeStamp & "</td><td><span style='color:" & statusColor & "'><b>" & statusText & "</b></span></td>"
    htmlRow = htmlRow & "<td>" & stepName & "</td><td>Step status: " & statusText & "</td>"
    If screenshotPath <> "" Then
        htmlRow = htmlRow & "<td><a href='" & screenshotPath & "' target='_blank'>View</a></td>"
    Else
        htmlRow = htmlRow & "<td>–</td>"
    End If
    htmlRow = htmlRow & "</tr>"

    g_fileObj.WriteLine htmlRow

    Set statusDict = Nothing
    Set colorDict = Nothing
End Sub

'=========================
' End Iteration
'=========================
Sub EndIteration()
    g_fileObj.WriteLine "</table><br>"
End Sub

'=========================
' Finalize HTML Report
'=========================
Sub EndTestReport()
    Dim endTime, duration

    endTime = Now
    duration = DateDiff("s", g_testStartTime, endTime)

    g_fileObj.WriteLine "<hr>"
    g_fileObj.WriteLine "<p><b>Total Iterations:</b> " & g_iterationCount & "<br>"
    g_fileObj.WriteLine "<b>Test Started:</b> " & g_testStartTime & "<br>"
    g_fileObj.WriteLine "<b>Test Ended:</b> " & endTime & "<br>"
    g_fileObj.WriteLine "<b>Duration:</b> " & duration & " seconds</p>"
    g_fileObj.WriteLine "</body></html>"
    g_fileObj.Close

    Set g_fileObj = Nothing
End Sub

' UFT Custom HTML Report Generator
Option Explicit

Dim objFSO, objFile, strReportPath, dtmStartTime, dtmEndTime

' Initialize FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Set report path
strReportPath = "C:\Windows\Temp\report_" & Replace(Replace(Now(), ":", "-"), "/", "-") & ".html"

' Create HTML report file
Set objFile = objFSO.CreateTextFile(strReportPath, True)

' Write HTML header
objFile.WriteLine "<!DOCTYPE html>" & vbCrLf & _
"<html lang='en'>" & vbCrLf & _
"<head>" & vbCrLf & _
"    <meta charset='UTF-8'>" & vbCrLf & _
"    <title>UFT Test Execution Report</title>" & vbCrLf & _
"    <style>" & vbCrLf & _
"        body { font-family: Arial, sans-serif; margin: 20px; }" & vbCrLf & _
"        .header { background-color: #4472C4; color: white; padding: 20px; margin-bottom: 20px; }" & vbCrLf & _
"        .summary { margin: 20px 0; }" & vbCrLf & _
"        table { width: 100%; border-collapse: collapse; margin: 20px 0; }" & vbCrLf & _
"        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }" & vbCrLf & _
"        th { background-color: #f5f5f5; }" & vbCrLf & _
"        .pass { background-color: #92d050; }" & vbCrLf & _
"        .fail { background-color: #ffc7ce; }" & vbCrLf & _
"        .warning { background-color: #ffeb9c; }" & vbCrLf & _
"        .details { display: none; }" & vbCrLf & _
"        .expandBtn { cursor: pointer; color: blue; text-decoration: underline; }" & vbCrLf & _
"    </style>" & vbCrLf & _
"    <script>" & vbCrLf & _
"        function toggleDetails(id) {" & vbCrLf & _
"            var details = document.getElementById('details_' + id);" & vbCrLf & _
"            var btn = document.getElementById('btn_' + id);" & vbCrLf & _
"            if (details.style.display === 'none') {" & vbCrLf & _
"                details.style.display = 'table-row';" & vbCrLf & _
"                btn.textContent = 'Hide Details';" & vbCrLf & _
"            } else {" & vbCrLf & _
"                details.style.display = 'none';" & vbCrLf & _
"                btn.textContent = 'Show Details';" & vbCrLf & _
"            }" & vbCrLf & _
"        }" & vbCrLf & _
"    </script>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body>" & vbCrLf & _
"    <div class='header'>" & vbCrLf & _
"        <h1>UFT Test Execution Report</h1>" & vbCrLf & _
"        <p>Generated on: " & Now() & "</p>" & vbCrLf & _
"    </div>"

' Write summary section
objFile.WriteLine "    <div class='summary'>" & vbCrLf & _
"        <h2>Test Summary</h2>" & vbCrLf & _
"        <p>Total Tests: 3</p>" & vbCrLf & _
"        <p>Passed: 1 | Failed: 1 | Warning: 1</p>" & vbCrLf & _
"    </div>"

' Write table header
objFile.WriteLine "    <table>" & vbCrLf & _
"        <tr>" & vbCrLf & _
"            <th>Test Case ID</th>" & vbCrLf & _
"            <th>Test Case Name</th>" & vbCrLf & _
"            <th>Status</th>" & vbCrLf & _
"            <th>Start Time</th>" & vbCrLf & _
"            <th>End Time</th>" & vbCrLf & _
"            <th>Actions</th>" & vbCrLf & _
"        </tr>"

' Sample test results
Sub AddTestResult(testID, testName, status, startTime, endTime, details)
    Dim statusClass
    Select Case UCase(status)
        Case "PASS"
            statusClass = "pass"
        Case "FAIL"
            statusClass = "fail"
        Case "WARNING"
            statusClass = "warning"
    End Select
    
    objFile.WriteLine "        <tr>" & vbCrLf & _
    "            <td>" & testID & "</td>" & vbCrLf & _
    "            <td>" & testName & "</td>" & vbCrLf & _
    "            <td class='" & statusClass & "'>" & status & "</td>" & vbCrLf & _
    "            <td>" & startTime & "</td>" & vbCrLf & _
    "            <td>" & endTime & "</td>" & vbCrLf & _
    "            <td><span class='expandBtn' id='btn_" & testID & "' onclick='toggleDetails(""" & testID & """)'>Show Details</span></td>" & vbCrLf & _
    "        </tr>" & vbCrLf & _
    "        <tr id='details_" & testID & "' class='details'>" & vbCrLf & _
    "            <td colspan='6'>" & details & "</td>" & vbCrLf & _
    "        </tr>"
End Sub

' Add sample test results
AddTestResult "TC001", "Login Validation", "PASS", Now(), Now(), "Test executed successfully. All assertions passed.<br>Steps:<br>1. Enter credentials<br>2. Click login<br>3. Verify dashboard access"
AddTestResult "TC002", "Search Functionality", "FAIL", Now(), Now(), "Element not found: searchButton<br>Error at step 2: Unable to locate element using xpath: //button[@id='search']"
AddTestResult "TC003", "Data Export", "WARNING", Now(), Now(), "Test completed with warnings.<br>Performance degradation noticed.<br>Response time: 5.2s (Expected: <3s)"

' Close HTML tags
objFile.WriteLine "    </table>" & vbCrLf & _
"</body>" & vbCrLf & _
"</html>"

' Close file
objFile.Close

' Clean up objects
Set objFile = Nothing
Set objFSO = Nothing

WScript.Echo "Report generated successfully at: " & strReportPath
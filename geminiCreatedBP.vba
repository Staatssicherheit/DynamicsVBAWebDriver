' VBA code to automate Microsoft Edge, input data into Microsoft Dynamics 365, and write to Excel
' This code requires Microsoft Excel and the Microsoft Edge WebDriver.
' Ensure the Edge WebDriver is downloaded and placed in a suitable location (and the path is updated in the code).
' No external libraries are used.

Option Explicit

' Declare constants for WebDriver and Excel interaction
Const WEBDRIVER_PATH As String = "C:\msedgedriver.exe" ' **CHANGE THIS TO YOUR Edge WebDriver PATH**
Const DYNAMICS_URL As String = "https://your-dynamics-365-url.crm.dynamics.com/"      ' **CHANGE THIS TO YOUR Dynamics 365 URL**
Const INPUT_FIELD_ID As String = "yourInputFieldId" ' **CHANGE THIS TO THE ID OF THE INPUT FIELD**
Const INPUT_DATA As String = "Your Data Here"       ' **CHANGE THIS TO THE DATA YOU WANT TO INPUT**
Const OUTPUT_FIELD_ID As String = "yourOutputFieldId"    ' **CHANGE THIS TO THE ID of the OUTPUT FIELD**
' Add more constants for other elements as needed for your Dynamics 365 interaction

Sub AutomateDynamicsAndExcel()
    ' Declare variables
    Dim objShell As Object, objExec As Object
    Dim objHTTP As Object
    Dim json As Object
    Dim sessionId As String, urlString As String
    Dim elementId As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim result As String
    Dim command As String
    Dim script As String
    Dim title As String
    Dim dynamicsTitle As String

    'On Error GoTo ErrorHandler ' Enable error handling -  Important for Dynamics

    ' 1. Start Edge WebDriver
    Set objShell = CreateObject("WScript.Shell")
    command = Chr(34) & WEBDRIVER_PATH & Chr(34) & " --port=9515" ' Start on port 9515
    Set objExec = objShell.Exec(command)
    ' Wait for WebDriver to start (adjust the wait time if needed)
    Application.Wait Now + TimeValue("00:00:02") ' Wait 2 seconds. Increased for reliability.

    ' 2. Create a new Excel worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "DynamicsData"
    lastRow = 1 ' Initialize lastRow

    ' 3. Create HTTP object for communication with WebDriver
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    ' 4. Create a new session
    objHTTP.Open "POST", "http://localhost:9515/session", False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send "{""capabilities"":{""alwaysMatch"":{""browserName"":""msedge""},""firstMatch"":[{}]}}"
    'Debug.Print objHTTP.responseText ' For debugging
    Set json = ParseJson(objHTTP.responseText) ' Use the JSON parser
    sessionId = json.value.sessionId
    Debug.Print "Session ID: " & sessionId

    If sessionId = "" Then
        MsgBox "Failed to create a WebDriver session. Check WebDriver path and version."
        GoTo Cleanup
    End If

    ' 5. Navigate to the Dynamics 365 URL
    urlString = "http://localhost:9515/session/" & sessionId & "/url"
    objHTTP.Open "POST", urlString, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send "{""url"":""" & DYNAMICS_URL & """}"
    Debug.Print "Navigate Status: " & objHTTP.Status

    Application.Wait Now + TimeValue("00:00:05") ' Wait for Dynamics to load.  Increased significantly.  Dynamics can be slow.

    ' 6. Get the title of the Dynamics 365 page
    urlString = "http://localhost:9515/session/" & sessionId & "/title"
    objHTTP.Open "GET", urlString, False
    objHTTP.send
    dynamicsTitle = ParseJson(objHTTP.responseText).value
    Debug.Print "Dynamics 365 Page Title: " & dynamicsTitle
    ws.Cells(lastRow, 1).Value = "Dynamics Title"
    ws.Cells(lastRow, 2).Value = dynamicsTitle
    lastRow = lastRow + 1

    ' 7.  Find the input element on the Dynamics 365 page
    urlString = "http://localhost:9515/session/" & sessionId & "/element"
    objHTTP.Open "POST", urlString, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send "{""using"":""id"",""value"":""" & INPUT_FIELD_ID & """}" '  Find by ID
    Set json = ParseJson(objHTTP.responseText)
    elementId = json.value.ELEMENT
    Debug.Print "Input Element ID: " & elementId

    If elementId = "" Then
        MsgBox "Input field not found. Check the INPUT_FIELD_ID.  May need to use a different selector (e.g., XPath) for Dynamics."
        GoTo Cleanup
    End If

    ' 8. Enter data into the input field
    urlString = "http://localhost:9515/session/" & sessionId & "/element/" & elementId & "/value"
    objHTTP.Open "POST", urlString, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send "{""text"":""" & INPUT_DATA & """,""value"":[""" & INPUT_DATA & """]}"
    Debug.Print "Input Status: " & objHTTP.Status
    Application.Wait Now + TimeValue("00:00:01") ' Wait

    ' 9. Find the output element.
    urlString = "http://localhost:9515/session/" & sessionId & "/element"
    objHTTP.Open "POST", urlString, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send "{""using"":""id"",""value"":""" & OUTPUT_FIELD_ID & """}"
    Set json = ParseJson(objHTTP.responseText)
    elementId = json.value.ELEMENT
    Debug.Print "Output Element ID: " & elementId
    If elementId = "" Then
        MsgBox "Output field not found.  Check the OUTPUT_FIELD_ID."
        GoTo Cleanup
    End If

    ' 10. Get the text from the output element
    urlString = "http://localhost:9515/session/" & sessionId & "/element/" & elementId & "/text"
    objHTTP.Open "GET", urlString, False
    objHTTP.send
    result = ParseJson(objHTTP.responseText).value
    Debug.Print "Result: " & result

    ' 11. Write the result to the Excel spreadsheet
    ws.Cells(lastRow, 1).Value = "Output Data"
    ws.Cells(lastRow, 2).Value = result
    lastRow = lastRow + 1

    ' 12.  Example of running Javascript (May need adjustments for Dynamics)
    script = "return document.title;"
    urlString = "http://localhost:9515/session/" & sessionId & "/execute/sync"
    objHTTP.Open "POST", urlString, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.send "{""script"":""" & script & """,""args"":[]}"
    Set json = ParseJson(objHTTP.responseText)
    title = json.value
    Debug.Print "Title from JS: " & title

    ws.Cells(lastRow, 1).Value = "Title (JS)"
    ws.Cells(lastRow, 2).Value = title
    lastRow = lastRow + 1

Cleanup:
    ' 13. Close the session
    urlString = "http://localhost:9515/session/" & sessionId
    objHTTP.Open "DELETE", urlString, False
    objHTTP.send
    Debug.Print "Close Session Status: " & objHTTP.Status

    ' 14. Quit WebDriver (attempt to, even with errors)
    On Error Resume Next
    If Not objExec Is Nothing Then
        objExec.Terminate
    End If
    Set objExec = Nothing
    Set objShell = Nothing
    Set objHTTP = Nothing
    Set ws = Nothing
    On Error GoTo 0 ' Restore default error handling

    MsgBox "Automation complete. Data written to sheet 'DynamicsData'."

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description & " (Error Code: " & Err.Number & ")"
        ' Clean up resources in case of error
        On Error Resume Next
        If Not objExec Is Nothing Then
            objExec.Terminate
        End If
        Set objExec = Nothing
        Set objShell = Nothing
        Set objHTTP = Nothing
        Set ws = Nothing
        On Error GoTo 0
    End If
End Sub

' Function to parse JSON (simplified - handles basic JSON structures)
Function ParseJson(jsonString As String) As Object
    Dim scriptEngine As Object
    Set scriptEngine = CreateObject("ScriptControl")
    scriptEngine.Language = "JScript"
    scriptEngine.AddCode "function parse(jsonString) { return JSON.parse(jsonString); }"
    Set ParseJson = scriptEngine.Run("parse", jsonString)
    Set scriptEngine = Nothing
End Function

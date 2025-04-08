Private Sub SlideShowWindows(Index As Integer) Handles Application.SlideShowWindows

    Dim oSlide As Slide
    Dim oShape As Shape

    ' Check if it's the start of a new slide show
    If Index > 0 Then
        Set oSlide = SlideShowWindows(Index).View.Slide
        For Each oShape In oSlide.Shapes
            ' Check if the shape is a media object (video or audio)
            If oShape.Type = msoMedia Then
                ' Store the shape object in a public variable for the event handler
                Set g_oCurrentMedia = oShape
                Exit For ' Assuming only one video per slide for simplicity
            End If
        Next oShape
    End If

End Sub

Private Sub Document_KeyDown(ByVal KeyCode As Long, ByVal Shift As Integer) Handles Application.DocumentBeforeClose, Application.SlideShowBegin, Application.SlideShowEnd, Application.SlideShowNextSlide, Application.SlideShowPreviousSlide, Application.SlideShowGotoSlide, Application.SlideShowOnSlideChange

    Static bIsPlaying As Boolean
    Const NAVIGATE_SECONDS As Single = 5 ' Define the navigation step in seconds

    ' Check if a media object is currently active
    If Not g_oCurrentMedia Is Nothing Then
        ' Check for Spacebar (Pause/Resume)
        If KeyCode = 32 Then ' ASCII code for Spacebar
            If bIsPlaying Then
                g_oCurrentMedia.MediaFormat.Player.Pause
                bIsPlaying = False
            Else
                g_oCurrentMedia.MediaFormat.Player.Play
                bIsPlaying = True
            End If
        ' Check for Left Arrow (Rewind)
        ElseIf KeyCode = 37 Then ' ASCII code for Left Arrow
            If g_oCurrentMedia.MediaFormat.Player.CanSeek Then
                g_oCurrentMedia.MediaFormat.Player.CurrentPosition = g_oCurrentMedia.MediaFormat.Player.CurrentPosition - NAVIGATE_SECONDS
            End If
        ' Check for Right Arrow (Forward)
        ElseIf KeyCode = 39 Then ' ASCII code for Right Arrow
            If g_oCurrentMedia.MediaFormat.Player.CanSeek Then
                g_oCurrentMedia.MediaFormat.Player.CurrentPosition = g_oCurrentMedia.MediaFormat.Player.CurrentPosition + NAVIGATE_SECONDS
            End If
        End If
    End If

End Sub

' Public variable to store the currently active media shape
Public g_oCurrentMedia As Shape

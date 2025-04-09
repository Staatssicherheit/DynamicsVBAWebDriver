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

    ' Check if a media object is currently active
    If Not g_oCurrentMedia Is Nothing Then
        ' Check if the pressed key is the Spacebar
        If KeyCode = 32 Then ' ASCII code for Spacebar
            ' Toggle play/pause
            If bIsPlaying Then
                g_oCurrentMedia.MediaFormat.Player.Pause
                bIsPlaying = False
            Else
                g_oCurrentMedia.MediaFormat.Player.Play
                bIsPlaying = True
            End If
        End If
    End If

End Sub

' Public variable to store the currently active media shape
Public g_oCurrentMedia As Shape

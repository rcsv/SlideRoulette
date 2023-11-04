Option Explicit

' Constants for special slides:
Public Const TitleSlideIndex As Integer = 1
Public Const InstructionSlideIndex As Integer = 2
Public Const FirstSlideIndex As Integer = 3

' Timer constants:
Public Const InitialDelay As Integer = 100
Public Const SlowdownTimeInterval As Integer = 2000
Public Const TimerDelayIncrease As Integer = 100

' Sound file names:
Private Const SoundButtonClick As String = "button-click.wav"
Private Const SoundDrumroll As String = "drumroll.wav"
Private Const SoundFanfare As String = "fanfare.wav"

' TextBox Name
Private Const TextBoxName_StopSlides As String = "StoppedSlideNumbers"

' Window API functions for timer and sound:
Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Integer
Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long

Private TimerID As LongPtr
Private StopTimerID As LongPtr
Private Running As Boolean
Private IncreaseDelay As Boolean
Private Delay As Long

' Toggle switch for enables playing sounds: required preparing three .wav files listed above.
Private Const EnableSound As Boolean = False

Private Const SND_SYNC As Long = &H0
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000

Private SelectedSlides As Collection

Public Sub StartRoulette()
    If Running Then Exit Sub
    Running = True

    If SelectedSlides Is Nothing Then
        Set SelectedSlides = New Collection
    End If

    Dim MaxNumber As Integer
    MaxNumber = ActivePresentation.Slides.Count - FirstSlideIndex + 1

    If SelectedSlides.Count >= MaxNumber Then
        MsgBox "All slides have been displayed." & vbCrLf & _
               "To reset the roulette history, please press the 'Reset' button.", vbOKOnly, "Reset Slide Number"
        Exit Sub
    End If

    If EnableSound Then
        PlaySound ActivePresentation.Path & "\" & SoundButtonClick, 0&, SND_SYNC Or SND_FILENAME
    End If

    Delay = InitialDelay
    TimerID = SetTimer(0&, 0&, Delay, AddressOf TimerProc)

    If EnableSound Then
        PlaySound ActivePresentation.Path & "\" & SoundDrumroll, 0&, SND_ASYNC Or SND_FILENAME
    End If
End Sub

Public Sub StopRoulette()
    If Not Running Then Exit Sub

    If EnableSound Then
        PlaySound ActivePresentation.Path & "\" & SoundButtonClick, 0&, SND_ASYNC Or SND_FILENAME
    End If

    IncreaseDelay = True
    StopTimerID = SetTimer(0&, 0&, SlowdownTimeInterval, AddressOf StopTimerProc)
End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTime As Long)
    With ActivePresentation.SlideShowWindow.View
        Dim totalSlides As Integer
        Dim randomSlideIndex As Integer

        totalSlides = ActivePresentation.Slides.Count
        Do
            randomSlideIndex = Int(Rnd * (totalSlides - FirstSlideIndex + 1) + FirstSlideIndex)
        Loop While randomSlideIndex = InstructionSlideIndex Or IsInSelectedSlides(randomSlideIndex)

        .GotoSlide randomSlideIndex
    End With

    If IncreaseDelay Then
        Delay = Delay + TimerDelayIncrease
        KillTimer 0&, TimerID
        TimerID = SetTimer(0&, 0&, Delay, AddressOf TimerProc)
    End If
End Sub

Private Sub StopTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTime As Long)
    KillTimer 0&, TimerID
    KillTimer 0&, StopTimerID
    Running = False
    IncreaseDelay = False

    If EnableSound Then
        PlaySound ActivePresentation.Path & "\" & SoundFanfare, 0&, SND_ASYNC Or SND_FILENAME
    End If

    Dim CurrentSlideIndex As Integer
    With ActivePresentation.SlideShowWindow.View
        CurrentSlideIndex = .CurrentShowPosition
        .GotoSlide CurrentSlideIndex
    End With

    On Error Resume Next
    SelectedSlides.Add CurrentSlideIndex, CStr(CurrentSlideIndex)
    On Error GoTo 0

    AddSlideNumberToSecondSlide CurrentSlideIndex
End Sub

Public Sub ResetHistory()
    Set SelectedSlides = Nothing
    Dim shp As Shape
    Set shp = GetSlideNumberTextBox()
    shp.TextFrame.TextRange.Text = ""
End Sub

Private Function IsInSelectedSlides(ByVal slideIndex As Integer) As Boolean
    Dim slide As Variant
    For Each slide In SelectedSlides
        If slide = slideIndex Then
            IsInSelectedSlides = True
            Exit Function
        End If
    Next slide
    IsInSelectedSlides = False
End Function

Private Function GetSlideNumberTextBox(Optional ByVal slideNumber As Integer = InstructionSlideIndex) As Shape
    Dim sl As Object
    Dim shp As Shape
    Set sl = ActivePresentation.Slides(slideNumber)

    On Error Resume Next
    Set shp = sl.Shapes(TextBoxName_StopSlides)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = sl.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=100, Top:=100, Width:=400, Height:=300)
        shp.Name = TextBoxName_StopSlides
    End If

    Set GetSlideNumberTextBox = shp
End Function

Public Sub AddSlideNumberToSecondSlide(Optional ByVal slideNumber As Integer = InstructionSlideIndex)
    Dim shp As Shape
    Set shp = GetSlideNumberTextBox()

    If shp.TextFrame.TextRange.Text = "" Then
        shp.TextFrame.TextRange.Text = CStr(slideNumber)
    Else
        shp.TextFrame.TextRange.Text = shp.TextFrame.TextRange.Text & ", " & CStr(slideNumber)
    End If
End Sub


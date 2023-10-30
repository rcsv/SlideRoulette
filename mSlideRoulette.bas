'
' SlideRoulette
' A PowerPoint VBA tool to randomize slide presentations with aroulette-like effect. Spice up your
' presentations by adding an element of surprise!
'
' Rcsvpg
' @since 2023-10-10
'
Option Explicit

' path to sound files
' Note: Embedded audio files within PowerPoint can be challenging to invoke programmatically.
' Therefore, for ease of access and stability, it's assumed that audio files are stored externally.
' Adjust the path below to point to the location of your external sound files.
Public Const soundpath As String = ""

' Slide indices for special purposes:

' TitleSlideIndex: This slide serves as the introduction or cover for the presentation.
' It's not included in the roulette spins to ensure the audience always starts from a consistent point.
Public Const TitleSlideIndex As Integer = 1

' InstructionSlideIndex: A dedicated slide providing instructions or guidelines for the audience.
' As with the title slide, this isn't included in the roulette to avoid disrupting the flow.
Public Const InstructionSlideIndex As Integer = 2

' FirstSlideIndex: Indicates the starting point for the actual content slides which will be part of the roulette.
' Slides before this index are considered as meta-slides and won't be part of the random spins.
Public Const FirstSlideIndex As Integer = 3

' Initial timer delay in milliseconds.
Public Const InitialDelay As Integer = 100

' Slowdown Time interval
Public Const SlowdownTimeInterval As Integer = 2000

' Window API functions for timer and sound.
' These allow us to execute certain actions at regular intervals and to play sounds.
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Integer

Public Declare PtrSafe Function PlaySound Lib "winmm.dll" _
    Alias "PlaySoundA" (ByVal lpszName As String, _
    ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000

' Global variables.
Public TimerID As LongPtr
Public StopTimerID As LongPtr
Public Running As Boolean
Public IncreaseDelay As Boolean
Public Delay As Long

' Store slide indices that were already displayed.
Public SelectedSlides As Collection

'
' ENTRY POINT : 1 - StartRoulette
' 
Public Sub StartRoulette()
  ' Initialize and start the SlideRoulette with Randomizer
  If Running Then Exit Sub
  Running = True
  If SelectedSlides Is Nothing Then
    Set SelectedSlides = New Collection
  End If
  Delay = InitialDelay
  TimerID = SetTimer(0&, 0&, Delay, AddressOf TimerProc)
End Sub

'
' ENTRY POINT : 2 - StopRoulette
'
Public Sub StopRoulette()
  ' Stop the slide randomizer and add the current slide to the selected slide collection

  If Not Running Then Exit Sub
  IncreaseDelay = True
  StopTimerID = SetTimer(0&, 0&, SlowdownTimeInterval, AddressOf StopTimerProc)

  Dim currentSlideIndex As Integer
  currentSlideIndex = ActivePresentation.SlideShowWindow.View.slide.slideIndex
  On Error Resume Next
  SelectedSlides.Add currentSlideIndex, CStr(currentSlideIndex)
  On Error GoTo 0
End Sub

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal nIDEvent As Long, ByVal dwTime As Long)
    ' Function to be called at regular intervals. Determines which slide to show next.

  With ActivePresentation.SlideShowWindow.View
    Dim totalSlides As Integer = ActivePresentation.Slides.Count
    Dim randomSlideIndex As Integer
    Do
      randomSlideIndex = Int(Rnd * (totalSlides - FirstSlideIndex + 1) + FirstSlideIndex)
    Loop While randomSlideIndex = InstructionSlideIndex Or IsInSelectedSlides(randomSlideIndex)
      .GotoSlide randomSlideIndex
  End With

  If IncreaseDelay Then
    Delay = Delay + 100
    KillTimer 0&, TimerID
    TimerID = SetTimer(0&, 0&, Delay, AddressOf TimerProc)
  End If
End Sub

Public Sub StopTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
  ByVal nIDEvent As Long, ByVal dwTime As Long)
  ' Function to stop the timer

  KillTimer 0&, TimerID
  KillTimer 0&, StopTimerID
  Running = False
  IncreaseDelay = False
  With ActivePresentation.SlideShowWindow.View
    .GotoSlide .CurrentShowPosition
  End With
End Sub

'
' ENTRY POINT : 3 - Reset History
'
Public Sub ResetHistory()
  ' Reset the collection of displayed slides.
  Set Selected Slides = Nothing
End Sub

Public Function IsInSelectedSlides(ByVal slideIndex As Integer) As Boolean
  ' Check if a slide index is in the SelectedSlides collection.

  Dim slide As Variant
  For Each slide In SelectedSlides
    If slide = slideIndex Then
      IsInSelectedSlides = True
      Exit Function
    End If
  Next slide
  IsInSelectedSlides = False
End Function

' EOF

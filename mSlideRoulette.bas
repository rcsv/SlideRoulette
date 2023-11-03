'
' SlideRoulette
' A PowerPoint VBA tool to randomize slide presentations with aroulette-like effect. Spice up your
' presentations by adding an element of surprise!
'
' Rcsvpg
' @since 2023-10-10
'
Option Explicit

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

' Duration of timer is slowing down
Public Const SlowdownTimeInterval As Integer = 2000

' Window API functions for timer and sound
' These allow us to execute certain actions at regular intervals and to play sounds.

' SetTimer
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

' KillTimer
Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Integer

' PlaySound
Public Declare PtrSafe Function PlaySound Lib "winmm.dll" _
    Alias "PlaySoundA" (ByVal lpszName As String, _
    ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long

' Sleep
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000

' Global variables.
Private TimerID As LongPtr
Private StopTimerID As LongPtr
Private Running As Boolean
Private IncreaseDelay As Boolean
Private Delay As Long

' Store slide indices that were already displayed.
Private SelectedSlides As Collection

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

  ' 回転するスライドの枚数と SelectedSlides Collection の数が同じ場合、無限ループに陥る
  ' エラーを出して終了する
  Dim MaxNumber As Integer
  MaxNumber = ActivePresentation.Slides.Count - FirstSlideIndex + 1

  If SelectedSlides.Count >= MaxNumber Then
    ' メッセージボックスを使って停止する理由を丁寧に伝える
    MsgBox "全てのスライドを表示しました。" & vbCrLf & _
           "ルーレット履歴をリセットするには「Reset」ボタンを押してください。", vbOKOnly, "スライド番号リセット"
    Exit Sub
  End If

  ' PlaySound Timing 1: Click Start (sync)
  ' PlaySound ActivePresentation.Path & "\" & "button-click.wav", 0&, SND_SYNC Or SND_FILENAME

  Delay = InitialDelay
  TimerID = SetTimer(0&, 0&, Delay, AddressOf TimerProc)

  ' PlaySound Timing 2: Drumrolling (async)
  ' PlaySound ActivePresentation.Path & "\" & "drumroll.wav", 0&, SND_ASYNC Or SND_FILENAME
End Sub

'
' ENTRY POINT : 2 - StopRoulette
'
Public Sub StopRoulette()
  ' Stop the slide randomizer and add the current slide to the selected slide collection

  If Not Running Then Exit Sub

  ' PlaySound Timing 3: Stop Roulette Button
  ' PlaySound ActivePresentation.Path & "\" & "button-click.wav", 0&, SND_ASYNC Or SND_FILENAME

  IncreaseDelay = True
  StopTimerID = SetTimer(0&, 0&, SlowdownTimeInterval, AddressOf StopTimerProc)

End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal nIDEvent As Long, ByVal dwTime As Long)
    ' Function to be called at regular intervals. Determines which slide to show next.

  With ActivePresentation.SlideShowWindow.View

    Dim totalSlides As Integer
    Dim randomSlideIndex As Integer

    totalSlides = ActivePresentation.Slides.Count

    ' loop for avoiding slide specified by InstructionSlideIndex and SelectedSlides collection
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


Private Sub StopTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
  ByVal nIDEvent As Long, ByVal dwTime As Long)
  ' Function to stop the timer

  ' stop timers
  KillTimer 0&, TimerID
  KillTimer 0&, StopTimerID
  Running = False
  IncreaseDelay = False ' Reset increase delay flag

  ' PlaySound Timing 4: Fanfare
  ' PlaySound ActivePresentation.Path & "\" & "fanfare.wav", 0&, SND_ASYNC Or SND_FILENAME

  Dim CurrentSlideIndex As Integer

  With ActivePresentation.SlideShowWindow.View

    CurrentSlideIndex = .CurrentShowPosition ' Redundant...?
    .GotoSlide CurrentSlideIndex

  End With

  On Error Resume Next
  SelectedSlides.Add CurrentSlideIndex, CStr(CurrentSlideIndex)
  On Error GoTo 0

  ' Append a FinalSlideNumber into Second Slide TextBox
  AddSlideNumberToSecondSlide CurrentSlideIndex

End Sub

'
' ENTRY POINT : 3 - Reset History
'
Public Sub ResetHistory()
  ' Reset the collection of displayed slides.

  Set Selected Slides = Nothing
  Debug.Print Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & " Flush SelectedSlides"

  Dim shp As Shape
  Set shp = GetSlideNumberTextBox()
  shp.TextFrame.TextRange.Text = ""

End Sub

'
' SelectedSlides Collectionの中身を一つずつ確認して、
' slideIndex と合致するものがあれば、True を戻す
'
Private Function IsInSelectedSlides(ByVal slideIndex As Integer) As Boolean
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

'
' For Testing
'
Public Sub SlideRoutelle_Test()

  ' Initialize
  Call ClearCollection

  ' Test 1: Create a New Collection
  Set SelectedSlides = New Collection

  ' Test 2: Add two indices
  SelectedSlides.Add 5, CStr(5)
  SelectedSlides.Add 8, CStr(8)

  ' Test 3: IsSelectedSlides
  Debug.Print ("Expect True : " & IsInSelectedSlides(5))
  Debug.Print ("Expect False: " & IsInSelectedSlides(8))

End Sub

'
' Helper function: GetSlideNumberTextBox
'
Private Function GetSlideNumberTextBox(Optional ByVal slideNumber As Integer = InstructionSlideIndex) As Shape

  Dim sl As Object
  Dim shp As Shape

  Set sl = ActivePresentation.Slides(slideNumber)

  ' Get TextBox object instance
  On Error Resume Next
  Set shp = sl.Shapes("StoppedSlideNumbers")
  On Error GoTo 0

  ' Create a new TextBox object instance "StoppedSlideNumbers" when it is missing
  If shp Is Nothing Then
    Set shp = sl.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=100, Top:=100, Width:=400, Height:=300) ' 適当
    shp.Name = "StoppedSlideNumbers"
  End If

  Set GetSlideNumberTextBox = shp

End Function

'
' Helper function: AddNumber
'
Public Sub AddSlideNumberToSecondSlide(Optional ByVal slideNumber As Integer = InstructionSlideIndex)

  Dim shp As Shape
  Set shp = GetSlideNumberTextBox()

  ' append previous string and new number
  Dim newText As String
  newText = shp.TextFrame.TextRange.Text & " " & (slideNumber -1)
  shp.TextFrame.TextRange.Text = newText

End Sub

' EOF

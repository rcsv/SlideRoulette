'
' SlideRoulette
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

' Window API functions for timer and sound.
' These allow us to execute certain actions at regular intervals and to play sounds.

---
title: "Slake - Creating snake in Excel"
date: 2019-12-27T18:35:26Z
draft: false
toc: false
images:
tags: 
  - vba
  - games
  - excel
---

It's that time of the year again and I'm starting to form what may become a personal tradition - creating a small game at the end of the year. Almost everyone has played snake in one form or another and it has a simple ruleset, making it a quick and fun project to put together.

![Slake](/gifs/Slake.gif#center)

I pay homage to the [great] game inside a piece of 90's music software called [FastTracker II](https://en.wikipedia.org/wiki/FastTracker_2). I spent many hours listening to 90's trance and drum and bass while completing the inbuilt version of snake, called Nibbles.

You can view the source and [download the game from Gitlab here](https://gitlab.com/dieter.g/slake_excel).

The most challenging bit of code was the responsiveness of keystrokes, directing the snake to change direction at a moments notice, and in many cases, performing two or more actions in quick succession.

By inserting the change of direction (executed by the worksheet change event) within the master loop, the change of direction is picked up as fast as a loop can run on the machine. When tied together with a string buffer containing keystrokes, the snake can be directed to pick up and make every direction change requested when the update timer triggers, ensuring that nothing is missed.

The primary loop:
{{< highlight vb.net >}}
  While Not GameOver
    If moveCompleted Then
      ChangeDirection '<--- Change direction is called at every possible interval.
    End If
    moveCompleted = False
    If Timer >= frameUpdate Then
      frameUpdate = Timer + GameSpeed
      MoveSlake '<--- The snake is only updated once every frame update though.
      moveCompleted = True
      PlaceFood
    End If
    DoEvents
  Wend
{{< /highlight >}}

The direction change routine's only purpose is to locate the range (cell) that the snake's head should move to:
{{< highlight vb.net >}}
Private Sub ChangeDirection()
  Dim oldLocation As Range, previousDirection As Long

  previousDirection = CurrentDirection
  If Len(MoveBuffer) > 0 Then
    CurrentDirection = Left(MoveBuffer, 1)
    MoveBuffer = Mid(MoveBuffer, 2)
  End If
  
  'Safeguard against traveling in the opposite direction
  If previousDirection = D_UP And CurrentDirection = D_DOWN Then
    CurrentDirection = D_UP
  ElseIf previousDirection = D_DOWN And CurrentDirection = D_UP Then
    CurrentDirection = D_DOWN
  ElseIf previousDirection = D_LEFT And CurrentDirection = D_RIGHT Then
    CurrentDirection = D_LEFT
  ElseIf previousDirection = D_RIGHT And CurrentDirection = D_LEFT Then
    CurrentDirection = D_RIGHT
  End If

  Set oldLocation = slakeSheet.Range(Slake(0))
  Select Case CurrentDirection
    Case D_UP
      Set NewLocation = oldLocation.Offset(-1)
    Case D_DOWN
      Set NewLocation = oldLocation.Offset(1)
    Case D_LEFT
      Set NewLocation = oldLocation.Offset(0, -1)
    Case D_RIGHT
      Set NewLocation = oldLocation.Offset(0, 1)
  End Select
End Sub
{{< /highlight >}}

And finally, once the snake is moved, only the new location retrieved from the ChangeDirection routine is added to the snake, and the tail bit removed. While the entirety of the snake is kept in an array, movement is only created by adding on a new cell in the front, and removing the last cell at the end, shifting the array along with each piece of food munched.

Happy holidays!
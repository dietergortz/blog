---
title: "Celltris - Tetris in Excel from Scratch"
date: 2018-12-31T16:51:57Z
draft: false
toc: false
images: 
tags: 
  - excel 
  - tetris 
  - vba
---

I wanted to create a Tetris-like game in Excel without referencing code on the internet. Thus what follows is a brief look at the process I followed to create Celltris (*cough* I know).

Total coding time: Roughly 3-4 days over the Christmas period.

![Celltris](/gifs/celltris.gif#center)
If you would like to take a look at the code (and play the game!), grab it from [the Gitlab page here](https://gitlab.com/dieter.g/celltris).

Disclosure: I ended up having to look up two things:  
  - The rules for Tetris since I couldn't remember the column and row counts.  
  - The vba timer function. I started off attempting to use sleep(), bad idea.  

I used two 2d arrays, the first for the game board's data, the second for the current falling shape. The latter having 2 rows more than the first, thus a 22x10 array and a 20x10. The two additional rows in the shape array allowed spawning of the shapes above the game board. Since the shapes drop one row at a time, these two hidden rows provided a smooth single row drop into the game board.

At the start of the game a timer variable (type single) stores the VBA timer. A simple while-game-is-not-over loop runs, creates a shape, then a secondary loop runs which checks the elapsed time (game speed and level). If the correct amount of time has elapsed, the shape is dropped one row, and if not possible to drop, the secondary loop exits and a new shape is created.

All the shapes are represented as numbers in the arrays, the square or O is 1, the Z is 2, S is 3 and so on. I used Excel's conditional formatting for each shape's colouring.

![ShapeArrays](/images/celltris2.png#center)

Shapes are hardcoded when spawning in and rotated. Rotation checks the shape's current position by verifying blank co-ordinates in the shape's region, then by recreating the shape based on the result. This could probably be done a lot smarter, but since I did not want to reference any code, this was the best I could come up with. Rotation is fairly rudimentary and does not follow the rotation rules for Tetris, which I may revisit and correct one day.

Overall, I was pretty pleased with the result and threw in a nod to Acid Tetris with the smiley face changing as the game is played. All shapes/images have been created in Excel and are simple excel shapes/flowchart items.

Last updated 23/04/2019 to fix a typo.
# Excel-VBA-Projects
Excel games, procedure, user define function, and other interesting stuff built on excel file with VBA


## Games
- Sliding Puzzle

  This game originally from puzzle toys. Simply rearrange puzzle to 1 - 8 sequence after being shambles. 
<p align="center">
  <img src="demo/Sliding Puzzle demo.gif" height="300"><br/>
  <i>Sliding Puzzle</i>
</p>

- Tetris

  This game popular in Nintendo era. Make full row to gain points.
<p align="center">
  <img src="demo/Tetris demo.gif" height="300"><br/>
  <i>Tetris</i>
</p>

- Snake

  This game popular in Nokia mobile phone as built-in game. Just make your snake keep eating apple's and your snake will grow longer.
  But your snake will die when it crash into it's own body. 
<p align="center">
  <img src="demo/Snake demo.gif" height="300"><br/>
  <i>Snake</i>
</p>

- Maze

  You lost on a maze (or a labyrinth, i dont know the difference), and you must go to the other side to escape from this maze.
  But you only bring a torch which not too strong to light all your path. There are 3 level stage.
  This game use kruskal's algorithm to generate a maze. 
<p align="center">
  <img src="demo/Maze demo.gif" height="300"><br/>
  <i>Maze</i>
</p>

## Procedure
- Color Diff
  Fill different color based on specific column

- Go To Advance
  Go To by Value, by Format (Bold / Italic / Underline / Specific Cell Color / Specific Text Color),
  by Calculation (Max, Min, Mode, Above AVG, Below AVG, Equal, Greater, Lower than specific value)

- Sort Sheets
  Sort All sheets A-Z

- Misc
  Custom shortcut keyboard for task: zoom in, zoom out, full screen, normal screen, and rename sheet


## UDF (User Defined Function)
- COLORCODE(cell)
  Get colorcode from a cell

- CONTAIN(within_text, find_text, start_num)
  Check within_text contain string of find_text

- DISTINCTCOUNT(range)
  Count unique row, equal with data count after remove duplicate

- EXTRACTNUMBER(text)
  Get number only from a text

- EXTRACTTEXT(text)
  Get alphabet only from a text

- IFZERO(value, value_if_zero)
  Equal with IF(value=0, value_if_zero, value) in IFERROR style

- ISINSET(value, range_data)
  Check if value is an element of range_data set
  
- MAXIF(max_range, range_criteria, criteria)
  Find max value specified by a given condition or criteria

- MAXIFS(max_range, range_criteria, criteria)
  Find max value specified by multiple given condition or criteria

- MINIF(min_range, range_criteria, criteria)
  Find min value specified by a given condition or criteria

- MINIFS(min_range, range_criteria, criteria)
  Find min value specified by multiple given condition or criteria

- REMOVESPACE(text)
  Remove all space from text

- SCRAMBLE(text)
  Scramble a text

- SELECTONE(range_data)
  Randomly select one element from range_data

- TERBILANG(number, suffix)
  Pronounce number with suffix, use for receipt in indonesian language


## Interesting Stuff
- Hypocloid graph
- 3DRotate graph
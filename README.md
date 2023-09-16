# VBA-Challenge

# External Help Credit


Within a 'huddle' with Lindsey Jessurun (mentioned in the README of my first module assesment too), I got help with what command will run my script across all worksheets in the excel workbook.

Had help from Google and ChatGPT searches in finding each worksheets respective last row (with input), the code provided was:
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
This led me to research what the Rows.Count and End(xlUp) commands were as I did not recall them from classes.

Had help from ChatGPT in revising my last For Loop in my ChallengeCode subroutine- I had initially used 3 separate for loops for each of the coniditonals for Greatest % increase, Greatest % decrease, and Greatest total volume but it revised it to 1 For Loop for all of them.

In my subroutine to clear the results and reformat the column widths, I had help from Google to learn the
  ws.Cells.Columns("___").ColumnWidth
command.

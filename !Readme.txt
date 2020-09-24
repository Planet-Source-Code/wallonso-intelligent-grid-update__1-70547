intelligent Datagrid (UCDataShow) :

Have you ever needed to display unsorted data sorted and grouped in a grid,
or tried to show them in a matrix 
- You didn't want to use the big MSFlex/Flexgrid/FlexH control.
- You get your data in an unarranged way and need to arange it grouped by row and column
- your data often changes row and column values ...

So this is for You.


It will display unsorted/unarraged data in rows and columns where they belong to.
If data belongs to the Same Row and Column it will be added in this "matrix"

as an extra option : it creates an outputstring as HTML.

I added some examples/testcases :
calendar, timeplan, scheduler and how to simply access information from database (uses ADO)
other scenarios are possible : 
ranking, valuebars/gantt-diagram (not really but possible to create), reports


At the end : everything where a 2 dimensional matrix is needed to display indexed data



Changes :
----------
- Added Scrollbars 
- changed PaintGrid for Speed and readybility
- moved generating of HMTL to separate function
- HTML color can be independent from colors in control
- some html-options added 
- added direct database access






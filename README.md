# gSheet-Capacity-Planner
a gscript for creating capacity planning / project tracking gSheet

Simply make a new sheet, click tools->script editor.

paste this script in there, and save, 

From the menu choose *Run -> run function -> onOpen*

You'll have to grant permissions at this step.

go back to the sheet you should see
* Project Overview
* People Overview

They are of course empty.

## add a person
click 'capacity functions - > add person' from the top menu
give them a name

## Assign a person work
In the correct week, add a project name for the person in column B, underneat their 'base' assignments (repeating work which must always happen)

Now assign them a % under the correct week, in the row for that project.

You've assigned them!

## Update Overviews
under capacity planning, click 'update projects'

## Adding a project
go to the project overview sheet, you should see your new project name with a red background, this means the project sheet doesnt exist. Copy the name of the project from that cell and select
Capacity Functions -> add project
paste the name (these must match exactly)
The project sheet will now be created.

## set a projects scope
on the newly created project sheet, enter a number of days/story points effort in cell B2

Go back and look at your project overview, it should now be populated with the work allotments

Look at your person overview, you should see their allotments.

If they go over 80% the cell will turn red to show they are over-assigned

## Track a projects progress
at the end of the week, or the start of the next one, review the points closed off in the past week, and enter them in the 'actual' row of the project sheet.

The plan/actual plan will adjust, and colours will be added to indicate drift (NOT ADDED YET)

If you find the scope has increased, you can enter that in the 'Added Points' row, for the appropriate week.

## Add people to a project
If you need to add people to a project to bring its timeline in range, review the 'person overview' sheet to see who has capacity, click their name to go their sheet.

Add a row for the project in question, and their % assignment in the approriate week.

Now go back to the project sheet, and click
Capacity Functions -> update people on this project

This will update the people for the currently active project sheet.



# Calculate dive time

Calculate dive time based on dive start and end times, using a few rules.
This script imports an excel file with start and end times by date, and transect name, 
and can accomodate more than one dive per day, multiple divers, as well as multiple days.

To calculate dive time:

1. Set your **R** working directory to the folder that contains the script "CalcDiveTime.R".
2. Source the script `source( file="CalcDiveTime.R" )`.
3. A pop-up window will open; select the Excel file with dive start and end times.

Note that the Excel file with dive start and end times must have a specific format.
The first column should be the `Date` in the format "August 14, 2015".
The second column should be the `Transect` name or number.
Then there should be two columns for each diver containing
the start and end times of each dive formatted as `15:12`.
The two columns for each diver should be the diver name followed by `.Start` and `.End`
(e.g., `Matt.Start` and `Matt.End`).
See the example dive times in "Example.xlsx."
This data must be in the first worksheet in the excel file.
If there are multiple excel sheets with dive times (e.g., two different surveys), put them in separate folders.
**ProTip**: Make a copy of the example Excel file, and enter your dive times there.

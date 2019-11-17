Small app that creates frequency tables and pie charts(can be changed through settings)
from raw data in excel, using exceljs (https://github.com/exceljs/exceljs)
and xlsx-chart(https://github.com/objectum/xlsx-chart).

The script assumes that the raw data is fomatted in columns and the first one is the title e.g

height weight <br>
180 90 <br>
190 50 <br>
120 50 <br>

It should be noted that this exact format is used by Google Forms.
you can change the filename in the first line of the script. The new charts will be
created in the same folder as the script, one in each file, with the filename being
the title.

Hi, @all !

If you would like to contribute, please, contact me by email: admin <doggy> dragokas.com
and we can talk together about your skills and possible ways of improvement of the project, including all points that need updating.

Also, you can fork the project, make modifications and contribute directly.

--------------------------------------------
Some points about structure of the project
--------------------------------------------

After downloading the project, you have to open it with bat-file "_0_Open Project Elevated  - !!! - .cmd"

The entry point is a form: frmEULA.frm

After initialization, form frmMain.frm has been started.

Functions call stack while system scanning is looking like this:

"Do a system scan and save log file" button -> cmdN00bLog_Click -> cmdScan_Click -> SaveReport -> StartScan -> CreateLogFile (process list)

'StartScan' contains the list of all sections to scan, like CheckO1Item(), e.t.c ...

The results of scanning are beeing saved in SCAN_RESULT structure as well as fix directives.

Each 'Check' procedure has appropriate 'Fix', like CheckO1Item() <-> FixO1Item().

All other modules and forms are self-explained according to its names.

Best wishes,
Alex.
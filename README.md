# Cal Log.xlsm


This is an Excel document I made to keep track of calibrated devices at work. 

The macro executes as soon as the file is opened.

When used with Window’s “Task Scheduler” you can have it execute automatically whenever you want. 

Open Excel then open the Cal Log in Excel while holding “Shift” at the same time to stop the Marco from executing.

Once the spreadsheet is opened the macro looks at every calibration due date then compares it to the date on the computer 
and determines what is the urgency is; This Month, Next Month, Two Months, Three Months, Calibration Good, Out Of Service,
Should Be Out Of Service, Reference Only, and In Service. 

The macro then saves a copy of itself calling it CalReport.xlsm without the macro so you can open the copy without executing the macro.

Then it sends an email with the CalReport and the email documenting all of the items that need to be addressed.

Then saves changes then closes Excel.

*
Auto_Open() is the macro.

Cal Log.xlsm is the Excel spreadsheet with Auto_Open() embedded in it.

*

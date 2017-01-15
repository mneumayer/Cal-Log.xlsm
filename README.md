# Cal Log.xlsm

* This is an Excel document I made to keep track of calibrated instruments.

## How to run
* Download the file and open Cal log.xlsm.
* The macro executes as soon as the file is opened. 
* When used with Window’s “Task Scheduler” you can have it execute whenever you want. 

## How to open an excel document t without the macro going off.
* Open Excel then open your Excel documnet with the Cal Log macro install while holding “Shift” at the same time to stop the Marco from executing.

## How does it work?
1. Once the spreadsheet is opened the macro looks at every calibration due date then compares it to the date on the computer 
and determines what is the urgency is; This Month, Next Month, Two Months, Three Months, Calibration Good, Out Of Service,
Should Be Out Of Service, Reference Only, and In Service.

2. The macro then saves a copy of itself calling it CalReport.xlsm without the macro so you can open the copy without executing the macro.

3. Then it sends an email with the CalReport and the email documenting all of the items that need to be addressed.

4. Then saves changes then closes Excel.

# How to install in my own Excel document?
* Auto_Open() is the macro.
* Cal Log.xlsm is the Excel spreadsheet with Auto_Open() embedded in it.
* Copy and paste the code into your own macro with the name Auto_Open()(Auto_Open() Tell excel the macro start when the excel sheet is opened).









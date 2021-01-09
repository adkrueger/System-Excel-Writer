# SYSTEM EXCEL WRITER
Author: Aaron Krueger
Publication Date: 8/19/2019
Update Date: 1/8/2021


This application was made by Aaron Krueger for use at Boston Children's Hospital to aid in the research and compilation
of data obtained by many aquatic systems. It aims to take data from one spreadsheet and set of sheets, then compiles
this data in a clean and concise fashion onto another spreadsheet.


&nbsp; &nbsp; 

# **USAGE**

**_Note_**: skip steps 1-3 if you've already downloaded python, pip, and the openpyxl library. step 4 can be skipped if there 
    are no updates to the program

&nbsp; 

**1.** First step is to make sure you've downloaded Python 3.x. There are many versions of Python 3 to download, but this
program was written in Python 3.7.3. To download the most recent version of Python, go to the following link:

        https://www.python.org/downloads/

After this, run the exe file (found where you downloaded it) and the wizard will walk you through installation.

&nbsp; 

**2.** We are going to need to download pip, which is used to install libraries (allows different functionality for the
program). To do this, open Terminal (Mac/Linux) or PowerShell (Windows), and type:

    curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py

And hit enter. Once this is done, type:

    python get-pip.py

And hit enter again.

&nbsp; 

**3.** Now, using Terminal/PowerShell, type:

    pip3 install openpyxl

And hit enter. If you get an error saying something like 'pip3: command not found' then trying using pip instead of
pip3.

&nbsp; 

**4.** Next, download writer.py by clicking on the "Clone or Download" button on the "Code" page of the GitHub repository
(this is found at github.com/adkrueger/System-Excel-Writer). Now, click "Download ZIP". You may need to extract the
contents of this folder after done downloading (which can be done by either going into the folder and selecting the
"Extract" option in the menu, or by right clicking and selecting the "Extract All" option).

&nbsp; 

**5.** There are multiple ways to run the script, but we will use the easier of the two first:
Open the 'writer.py' file by double-clicking on it. This contains all code, and should open up in a new window. Now click the 'Run' option in
the menu bar at the top of the screen (if this doesn't appear, make sure you're on the correct window). In the dropdown
menu that appears, click 'Run Module'.

If the above instructions didn't work, follow the OS-specific instructions below:

&nbsp; 

**WINDOWS**:
Using Command Prompt (type "cmd" in the search bar and click on Command Prompt), type in the following:
python "path where writer.py is stored"

i.e.

    python C:\Users\user\Downloads\System-Excel-Writer-master\writer.py

alternatively, use the "cd" command to navigate to the same directory as writer.py (https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/cd), and type:

    python ./writer.py

&nbsp; 

**MAC OS**:
Open Terminal (i.e. Command+Spacebar, then type "Terminal" and click on the program) and type something like the following:
python3 "path where writer.py is stored"

i.e.

    python3 C:/Users/user/Downloads/System-Excel-Writer-master/writer.py

*or*

    python3 ./writer.py                 <-- note that this only works if your terminal is in the same directory as writer.py

This will run the program, where a window will pop up with instructions on how to run the program.

&nbsp; 

**LINUX**:
Use terminal to navigate to the folder where writer.py is stored.
Type "chmod u+x writer.py" to allow the program to execute.
Type "./writer.py" and the program will run


6. Follow the instructions given in the program to compile the spreadsheet.

&nbsp; 

&nbsp; 



# NOTES:
The initial spreadsheet should follow a certain "expected format"
    Column A should have dates in "MM/DD/YYYY" or similar format
    Column B should have individual days for each date (i.e. Monday, Tuesday, etc.)
    Columns with actual data need to have the data type listed above the numbers (i.e. pH or Alkalinity)
    See Troubleshooting concern #2 for information on system/sheet titling

If a word is written instead of actual data or the cell is left blank, the program will ignore it when calculating means.
    However, the program can't discern whether or not a number is too high or too low, so make sure numbers are input
    properly and without error.

Data can be edited freely after successful compilation.

The program *will* overwrite any excel files in the current directory with the name "writebook"
    This includes past years' compilations!
    To disregard this, simply move the other "writebook" to another directory.

The program will not save properly if "writebook" is already open on the computer.

&nbsp; 

&nbsp; 

# TROUBLESHOOTING:
I clicked the "X" button and the program stopped, what should I do?
    Run the program again and make sure you only use the "Exit" button given at the end of the program.
    (Exiting the program preemptively won't cause any harm, you'll just need to start all over)

The system names I am looking to compile won't appear
    For system names to appear in the second step, you must have them labeled with numbers at the end, starting at 1
    and increasing by one for each subsequent system
    GOOD: ABC SYS1, ABC SYS2, ABC SYS3
          ABC SYS 1, ABC SYS 2, ABC SYS 3 (spaces are recognized!)
    BAD: System A, System B, System C
         System 1, System 12, System 142 (this will appear properly, but sheet may not compile properly)
         SystemQ1, System Q2, SystemQ3 (note the difference in spacing - will produce two different possible names)
    Overall: Make sure systems are ordered 1, 2, 3, etc. otherwise the program may have issue recognizing sheets

I don't want the compiled file to be named "writesheet.xlsx"
    You can rename and edit the file as wanted after it's done compiling - it is named "writesheet" for the sake of it
    being unique enough to notice the file name.

I got a mean that is considerably higher/lower than it should be in the final product.
    Make sure the original sheet has no issues in data entered.
    For example, the program is unable to recognize if you meant 6.13 or 613.
    To fix this, change the erroneous number on the original sheet and run the program again.

I want the titles on the graphs to be different.
    You can edit the data as much as needed after compilation, and this includes the graph names.

An error occurs when using the most up to date year
    Check to make sure that there is date padding below the last date of the current year, i.e.
            12/31/2019
            1/1/2020
    Otherwise, the program tries to perform an operation on an empty cell and throws an error

I can't use the arrows on the year selection to go below a certain year
    Try typing in the date, then hit 'Select' and 'Compile.'

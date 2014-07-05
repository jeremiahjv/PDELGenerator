PDELGenerator
=============
NON-PROPRIETARY CODE
VBA project to convert Excel data into a database-readable text file


###clsDataFile.cls
Class takes two inputs, the required test-type (named Partition), and unit number. The output is a sheet object from which data could be extracted. The class has full error checking for each input and each step of the process, along with logging.

###clsLogger.cls
Class contains methods for logging to user-visible console and/or a logfile. The class is called from within another class which is performing the actual work on the desired datafile. The clsLogger methods are easily called and take any string as an input.

###frmMultiDirFile.frm
Windows form code (actual form GUI not included) which presents the user multiple options when multiple valid results are found for searched directories, files, or worksheets. The user is permitted to select only one to continue processing.

###Tools.bas
Module containing useful tools for exporting all code components (classes, methods, forms) for the purpose of merging similar files. Also contains shortcut tools.

###Testing.bas
Module containing unit tests for clsDataFile to verify inputs, outputs, and intermediate steps.

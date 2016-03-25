# keysurvey-parser

##1-Month export
1. Delete rows 1-5
2. Clear text from new row 1
3. Paste headers from corresponding file in **header\_files** folder to new row 1
4. Save file to **survey\_files** folder
5. Open **.env** file and set **FILE\_TO\_PARSE** equal to the name of the file (**including the .xls extension**) and set the **FIELDS\_TO\_USE\_FOR\_CSV** equal to which survey month it is (1, 3, 6, 12, or 18)
6. npm start to run
7. New CSV will appear in **parsed\_files** folder upon completion

'*************************
' Purpose:
' This function is a VBscript CSV data parser (no external controls). I used to use
' the SPLIT function, but I found it problematic, depending on HOW the CSV file
' was written.
'
' An ordinary CSV file has records separated by a text delimiter (usually a comma).
' However, in it's strictest form, all records in a CSV file will have a text
' qualifyer (usually a quote mark), in addition to a text delimiter. Text qualifyers
' are usually used when the the text delimiter is contained in the record:
' eg. an address field. Some software, like Excel, will only export a record with a
' text qualifyer when needed. See examples below:
'
' Apple,Orange Nectarine,Pear <- Regular:SPLIT works fine with this
' "Apple","Orange Nectarine","Pear" <- Strict :SPLIT can work (with additional code)
' Apple,"Orange, Nectarine",Pear <- Excel :SPLIT would cause a problem here. It
' would end up creating two records
' where should be one.
' ("Orange, Nectarine")
'
' I didn't want to think about all the possible variations so, I created this code.
' It should be able to handle anything you throw at it.
'
'
'
' Usage: MyRecords = CSVParser(CSVDataToProcess)
'
' Inputs: CSVDataToProcess - A string expression containing the CSV data to process.
'
' Returns: The function returns an array containing the records processed (ala SPLIT function).
'
' NB. Text qualifyers and delimeters are stripped from records before being
' added to the array.
'
'
' This is my first script so be gentle :) Bug reports, comments, suggestions, optimizations are
' welcome. Anyhoo.. enough rambling.. on with the script..
'*******************

 Option Explicit


 Function CSVParser (CSVDataToProcess)

   'Declaring variables for text delimiter and text qualifyer
    Dim TextDelimiter, TextQualifyer

   'Declaring the variables used in determining action to be taken
    Dim ProcessQualifyer, NewRecordCreate

   'Declaring variables dealing with input string
    Dim CharMaxNumber, CharLocation, CharCurrentVal, CharCounter, CharStorage

   'Declaring variables that handle array duties
    Dim CSVArray(), CSVArrayCount



'-------------------

   'Setting default values for various variables

   '<- Text delimiter is a comma
    TextDelimiter = ","

   '<- Chr(34) is the ascii code for "
    TextQualifyer = Chr(34)

   '<- Determining how record should be processed
    ProcessQualifyer = False

   '<- Calculating no. of characters in variable
    CharMaxNumber = Len(CSVDataToProcess)

   '<- Determining how to handle record at different
   '   stages of operation
   '   0 = Don't create new record
   '   1 = Write data to existing record
   '   2 = Close record and open new one
    NewRecordCreate = 0

   '<- Priming the array counter
    CSVArrayCount = 0

   '<- Initializing the array
    Redim Preserve CSVArray(CSVArrayCount)

   '<- Record character counter
    CharCounter = 0



'-------------------

   'Starting the main loop

    For CharLocation = 1 to CharMaxNumber

      'Retrieving the next character in sequence from CSVDataToProcess
       CharCurrentVal = Mid(CSVDataToProcess, CharLocation, 1)

      'This will figure out if the record uses a text qualifyer or not
       If CharCurrentVal = TextQualifyer And CharCounter = 0 Then
         ProcessQualifyer = True
         CharCurrentVal = ""
       End If

      'Advancing the record 'letter count' counter
       CharCounter = CharCounter + 1


      'Choosing data extraction method (text qualifyer or no text qualifyer)
       If ProcessQualifyer = True Then

          'This section handles records with a text qualifyer and text delimiter
          'It is also handles the special case scenario, where the qualifyer is
          'part of the data.  In the CSV file, a double quote represents a single
          'one  ie.  "" = "
           If Len(CharStorage) <> 0 Then
              If CharCurrentVal = TextDelimiter Then
                 CharStorage = ""
                 ProcessQualifyer = False
                 NewRecordCreate = 2
              Else
                 CharStorage = ""
                 NewRecordCreate = 1
              End If
           Else
              If CharCurrentVal = TextQualifyer Then
                 CharStorage = CharStorage & CharCurrentVal
                 NewRecordCreate = 0
              Else
                 NewRecordCreate = 1
              End If
           End If

      'This section handles a regular CSV record.. without the text qualifyer
       Else
           If CharCurrentVal = TextDelimiter Then
              NewRecordCreate = 2
           Else
              NewRecordCreate = 1
           End If

       End If


      'Writing the data to the array
       Select Case NewRecordCreate

        'This section just writes the info to the array
         Case 1
           CSVArray(CSVArrayCount) = CSVArray(CSVArrayCount) & CharCurrentVal

        'This section closes the current record and creates a new one
         Case 2
           CharCounter = 0
           CSVArrayCount = CSVArrayCount + 1
           Redim Preserve CSVArray(CSVArrayCount)

       End Select

    Next



'-------------------

   'Finishing Up

    CSVParser = CSVArray

 End Function
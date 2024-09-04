# VBA-Challenge
This code was written to get the desired outcome for the VBA Challenge in Module 2.
## Declarations
The code begins with declaring our variables and its data types. <br/><br/>
    This First Statement (after begining the code) makes sure that the code is compatible with all the worksheets in the workbook. The rest of the variables are declared with appropriate data types to perform the functions within the code efficiently. <br/>
  <pre>  Sub ModTwoSolution()
    Dim ws As Worksheet  
    Dim searchRange As Range
    Dim searchCell As Range
    Dim uniqueStrings As Collection
    Dim sortedStrings() As String
    Dim i As Long
    Dim sumF As Double
    Dim sumC As Double
    Dim sumG As Double
    Dim diff As Double
    Dim percentage As Double
    Dim outputRow As Long
    Dim firstValueC As Double
    Dim firstOccurrence As Boolean
    Dim LastOccurrenceC As Double
    Dim LastOccurrence As Boolean
    Dim highestDiff As Double
    Dim highestDiffString As String
    Dim lowestDiff As Double
    Dim lowestDiffString As String
    Dim highestvol As Double
    Dim highestvolString As String
    Dim originalDate As String
    Dim yearPart As Integer
    Dim monthPart As Integer
    Dim dayPart As Integer
    Dim convertedDate As Date</pre> 
## Initialisation 
 The following set of statemnts are setting the initial values of the Highest Percent Change, the lowest percent Change and the Total Stock Volume. These initialisations will set the value for our variables to the minimum or the maximum value possible, which will help us later in the program to print the final trends on our worksheet.<br/>
   <pre> highestDiff = -1E+308 
    highestDiffString = ""
    lowestDiff = 1E+308
    lowestDiffString = ""
    highestvol = -1E+308
    highestvolString = ""</pre>
    Sometimes the values are Initialised within the loops or inside the functions to get the required output. There will be further initialisations as required in this program.<br/>
## Loops 
The statement below will loop through each worksheet to make sure our program runs efficiently for this worksheet and then loops back for the next worksheet.<br/>
    <pre>For Each ws In ThisWorkbook.Worksheets</pre>   
### Format Date and Print in Column B
Find the Last row using : 
<pre>lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row</pre>
From row 2 to lastrow: 
<pre>For i = 2 To lastRow"</pre>
Set the Variable 'originalDate' value, which will begin with "ws.Cells(2, 2).Value"(B2): <pre>originalDate = ws.Cells(i, 2).Value</pre><br/>
YYYYMMDD to MMDDYYYY<BR/>
Now check the value stored in variable 'originalDate', which is cell B2 and see if the length of the value is 8 strings and its stored in numeric form such as "20220103":<br/>
 <pre>If Len(originalDate) = 8 And IsNumeric(originalDate) Then<br/>
                yearPart = CInt(Mid(originalDate, 1, 4))<br/>
                monthPart = CInt(Mid(originalDate, 5, 2))<br/>
                dayPart = CInt(Mid(originalDate, 7, 2))<br/>
                convertedDate = DateSerial(yearPart, dayPart, monthPart)<br/>
                ws.Cells(i, 2).Value = convertedDate<br/>
              End If<br/></pre>
Yes- set value of variable 'yearPart' to first 4 strings stored, which will be YYYY(2022) <br/>
----set value of variable 'monthPart' to the 5th value of strings stored and 2nd value after the 5th position which will be MM(01) <br/>
----set value of variable 'dayPart' to the 7th value of strings stored and 2nd value after the 7th position which will be DD(03)<br/>
----set value of variable 'convertedDate' using 'DateSerial function' to MM/DD/YYYY so that the output looks like 01/03/2022
And <br/> 
Print the value 'convertedDate' to the cell B2<br/> 
No- End the If condition and go to next statement<br/> <br/>
DDMMYYYY to MMDDYYYY<br/>
Now, check the value stored in variable 'originalDate', see if the value is in Date format such as 03/01/2022<br/>
            <pre>If IsDate(originalDate) Then
                 Dim dateParts() As String
                dateParts = Split(originalDate, "/")</pre><br>
Yes- Declare a new array called 'dateParts()' to store the date and split the date using delimeter "/". Now we have<br/>
dateParts(0)=03, dateParts(1)=01, and dateParts(2)= 2022<br/>
<br/>
Now see if 'dateParts()' array has 0,1,2 strings in it: <br/>             
            <pre> If UBound(dateParts) = 2 Then
                    Dim newDate As String
                    newDate = dateParts(1) & "/" & dateParts(0) & "/" & dateParts(2)
                    ws.Cells(i, 2).Value = newDate
                End If                
            End If
        Next i </pre><br/>
Yes- New date will be= Value stored at array 'datepart(1)'/ 'datepart(0)'/ 'datepart(2)', which when printed will give output of 01/03/2022<br/>
No- End all the if functions and go to next i, until last row.<br/>    
### Print the titles of the columns
Initialise the vallues of the cells of each column which will print the names of the columns on each worksheet:
      <pre> ws.Cells(1, "I").Value = "Ticker"
            ws.Cells(1, "J").Value = "Yearly Change"
            ws.Cells(1, "K").Value = "Percent Change"
            ws.Cells(1, "L").Value = "Total Stock Volume"            
            ws.Cells(2, "O").Value = "Greatest % Incraese"
            ws.Cells(3, "O").Value = "Greatest % decrease"
            ws.Cells(4, "O").Value = "Greatest Total Volume"            
            ws.Cells(1, "P").Value = "Ticker"
            ws.Cells(1, "Q").Value = "Value"</pre><br>
### Tickers         
Now find the last row in column A:
  <pre> lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row</pre>
And,set the column A as a ticker range:
<pre>    Set searchRange = ws.Range("A1:A" & lastRow)</pre>
Then, Initialize the collection for tickers
       <pre> Set uniqueStrings = New Collection</pre><br/>
Now, add tickers in the collection above using the statements below. If any error, ignore the error. Search each cell from column A and if the ticker 'uniqueString' is not empty, add the value from the cell to the collection.
   <pre>On Error Resume Next
        For Each searchCell In searchRange
            If searchCell.Value <> "" Then
                uniqueStrings.Add searchCell.Value, CStr(searchCell.Value)
            End If
        Next searchCell
        On Error GoTo 0</pre><br/>
### Alphabetically sort Tickers
To Sort the tickers alphabetically, check if the ticker range 'uniqueString' is not empty. If not, then declare the variable sortedStrings as Redim to accomodate the dynamic array of tickers and then sort it alphabetically using the external function 'Quicksort'. 
       <pre> If uniqueStrings.Count > 0 Then
            ReDim sortedStrings(2 To uniqueStrings.Count)
            For i = 2 To uniqueStrings.Count
                sortedStrings(i) = uniqueStrings(i)
            Next i
            QuickSort sortedStrings, LBound(sortedStrings), UBound(sortedStrings) </pre><br/>
## Calculating Yearly Change, Percent Change and Total Stock Volume
### Total Stock Volume
To print it from A to Z, we use LBound for the smallest value to Ubound for the highest value in sortedStrings.<br/> While printing the tickers alphabetically, let's do the calculations using For loop. So for A2 (Main For loop(i=2) and searchRange A), where output will be in second row of the column, and For A2=ticker (nested loop), we initialise  the variables 'sumG' to store 'Total stock Volume', 'firstValueC' to store 'opening value', 'LastOccurrenceC' to store the 'closing value', 'diff' to store the 'yearly change' to 0 and set the 'firstOccurrence' to False for that particular ticker.
      <pre> Print the sorted tickers to column I
            outputRow = 2
            For i = LBound(sortedStrings) To UBound(sortedStrings)  
                sumG = 0
                firstOccurrence = False
                firstValueC = 0
                LastOccurrenceC = 0
                diff = 0</pre><br/>
Now, to calculate 'Total stock Volume' for the tickers, we search each cell in Ticker range. <br/>If the ticker value same as the sortedString the 'sumG' will add new value from column "G" in itself.  
             <pre>   For Each searchCell In searchRange
                    If searchCell.Value = sortedStrings(i) Then                        
                        sumG = sumG + ws.Cells(searchCell.Row, "G").Value</pre>
### Find Opening and Closing Values
Then find opening and closing value for the ticker using the code below. As we know the firstOccurrence is initially set to False. So if firstOccurrence is not false, then firstValueC= the corresponding value against that ticker in column "C". Similarly, we set LastOccurence to 0 and the value of this variable will continue to change untill the last value in column F, which will provide us with the 'Closing Value' for that ticker. then the nested loop will continue for each cell until the ticker is changed.
                   <pre>     If Not firstOccurrence Then
                            firstValueC = ws.Cells(searchCell.Row, "C").Value
                            firstOccurrence = True
                            End If
                        LastOccurrenceC = ws.Cells(searchCell.Row, "F").Value                            
                    End If
                Next searchCell</pre>
### Yearly Change and Percent Change
Now, lets find the Yearly Change by subtracting the opening value from the closing value. The percent change can be found by dividing the Yearly Change by opening value and then multiplying it by 100. We are using elseif condition to avoid the division by 0.
        <pre> ' Calculate the difference and percentage
                diff = LastOccurrenceC - firstValueC
                If diff <> 0 Then
                     percentage = (diff / firstValueC) * 100
                Else
                    percentage = 0 ' Avoid division by zero
                End If</pre>
### Printing the Yearly Change, Percent Change and Total Stock Volume
We can now simply print the values calculated above in appropriate columns using the set of statements below:
<pre>          ' Print the results in columns I to L
                ws.Cells(outputRow, "I").Value = sortedStrings(i)
                ws.Cells(outputRow, "J").Value = diff
                ws.Cells(outputRow, "K").Value = percentage & "%"
                ws.Cells(outputRow, "L").Value = sumG '</pre>
### Changing Cell Format
Here, the Yearly Change positive or negative or none, will format our cells with Green,red and white background, respectively.<br/>
<pre>                ' Format Cells based on yearly change
                If diff > 0 Then
                    ws.Cells(outputRow, "J").Interior.Color = RGB(0, 255, 0) ' Green for positive
                ElseIf diff < 0 Then
                    ws.Cells(outputRow, "J").Interior.Color = RGB(255, 0, 0) ' Red for negative
                Else
                    ws.Cells(outputRow, "J").Interior.ColorIndex = xlNone ' No color for zero
                End If</pre>
## Greatest +ve/-ve Percent Change and Stock Volume
### Greatest % Increase and Decrease
To calculate the Greatest % Increase and Decrease, we have alraedy initialised the 'highestDiff' to store the value of Greatest Increase to -1E+308, which is the lowest number. Similarly, we used lowestDiff to store the value of Greatest Decrease to 1E+308 , which is the highest number. <br/>
Now, this set of statement will calculate if percent change is higher or lower than the value we initialised. if the value is true, our new Greatest Increase and Decreased will be replaced by the percentage, which will depict the highest or lowest percent change. <br/>
Whereas  highestDiffString and lowestDiffString have been assigned with "" , which means no string value. These variable will accomodate the ticker of the Greatest increase and decrease
    <pre>            ' Check for the highest/lowest value
                If percentage > highestDiff Then
                highestDiff = percentage
                highestDiffString = sortedStrings(i)                
                ElseIf percentage < lowestDiff Then
                lowestDiff = percentage
                lowestDiffString = sortedStrings(i)                
                End If</pre>
### Total Stock Volume
Similarly, if highestvol has been initialised to -1E+308. If sumG is greater that the value assigned, then new highestvol will be sumG and highestvolString will store the corresponding ticker.             
          <pre> If sumG > highestvol Then
                highestvol = sumG
                highestvolString = sortedStrings(i)
                End If</pre>
## Output
This statement will increase the value of the output row by 1. Thus, all the calculations will  loop again for the next row for the whole worksheet. Once Terminated, the Greatest % Increase will be printed in P2 and its strings will be printed in Q2.<br/>
Similarly, we can see the final output of Greatest % Decrease, Greatest Stock Volume, and its Tickers in P3,P4,Q3,Q4, respectively. Then the whole program will loop through the next worksheet.
     <pre>           outputRow = outputRow + 1
            Next i
        End If
        ws.Cells(2, "P").Value = highestDiffString
        ws.Cells(2, "Q").Value = highestDiff
        ws.Cells(3, "P").Value = lowestDiffString
        ws.Cells(3, "Q").Value = lowestDiff
        ws.Cells(4, "P").Value = highestvolString
        ws.Cells(4, "Q").Value = highestvol    
    Next ws  

End Sub</pre>
# External Function
This is an external function to sort the tickers alphabetically<br/>
#### Declaration and initialisation
<pre>Sub QuickSort(arr() As String, first As Long, last As Long)
     Dim low As Long, high As Long
     Dim midValue As String, temp As String
     low = first
    high = last</pre>
Whenever called, it will take the first find the middle of the array using
 <pre>  midValue = arr((first + last) \ 2)</pre> 
 Then while the first value is lower than the last value and while first value is also lower than the middle value, it continues to loop to create list of strings of lower than the mid value.
  <pre>  Do While low <= high
        Do While arr(low) < midValue
            low = low + 1
        Loop</pre>
In the same manner as above, a list of strings is created which are higher than the middle value using:
<pre>Do While arr(high) > midValue
            high = high - 1
        Loop</pre>
Then using the following calculations, the strings at low are incremented and at high are decremented, sort the entire array of strings from low to high.

<pre>If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub </pre>
## References
<pre>https://bootcampspot.instructure.com/courses/6174/external_tools/313</pre>
<pre> https://stackoverflow.com/questions/26587527/cite-a-paper-using-github-markdown-syntax</pre>
<pre>https://stackoverflow.com/questions/26587527/cite-a-paper-using-github-markdown-syntax</pre>
<pre>https://www.youtube.com/watch?v=MCo1UtflJHM</pre>
<pre>https://www.youtube.com/watch?v=Pm6fApmnOaQ&t=95s</pre>
<pre>https://www.youtube.com/watch?v=BE45npPXtH8</pre>
<pre>https://www.testunlimited.com/pdf/an/E5061-90033.pdf</pre>
<img width="455" alt="image" src="https://github.com/user-attachments/assets/70e88af1-f6ae-42d6-a29e-0724a1bbc2d1">
<pre>https://openai.com/chatgpt/</pre>




Module Module1
    '   Programming Assignment Two: Grade Report
    '                              By: David Rees
    '                                 11/04/18

    ' Classroom View
    ' 
    ' The If Statement:
    ' If A = B Then
    '   (Function or Process)
    ' End If

    ' Business View
    ' This program is designed to take a grade sheet from KVCC then print out the
    ' Student's Name, First Course ID, Credit Hours, Grade, and Grade Points then their Second 
    ' Course ID, Credit Hours, Grade and Grade Points then the GPA for the two courses. After 
    ' that it will print out the Total Number of Students, Credit Hours, and Grade Points. Finally
    ' it will print the Overall GPA for all of the students. This will all be formatted according
    ' to the Printer Spacing Chart provided to the programmer. It will be rounded To three Decimal
    ' points

    '   Input Variables
    '       From Data File

    '   Student Personal Info
    Private StudentNameString As String

    '   Student Course IDs
    Private CourseIDOneString As String
    Private CourseIDTwoString As String

    '   Student Credit Hours Info
    Private CreditHoursInteger As Integer         ' .ToString()
    Private CreditHours2Integer As Integer        ' .ToString()

    '   Student Grade Info
    Private GradeOneDecimal As Decimal            ' .ToString("N1")
    Private GradeTwoDecimal As Decimal            ' .ToString("N1")

    '   End Of Input Variables
    '       From Data File


    '   Processing Variables
    Private CurrentRecordString() As String
    Private AccumPageCounterInteger As Integer = 1
    Private AccumLineCounterInteger As Integer = 99
    Private Const MAX_LINE_COUNT_INTEGER As Integer = 20
    '   End Of Proccessing Variables

    '   Output Variables
    Private GradePointsOneDecimal As Decimal      ' .ToString("N1")
    Private GradePointsTwoDecimal As Decimal      ' .ToString("N1")
    Private GradePointAverageDecimal As Decimal   ' .ToString("N")

    Private AccumTotalStudentCountInteger As Integer   ' .ToString()
    Private AccumTotalCreditHoursInteger As Integer    ' .ToString()
    Private AccumTotalGradePointsDecimal As Decimal    ' .ToString("N1")
    Private AccumOverallGPADecimal As Decimal          ' .ToString("N3")
    '   End Of Output Variables

    ' Defines GradeFile as the File Location and the Type of File
    Private GradeFile As New Microsoft.VisualBasic.FileIO.TextFieldParser("GRADPNTF18.TXT")

    Sub Main()
        Call HouseKeeping()
        Do While Not (GradeFile).EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Sub HouseKeeping()
        Call SetFileDelimiters()
    End Sub

    Sub ProcessRecords()
        Call ReadFile()
        Call DetailCalculation()
        Call AccumulateTotals()
        Call WriteDetailLine()
    End Sub

    Sub EndOfJob()
        Call SummaryCalculation()
        Call SummaryOutput()
        Call CloseFile()
    End Sub

    Sub SetFileDelimiters()

        ' Sets File Field Type to Delimited
        GradeFile.TextFieldType = FileIO.FieldType.Delimited

        ' Delimiter is set to character within quotes. In this case character is a comma.
        GradeFile.SetDelimiters(",")

    End Sub

    Sub ReadFile()

        ' ReadLine
        CurrentRecordString = GradeFile.ReadFields()

        ' Name
        StudentNameString = CurrentRecordString(0)

        ' Course One ID
        CourseIDOneString = CurrentRecordString(1)

        ' Credit Hours One
        CreditHoursInteger = CurrentRecordString(2)

        ' Grade One
        GradeOneDecimal = CurrentRecordString(3)

        ' Course Two ID
        CourseIDTwoString = CurrentRecordString(4)

        ' Credit Hours Two
        CreditHours2Integer = CurrentRecordString(5)

        ' Grade Two
        GradeTwoDecimal = CurrentRecordString(6)

    End Sub

    Sub DetailCalculation()

        ' Course One Grade Points Calculation
        GradePointsOneDecimal = CreditHoursInteger * GradeOneDecimal

        ' Course Two Grade Points Calculation
        GradePointsTwoDecimal = CreditHours2Integer * GradeTwoDecimal

        ' Student GPA Calculation
        GradePointAverageDecimal = (GradePointsOneDecimal + GradePointsTwoDecimal) / (CreditHoursInteger + CreditHours2Integer)

        ' If statement to decide if new page is needed.
        If AccumLineCounterInteger >= 20 Then

            ' Calls Write Headings for a new page
            Call WriteHeadings()

            ' Adds One to the Page Counter so the proper number is printed.
            AccumPageCounterInteger += 1

            ' Resets the line counter so Write Headings isn't repeated more than necessary.
            AccumLineCounterInteger = 0

        End If
        ' (End If) Ends the If Statement
    End Sub

    Sub AccumulateTotals()

        ' Counts lines until equals 20 or more
        AccumLineCounterInteger += 1

        ' Counts the total number of records or students
        AccumTotalStudentCountInteger += 1

        ' Adds the CreditHours from both courses of the current record to the Total Accum Credit Hours Variable.
        AccumTotalCreditHoursInteger += (CreditHoursInteger + CreditHours2Integer)

        ' Adds the Grade Points from both courses of the current record to the Total Accum Grade Points Variable.
        AccumTotalGradePointsDecimal += (GradePointsOneDecimal + GradePointsTwoDecimal)

    End Sub

    Sub WriteHeadings()

        ' Blank Line for Spacing
        Console.WriteLine()

        ' Printer Spacing Chart Line Number 01
        Console.WriteLine(Space(25) & "STUDENT GRADE POINT REPORT" & Space(18) & "Page: " & AccumPageCounterInteger.ToString().PadLeft(3))

        ' Printer Spacing Chart Line Number 02
        Console.WriteLine(Space(31) & "BY: David Rees")

        ' Printer Spacing Chart Line Number 03
        Console.WriteLine()

        ' Printer Spacing Chart Line Number 04
        Console.WriteLine(Space(14) & "-------- Course #1 -------" & Space(4) & "-------- Course #2 -------")

        ' Printer Spacing Chart Line Number 05
        Console.WriteLine("Student" & Space(7) &
                          "Course" & Space(2) &
                          "Credit" & Space(7) &
                          "Grade" & Space(4) &
                          "Course" & Space(2) &
                          "Credit" & Space(7) &
                          "Grade" & Space(4) &
                          "Grade")

        ' Printer Spacing Chart Line Number 06
        Console.WriteLine("Name" & Space(10) &
                          "ID #1" & Space(6) &
                          "Hrs Grade" & Space(3) &
                          "Pts" & Space(4) &
                          "ID #2" & Space(6) &
                          "Hrs Grade" & Space(3) &
                          "Pts" & Space(3) &
                          "Pt Avg")

        ' Printer Spacing Chart Line Number 07
        Console.WriteLine()

    End Sub

    Sub WriteDetailLine()

        ' Printer Spacing Chart Detail Line
        Console.WriteLine(StudentNameString.PadRight(12) &                      ' Name Field 
                          Space(2) &                                            ' Spacing
                          CourseIDOneString.PadLeft(6) &                        ' Course One ID Field
                          Space(7) &                                            ' Spacing
                          CreditHoursInteger.ToString.PadLeft(1) &              ' Credit Hours Course One Field
                          Space(3) &                                            ' Spacing
                          GradeOneDecimal.ToString("N1").PadLeft(3) &           ' Grade Course One Field
                          Space(2) &                                            ' Spacing
                          GradePointsOneDecimal.ToString("N1").PadLeft(4) &     ' Grade Points Course One Field
                          Space(4) &                                            ' Spacing & End Of Course One Info
                          CourseIDTwoString.PadLeft(6) &                        ' Course Two ID
                          Space(7) &                                            ' Spacing
                          CreditHours2Integer.ToString.PadLeft(1) &             ' Credit Hours Course Two Field
                          Space(3) &                                            ' Spacing
                          GradeTwoDecimal.ToString("N1").PadLeft(3) &           ' Grade Course Two Field
                          Space(2) &                                            ' Spacing
                          GradePointsTwoDecimal.ToString("N1").PadLeft(4) &     ' Grade Points Course Two Field
                          Space(4) &                                            ' Spacing & End Of Course Two Info
                          GradePointAverageDecimal.ToString("N2").PadLeft(5))   ' GPA Field = (GradePointsOne + GradePointsTwo) / (CreditHoursOne + CreditHoursTwo)

    End Sub

    Sub SummaryCalculation()

        ' When run, it divides the total grade points by the total credit hours.
        AccumOverallGPADecimal = AccumTotalGradePointsDecimal / AccumTotalCreditHoursInteger

    End Sub

    Sub SummaryOutput()

        ' Printer Spacing Chart Line(11)
        Console.WriteLine()

        ' Printer Spacing Chart Line(12)
        Console.WriteLine("FINAL TOTALS:")

        ' Printer Spacing Chart Line(13)
        Console.WriteLine(Space(5) & "Number of Students" & Space(13) & AccumTotalStudentCountInteger.ToString().PadLeft(2))

        ' Printer Spacing Chart Line(14)
        Console.WriteLine()

        ' Printer Spacing Chart Line(15)
        Console.WriteLine(Space(5) & "All Credit Hours" & Space(14) & AccumTotalCreditHoursInteger.ToString().PadLeft(3))

        ' Printer Spacing Chart Line(16)
        Console.WriteLine(Space(5) & "All Grade Points" & Space(14) & AccumTotalGradePointsDecimal.ToString("N1").PadLeft(5))

        ' Printer Spacing Chart Line(17)
        Console.WriteLine()

        ' Printer Spacing Chart Line(18)
        Console.WriteLine(Space(5) & "Overall Grade Point Average" & Space(3) & AccumOverallGPADecimal.ToString("N3").PadLeft(7))

    End Sub

    Sub CloseFile()

        ' Writes two extra blank lines
        Console.WriteLine()
        Console.WriteLine()

        ' Writes line to prompt user to close program.
        Console.WriteLine(
            Space(10) &
            "Press -ENTER- TO EXIT")
        Console.ReadLine()
    End Sub
End Module
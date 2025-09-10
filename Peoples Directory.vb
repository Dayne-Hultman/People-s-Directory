Imports System
Imports System.Globalization
Imports System.Net.Http
Imports System.Linq


Public Class Person
    Public Property firstName As String
    Public Property middleName As String
    Public Property lastName As String

    Public Property address As String
    Public Property age As String
    Public Property heightInches As String
    Public Property weight As String

    Private _value As String


    Public Function getFullInfo() As String
        Return firstName + " " + middleName + " " + lastName + " is " + age + " years old, has a height of " + heightInches + " inches, weighs " + weight + " pounds, while living at " + address + "."
    End Function


End Class

Module intro
    Function DisplayIntro()
        'Console.Clear()
        Console.WriteLine("Welcome to the People's Directory. Here you can find new information about a person or add some.")
        Console.WriteLine("Now what would you like to do today?")
        Console.WriteLine("To view all people in the directory, type 1 and press enter.")
        Console.WriteLine("To add someone to the directory, type 2 and press enter.")
        Console.WriteLine("To find someone based on specific criteria, type 3 and press enter.")
        Console.WriteLine("To quit the application, type 4 and press enter.")
        Console.Write("Please enter your choice: ")
        Return 0
    End Function

    Function CapitalFirst(word As String)
        If word.Length > 0 Then
            word = word.Substring(0, 1).ToUpper() & word.Substring(1)
        Else
            word = "" ' Handle empty string case
        End If
        Return word
    End Function

End Module

Module Program
    Sub Main(args As String())

        Console.SetWindowSize(130, 40)
        Console.SetBufferSize(130, 40)

        Dim per1 As New Person()
        per1.firstName = "John"
        per1.middleName = "Glass"
        per1.lastName = "Reed"
        per1.address = "374 Wishire Lane"
        per1.age = "34"
        per1.heightInches = "67"
        per1.weight = "164.75"
        Dim per2 As New Person()
        per2.firstName = "Amanda"
        per2.middleName = "Stone"
        per2.lastName = "Thompson"
        per2.address = "8900 Lancaster Road"
        per2.age = "55"
        per2.heightInches = "55"
        per2.weight = "110.05"
        Dim per3 As New Person()
        per3.firstName = "Masaq"
        per3.middleName = "Chadri"
        per3.lastName = "Ubuluq"
        per3.address = "1047 Collarbone Drive"
        per3.age = "24"
        per3.heightInches = "73"
        per3.weight = "155.34"
        Dim per4 As New Person()
        per4.firstName = "Tracy"
        per4.middleName = "Heather"
        per4.lastName = "Bloom"
        per4.address = "8900 Lancaster Road"
        per4.age = "21"
        per4.heightInches = "52"
        per4.weight = "122.51"

        Dim people As New List(Of Person)
        people.Add(per1)
        people.Add(per2)
        people.Add(per3)
        people.Add(per4)

        Dim choice As String = "0"

        intro.DisplayIntro()


        While choice <> "4"
            choice = Console.ReadLine
            Console.WriteLine("")

            While choice <> "1" And choice <> "2" And choice <> "3" And choice <> "4"
                Console.WriteLine("Please enter a valid choice, remember: ")
                Console.WriteLine("To view all people in the directory, type 1 and press enter.")
                Console.WriteLine("To add someone to the directory, type 2 and press enter.")
                Console.WriteLine("To find someone based on specific criteria, type 3 and press enter.")
                Console.WriteLine("To quit the application, type 4 and press enter.")
                Console.Write("Please enter your choice: ")
                choice = Console.ReadLine
                Console.WriteLine()
            End While

            If choice = "1" Then
                Console.Clear()
                For nums As Integer = 0 To (people.Count - 1)
                    Console.WriteLine(people(nums).getFullInfo())
                Next

                Console.WriteLine()
                Console.WriteLine("Here are all the displayed people in the directory. I hope you find what you are looking for.")
                Console.WriteLine()
                Console.WriteLine()
                intro.DisplayIntro()
            ElseIf choice = "2" Then
                Dim newPerson As New Person
                Console.Clear()
                Console.Write("Please enter the persons first name: ")
                newPerson.firstName = Console.ReadLine
                newPerson.firstName = CapitalFirst(newPerson.firstName)
                Console.Write("Please enter the persons middle name: ")
                newPerson.middleName = Console.ReadLine
                newPerson.middleName = CapitalFirst(newPerson.middleName)
                Console.Write("Please enter the persons last name: ")
                newPerson.lastName = Console.ReadLine
                newPerson.lastName = CapitalFirst(newPerson.lastName)
                Console.Write("Please enter the persons address: ")
                newPerson.address = Console.ReadLine
                Console.Write("Please enter the persons age: ")
                newPerson.age = Console.ReadLine

                While IsNumeric(newPerson.age) = False
                    Console.WriteLine("I am sorry but the data you have entered is not in the correct format. Try again.")
                    Console.Write("Please enter the persons age: ")
                    newPerson.age = Console.ReadLine
                End While

                Console.Write("Please enter the persons height in Inches: ")
                newPerson.heightInches = Console.ReadLine

                While IsNumeric(newPerson.heightInches) = False
                    Console.WriteLine("I am sorry but the data you have entered is not in the correct format. Try again.")
                    Console.Write("Please enter the persons height in inches: ")
                    newPerson.heightInches = Console.ReadLine
                End While

                Console.Write("Please enter the persons weight in pounds: ")
                newPerson.weight = Console.ReadLine

                While IsNumeric(newPerson.weight) = False
                    Console.WriteLine("I am sorry but the data you have entered is not in the correct format. Try again.")
                    Console.Write("Please enter the persons weight: ")
                    newPerson.weight = Console.ReadLine
                End While
                Console.Clear()
                Console.WriteLine("Thank you for your entry and updating our database.")
                Console.WriteLine()

                people.Add(newPerson)

                intro.DisplayIntro()
                Console.WriteLine()

            ElseIf choice = "3" Then
                Dim searchCrit As String = "0"
                Dim searched As String = ""
                Dim found As Boolean = False
                Console.Clear()
                Console.WriteLine("Sometimes, you may got more than one instance.")
                Console.WriteLine("To find a specific person, please choose which category to sort by: ")
                Console.WriteLine("To search by First name, type 1 and press Enter.")
                Console.WriteLine("To search by Middle name, type 2 and press Enter.")
                Console.WriteLine("To search by Last name, type 3 and press Enter.")
                Console.WriteLine("To search by address, type 4 and press Enter.")
                Console.WriteLine("To search by age, type 5 and press Enter.")
                Console.WriteLine("To search by height, type 6 and press Enter.")
                Console.WriteLine("To search by weight, type 7 and press Enter.")
                Console.Write("Please select your choice: ")
                searchCrit = Console.ReadLine()

                While searchCrit <> "1" And searchCrit <> "2" And searchCrit <> "3" And searchCrit <> "4" And searchCrit <> "5" And searchCrit <> "6" And searchCrit <> "7"
                    Console.Write("I am sorry. That is not a valid choice. Try again: ")
                    searchCrit = Console.ReadLine()
                End While
                Console.WriteLine()
                Console.WriteLine()

                Select Case searchCrit
                    Case "1"
                        Console.Write("Please enter the person's first name: ")
                        searched = Console.ReadLine()
                        Console.Clear()
                        For num2 As Integer = 0 To (people.Count - 1)
                            If people(num2).firstName = searched Then
                                found = True
                                Console.WriteLine("We have found the following: ")
                                Console.WriteLine(people(num2).getFullInfo())
                                Console.WriteLine()
                            End If
                        Next

                    Case "2"
                        Console.Write("Please enter the person's middle name: ")
                        searched = Console.ReadLine()
                        Console.Clear()
                        For num2 As Integer = 0 To (people.Count - 1)
                            If people(num2).middleName = searched Then
                                found = True
                                Console.WriteLine("We have found the following: ")
                                Console.WriteLine(people(num2).getFullInfo())
                                Console.WriteLine()
                            End If
                        Next

                    Case "3"
                        Console.Write("Please enter the person's last name: ")
                        searched = Console.ReadLine()
                        Console.Clear()
                        For num2 As Integer = 0 To (people.Count - 1)
                            If people(num2).lastName = searched Then
                                found = True
                                Console.WriteLine("We have found the followiing: ")
                                Console.WriteLine(people(num2).getFullInfo())
                                Console.WriteLine()
                            End If
                        Next

                    Case "4"
                        Console.Write("Please enter the person's address: ")
                        searched = Console.ReadLine()
                        Console.Clear()
                        For num2 As Integer = 0 To (people.Count - 1)
                            If people(num2).address = searched Then
                                found = True
                                Console.WriteLine("We have found the following: ")
                                Console.WriteLine(people(num2).getFullInfo())
                                Console.WriteLine()
                            End If
                        Next

                    Case "5"
                        Console.Write("Please enter the person's age: ")
                        searched = Console.ReadLine()
                        Console.Clear()
                        While IsNumeric(searched) = False
                            Console.WriteLine("I am sorry that is not a acceptable parameter. Please enter numbers only.")
                            Console.Write("Please enter the person's age: ")
                            searched = Console.ReadLine()
                            Console.Clear()
                        End While

                        For num2 As Integer = 0 To (people.Count - 1)
                            If people(num2).age = searched Then
                                found = True
                                Console.WriteLine("We have found the following: ")
                                Console.WriteLine(people(num2).getFullInfo())
                                Console.WriteLine()
                            End If
                        Next
                    Case "6"
                        Console.Write("Please enter the person's height in inches: ")
                        searched = Console.ReadLine()
                        Console.Clear()
                        While IsNumeric(searched) = False
                            Console.WriteLine("I am sorry that is not a acceptable parameter. Please enter numbers only.")
                            Console.Write("Please enter the person's height in inches: ")
                            searched = Console.ReadLine()
                            Console.Clear()
                        End While

                        For num2 As Integer = 0 To (people.Count - 1)
                            If people(num2).heightInches = searched Then
                                found = True
                                Console.WriteLine("We have found the following: ")
                                Console.WriteLine(people(num2).getFullInfo())
                                Console.WriteLine()
                            End If
                        Next

                    Case "7"
                        Console.Write("Please enter the person's weight: ")
                        searched = Console.ReadLine()
                        Console.Clear()
                        While IsNumeric(searched) = False
                            Console.WriteLine("I am sorry that is not a acceptable parameter. Please enter numbers only.")
                            Console.Write("Please enter the person's weight: ")
                            searched = Console.ReadLine()
                            Console.Clear()
                        End While

                        For num2 As Integer = 0 To (people.Count - 1)
                            If people(num2).weight = searched Then
                                found = True
                                Console.WriteLine("We have found the following: ")
                                Console.WriteLine(people(num2).getFullInfo())
                                Console.WriteLine()
                            End If
                        Next
                End Select
                If found = False Then
                    Console.WriteLine("I am sorry. We were unable to find the person you are looking for.")
                    Console.WriteLine("Please try again and check to make sure you have the right information.")
                    Console.WriteLine()
                    Console.WriteLine()
                    intro.DisplayIntro()
                Else
                    Console.WriteLine("Here is the information we have gathered based on your input.")
                    Console.WriteLine("We hope you have found what you are looking for.")
                    Console.WriteLine()
                    Console.WriteLine()
                    intro.DisplayIntro()
                End If
            End If

        End While


        Console.WriteLine("Thank you for using the People's Directory. Have a great day!")

    End Sub
End Module

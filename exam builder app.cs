using System;
using System.IO;
using System.Globalization;
using IronXL;



namespace Exam_Automation
{

    class Functions
    {
        public int getInt()
        {
            // Get user input, use temp to hold initial and result to store int from temp
            string temp = Console.ReadLine();
            int result;

            // Try and parse temp to see if it is an int
            if (int.TryParse(temp, out result))
            {
                return result;
            }
            else // If not int then use recursion
            {
                Console.WriteLine("That wasn't an integer, please try again.");
                return getInt();
            }

        }

        public DateTime getDate()
        {
            // Get user input
            string temp = Console.ReadLine();
            DateTime result;

            // Try and parse the date
            if (DateTime.TryParse(temp, out result))
            {
                return result;
            }
            else // If not int then use recursion
            {
                Console.WriteLine("That wasn't a valid date, please try again.");
                return getDate();
            }
        }

    }
    class Program
    {

        static void Main(string[] args)
        {


            // Init functions
            Functions fn = new Functions();

            // Get template sheet data
            WorkBook wb = WorkBook.Load($"J:/Lanschool Utility/IT Special Use/Exam Account Sheets/Exam Account Sheet template - do not delete.xlsx"); // Import base spreadsheet
            WorkSheet ws = wb.GetWorkSheet("Sheet1"); // Get the correct sheet

            // Get user data

            // Give instructions for inputs
            Console.WriteLine("Inputs cannot include: '#, %, &, {}, <>, *, ?, $, !, \'\', \"\", +, `, |, =, :' \n\n");

            // Ask year level
            Console.WriteLine("\nWhat is the year level of the exam as an integer?");
            // Get year level input
            int yrLvl = fn.getInt();

            // Ask exam type
            Console.WriteLine("\nWhat is the type of the exam?");
            // Get type input
            string type = Console.ReadLine();


            // Ask exam location
            Console.WriteLine("\nWhat is the location of the exam?");
            // Get location input
            string location = Console.ReadLine();


            // Ask exam teacher
            Console.WriteLine("\nWho is the teacher of the exam?");
            // Get teacher input
            string teacher = Console.ReadLine();


            // Ask date of exam
            Console.WriteLine("\nWhat is the date of the exam in dd/mm/yyyy format?");
            DateTime examDate = fn.getDate();


            // Get number of students
            Console.WriteLine("\nEnter the amount of students in the exam as an integer.");
            int numStudents = fn.getInt();


            // Get exam account start location as integer
            Console.WriteLine("\nEnter an integer for the first exam account in this exam?");
            int firstExamAccount = fn.getInt();



            // Logic
            // Create the title/filename
            string title = $"{teacher} - Yr {yrLvl} - {type} - {examDate.DayOfWeek} {examDate.Day} " +
                $"{CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(examDate.Month)}";

            // Title cell is row 4, column j (array[3], array[9])
            ws.Rows[3].Columns[9].Value = $"{title}";
            ws.Rows[11].Columns[3].Value = $"{title}";

            // Create string array for the lsc file
            string[] lsc = new string[numStudents + 1];
            lsc[0] = "[LanSchool Class List]";



            // Input exam accounts into spreadsheet
            for (int i = 0; i <= numStudents; i++)
            {
                // Create and edit the exam excel spreadsheet
                // See if the exam accounts will be in column 1
                if (i <= 12)
                {

                    ws.Rows[(11 + i)].Columns[3].Value = $"Exam0{firstExamAccount + i}";

                } // Check to see if exam account will be in column 2 
                else if (i > 12 && i <= 25)
                {
                    ws.Rows[i - 2].Columns[10].Value = $"Exam{firstExamAccount + i}";
                } // Check to see if exam account will be in column 3 
                else if (i > 25 && i <= 38)
                {
                    ws.Rows[i - 15].Columns[17].Value = $"Exam{firstExamAccount + i}";
                }
                else if (i > 38 && i <= 51)
                {
                    ws.Rows[i - 28].Columns[24].Value = $"Exam{firstExamAccount + i}";
                }
                else
                {
                    // Do nothing
                }
            }

            // For loop for creating and writing the LSC file
            for (int i = 0; i <= numStudents; i++)
            {
                if (i == 0)
                {
                    // Do nothing
                }
                else if (i < 10)
                {
                    // check if exam account is less than 10 and add leading 0 if it is
                    // Take 1 to account for the [lanschool class list] header
                    lsc[i] = $"student_{i - 1}=exam0{firstExamAccount + i - 1}";
                }
                else
                {
                    // Take 1 to account for the [lanschool class list] header
                    lsc[i] = $"student_{i - 1}=exam{firstExamAccount + i - 1}";
                }
            }

            StreamWriter sw = new StreamWriter($"J:/Lanschool Utility/IT Special Use/Exam Account Lists/{title}.lsc", true);
            foreach (string line in lsc)
            {
                Console.WriteLine(line);
                sw.WriteLine(line);

            }
            sw.Close();


            // Save the new exam spread sheet
            wb.SaveAs($"J:/Lanschool Utility/IT Special Use/Exam Account Sheets/{title}.xlsx");
        }
    }
}

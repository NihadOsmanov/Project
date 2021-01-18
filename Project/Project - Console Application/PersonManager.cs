using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;

namespace Project___Console_Application
{
    class PersonManager
    {
        public static List<Person> personList = new List<Person>();
        public static List<Person> tempList = new List<Person>();
        public static void AddPerson()
        {
            Console.WriteLine("Please, enter name");
            string name = Console.ReadLine();

            Console.WriteLine("Please, enter age");
            int age = Convert.ToInt32(Console.ReadLine());

            Person person = new Person(name, age);

            personList.Add(person);
        }
        public static void SearchPerson()
        {
            Console.WriteLine($"What are you searching : \na) ada gore \nb) yasa gore");
            string choice = Console.ReadLine();
           

            if (choice == "a")
            {
                Console.WriteLine("Please, enter name");
                string name = Console.ReadLine();

                int counter = 1;
                tempList = personList.Where(person => person.Name == name).ToList();
                foreach (var item in tempList)
                {
                    Console.WriteLine($"{counter}, {item.Name}, {item.Age}");
                    counter++;
                }
            }
            else if (choice == "b")
            {
                Console.WriteLine("Please, enter age");
                int age = Convert.ToInt32(Console.ReadLine());

                int counter = 1;
                tempList = personList.Where(person => person.Age == age).ToList();
                foreach (var item in tempList)
                {
                    Console.WriteLine($"{counter}, {item.Name}, {item.Age}");
                    counter++;
                }
            }
            Console.Write("---");
                Console.Write(Environment.NewLine);

            if (tempList.Count() > 0)
            {
                int i = 1;
                foreach (var item in PersonManager.tempList)
                {
                    Console.WriteLine($"{i}, {item.Name}, {item.Age}");
                    i++;
                }
                Console.WriteLine("Edit ve ya Delete etmek ucun indexsi daxil et");
                Console.Write(Environment.NewLine);

                string index = Console.ReadLine();
                EditOrDelete(index);
            }
            else
            {
                Console.Write("---");
                Console.Write(Environment.NewLine);


                Console.WriteLine("File tapilmadi");
            }

        }
        public static void EditOrDelete(string index)
        {
            int i = Convert.ToInt32(index);

            Console.Write(personList.ElementAt(i - 1).Name);
            Console.Write(Environment.NewLine);

            Console.WriteLine($" \n1) Edit et \n2) Delete et");


            string editOrDelete = Console.ReadLine();

            if (editOrDelete == "1")
            {
                Console.WriteLine($"\na: Name to \nb: Age to");


                string choice = Console.ReadLine();

                if (choice == "a")
                {
                    Console.Write("Name write: ");
                    string name = Console.ReadLine();
                    var person = personList.ElementAt(i - 1);
                    person.Name = name;
                }
                else if (choice == "b")
                {
                    Console.Write("Age write: ");
                    int age = Convert.ToInt32(Console.ReadLine());

                    var person = personList.ElementAt(i - 1);
                    person.Age = age;
                }

                Console.Write("---");
                Console.Write(Environment.NewLine);
                Console.WriteLine("Update edildi!");

                Console.WriteLine($"{ personList.ElementAt(i - 1).Name}, {personList.ElementAt(i - 1).Age }");

            }
            else if (editOrDelete == "2")
            {
                personList.RemoveAt(i - 1);

                Console.Write("---");
                Console.Write(Environment.NewLine);
                Console.WriteLine("Delete edildi!");
            }
        }
        public static void SavedData()
        {
            DateTime dateTime = DateTime.Now;
            DirectoryInfo directoryInfo = Directory.CreateDirectory($"Saved--{dateTime.ToString("dd--MM--yyyy")}");
            BinaryFormatter formatter = new BinaryFormatter();

            Console.WriteLine("Provide filename");
            string fileName = Console.ReadLine();

            using (FileStream fs = new FileStream(Path.Combine(directoryInfo.Name, $"{fileName}.dat"), FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, personList.ToArray());
            }
            
        }
        public static void LoadData()
        {
            Console.WriteLine("Please pr filename");
            string file_Name = Console.ReadLine();
            string searchFile = "";

            foreach (var item in Directory.GetDirectories(Environment.CurrentDirectory))
            {
                searchFile = Path.GetFullPath(Path.Combine(item, file_Name));
                if (searchFile != null)
                {
                    break;
                }
            }
            if (searchFile != null)
            {
                BinaryFormatter formatter = new BinaryFormatter();
                using (FileStream fs = new FileStream(searchFile + ".dat", FileMode.OpenOrCreate))
                {
                     Person[] deserailizePeople = (Person[])(formatter.Deserialize(fs));
                    personList = deserailizePeople.ToList();
                }
            }
        }

        public static void ExportData()
        {
            Console.WriteLine("What format to print : \n1) Text  \n2) Excel");
            string export = Console.ReadLine();
            if(export == "1")

            {
                using(TextWriter tw = new StreamWriter($"C:\\Users\\Administrator\\source\\repos\\Project - Console Application\\Project - Console Application\\bin\\Debug\\netcoreapp3.1-Text--{DateTime.Now.ToString("dd--MM--yyyy")}.txt"))
                foreach(Person item in personList)
                {
                    tw.WriteLine($"Name : {item.Name}");
                    tw.WriteLine($"Age : {item.Age}");
                    Console.WriteLine("Export olundu");
                    Console.WriteLine("---");

                }

            }
            else if(export == "2")
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.AddWorksheet("text");
                    worksheet.Cell("A1").Value = "Name";
                    worksheet.Cell("B1").Value = "Age";

                    foreach (Person item in personList)
                    {
                        worksheet.Cell("A1").Value = item.Name;
                        worksheet.Cell("B1").Value = item.Age;
                        workbook.SaveAs("export.xlsx");
                    }

                }
                
            }
        }
            public static void ShowMenu()
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("----");
                Console.WriteLine("Please make your choice");
                Console.WriteLine("1. Add et");
                Console.WriteLine("2. Search et");
                Console.WriteLine("3. Save et");
                Console.WriteLine("4. Load et");
                Console.WriteLine("5. Export et");
                Console.WriteLine("6. Display et");
                Console.WriteLine("7. Exit et");

                int choice = Convert.ToInt32(Console.ReadLine());
                switch (choice)
                {
                    case 1:
                        AddPerson();
                        ShowMenu();
                        break;
                    case 2:
                        SearchPerson();
                        ShowMenu();
                        break;
                    case 3:
                        SavedData();
                        ShowMenu();
                        break;
                    case 4:
                        LoadData();
                        ShowMenu();
                        break;
                    case 5:
                        ExportData();
                        ShowMenu();
                        break;
                    case 6:
                        DisplayData();
                        ShowMenu();
                        break;
                   
                    case 7:
                        break;
                    default:
                        Console.WriteLine("There is no such operation");
                        break;

                }
            }
            public static void DisplayData()
            {
                foreach (var item in personList)
                {
                    Console.WriteLine($"Name is : {item.Name}");
                    Console.WriteLine($"Age is : {item.Age}");
                }
            }
        




    }   
   
    
}

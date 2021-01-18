using System;
using System.Collections.Generic;
using System.Text;

namespace Project___Console_Application
{
    [Serializable]
    public class Person
    {
        public string Name { get; set; }
       public int Age { get; set; }
        public bool personId { get; internal set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }
}

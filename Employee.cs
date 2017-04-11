using System;
using System.Collections.Generic;

namespace ExcelByOpenXML
{
    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime DOB { get; set; }
        public decimal Salary { get; set; }
    }

    public sealed class Employees
    {
        static List<Employee> _employees;
        const int COUNT = 15;

        public static List<Employee> EmployeesList
        {
            private set { }
            get
            {
                return _employees;
            }
        }

        static Employees()
        {
            Initialize();
        }

        private static void Initialize()
        {
            _employees = new List<Employee>();

            Random random = new Random();

            for (int i = 0; i < COUNT; i++)
            {
                _employees.Add(new Employee()
                {
                    Id = i,
                    Name = "Employee " + i,
                    DOB = new DateTime(1999, 1, 1).AddMonths(i),
                    Salary = random.Next(100, 10000)
                });
            }
        }
    }
}
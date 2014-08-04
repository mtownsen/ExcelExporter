using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExporterTester
{
    public class ExcelExportTestData
    {
        Random random = new Random();

        public class TestData
        {
            public int CustomerID { get; set; }
            public string Name { get; set; }
            public string Address { get; set; }
            public string phoneNumber { get; set; }
            public int Age { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string Country { get; set; }
        }

        public List<TestData> GenerateTestData()
        {
            List<TestData> testData = new List<TestData>();

            for (int i = 0; i < 100; i++)
            {
                testData.Add(RandomPersonGenerator(i));
            }

            return testData;
        }

        private TestData RandomPersonGenerator(int customerID)
        {
            string[] randomFirstNames = { "Paul", "Michael", "Joseph", "Sunil", "Matt", "Heather", "Gabrielle", "Martha", "Christi", "Susan" };
            string[] randomLastNames = { "Townsen", "Smith", "Johnson", "Miller", "Williams", "Jones", "Thompson", "Davis", "Moore", "White" };
            string[] randomAddress = { "22 West 21st Avenue", "6768 North 12th Street", "123423 East 44th Block", "134423 South 23rd St", "100923 West 1st", "99 North 15th Junction apt# 1" };
            string[] randomCity = { "Blanchard", "Boulder", "Paris", "Lafayette", "Grand Station", "Moore", "Norman", "Cupertino" };
            string[] randomStates = { "CA", "FL", "OK", "RI", "WA", "CO", "WY", "TX", "NY", "DE" };
            string[] randomCountry = { "United States", "Mexico", "Brazil", "Canada", "Argentina" };

            TestData testData = new TestData
            {
                CustomerID = customerID,
                Name = string.Format("{0} {1}", randomFirstNames[random.Next(randomFirstNames.Length)], randomLastNames[random.Next(randomLastNames.Length)]),
                Age = random.Next(75),
                Address = randomAddress[random.Next(randomAddress.Length)],
                City = randomCity[random.Next(randomCity.Length)],
                State = randomStates[random.Next(randomStates.Length)],
                Country = randomCountry[random.Next(randomCountry.Length)]
            };

            return testData;
        }
    }
}

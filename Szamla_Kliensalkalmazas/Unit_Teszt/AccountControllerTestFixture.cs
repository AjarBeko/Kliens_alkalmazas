using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Unit_Teszt.Controller;

namespace Unit_Teszt
{
    public class AccountControllerTestFixture
    {
        [
            Test,
            TestCase("admín", false),
            TestCase("1234", false),
            TestCase("Admin", false), 
            TestCase("ABCD1234", false),
            TestCase("admin", true)
        ]
            public void TestValidateUserName(string username, bool expectedResult)
            {
                // Arrange
                var accountController = new AccountController();

                // Act
                var actualResult = accountController.ValidateUserName(username);

                // Assert
                Assert.AreEqual(expectedResult, actualResult);
            }
        [
            Test,
            TestCase("krumplifozelek2025", false),
            TestCase("KrumpliFozelek", false),
            TestCase("KrumpliFozelek2024", false),
            TestCase("KRUMPLIFOZELEK2025", false),
            TestCase("", false),
            TestCase(null, false),
            TestCase(" KrumpliFozelek2025 ", false),
            TestCase("KrumpliFozelek2025!", false),
            TestCase("Krumpli Fozelek2025", false),
            TestCase("KrumpliFozelek20252025", false),
            TestCase("KrumpliFozelek2025abc", false),
            TestCase("KrumpliFozelek2025", true)
        ]
        public void TestValidatePassword(string password, bool expectedResult)
        {
            // Arrange
            var accountController = new AccountController();

            // Act
            var actualResult = accountController.ValidatePassword(password);

            // Assert
            Assert.AreEqual(expectedResult, actualResult);
        }

    }
}

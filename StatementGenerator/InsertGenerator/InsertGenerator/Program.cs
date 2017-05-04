using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace InsertGenerator
{
    // Simple .NET core app, which generates some sql statements (or any other) based on template
    // TODO make it configurable!
    class Program
    {
        private static string insertCustomerTemplate = "insert into  Customer " +
                                  "(Id, CustomerGuid, Username, Email, EmailToRevalidate, AdminComment, IsTaxExempt, AffiliateId, VendorId, HasShoppingCartItems, RequireReLogin, FailedLoginAttempts, CannotLoginUntilDateUtc, Active, Deleted, IsSystemAccount, SystemName, LastIpAddress, CreatedOnUtc, LastLoginDateUtc, LastActivityDateUtc, RegisteredInStoreId, BillingAddress_Id, ShippingAddress_Id) " +
                                  "VALUES ({0}, '{1}', '{2}@yourStore.com', '{2}@yourStore.com', NULL, NULL, 0, 0, 0, 0, 0, 0, NULL, 1, 0, 0, NULL, '127.0.0.1', CURRENT_TIMESTAMP(), NULL, CURRENT_TIMESTAMP(), 1, 1, 1); ";

        private static string insertAddressTemplate =
                "insert into  Address (Id, FirstName, LastName, Email, Company, CountryId, StateProvinceId, City, Address1, Address2, ZipPostalCode, PhoneNumber, FaxNumber, CustomAttributes, CreatedOnUtc) " +
                "VALUES ({0}, '{1}', 'Smith', '{1}@yourStore.com', 'Nop Solutions Ltd', 1, 40, 'New York', '21 West 52nd Street', '', '10021', '12345678', '', NULL, CURRENT_TIMESTAMP());";

        static void Main(string[] args)
        {
            Func<int, string> nameFunc = i => $"Test{i}";
            Func<int, string> idFunc = i => i.ToString();

            List<string> statements = new List<string>();

            statements.Add(GenerateStatements(insertAddressTemplate,
                new[]{
                    idFunc,
                    nameFunc
                },
                21,
                50));

            statements.Add(GenerateStatements(insertCustomerTemplate,
                new[]
                {
                    idFunc,
                    i => Guid.NewGuid().ToString(),
                    nameFunc
                },
                9,
                50));

            File.WriteAllLines("output.txt", statements);

        }


        static string GenerateStatements(string template, Func<int, string>[] argumentFunc, int from, int to)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = from; i <= to; i++)
            {
                int index = i;
                var arguments = argumentFunc.Select(x => x(index));

                string statement = string.Format(template, arguments.ToArray());
                sb.AppendLine(statement);
            }
            return sb.ToString();
        }
    }


}
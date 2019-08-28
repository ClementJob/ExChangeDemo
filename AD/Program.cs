using System;

namespace AD
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            #region exchange 测试

            ExchangeMange manage = new ExchangeMange();
            Console.WriteLine("Start to enable user maibox...");
            try
            {
                manage.AddExchangeUser("Employee011@TestWin10.local", "Employee011");
            }
            catch (System.Management.Automation.RuntimeException ex)
            {
                Console.WriteLine("enable user maibox error...");
                Console.WriteLine(ex);
                Console.ReadLine();
            }
            Console.WriteLine("Finish to enable user maibox...");
            Console.ReadLine();



            #endregion exchange 测试

            #region ad 测试

            var adHelp = new ADHelp("10.10.10.100", "administrator", "Quattro@DH204", "IT", "TestWin10.local");

            #region 新建用户测试

            Console.WriteLine("Create User Start...");
            try
            {
                adHelp.CreateUser("Employee John", "Employee011");
                Console.WriteLine("Create User Finish...");
                Console.ReadLine();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException ex)
            {
                Console.WriteLine("Create User Error...");
                Console.WriteLine(ex);
                Console.ReadLine();
            }

            #endregion 新建用户测试

            #region 添加用户到租

            Console.WriteLine("Add user to group Start...");
            try
            {
                adHelp.AddUserToGroup(adHelp.GetDirectoryEntry(), "Employee012", "NewGroup01");
                Console.WriteLine("Add user to group Finish...");
                Console.ReadLine();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException ex)
            {
                Console.WriteLine("Add user to group Error...");
                Console.WriteLine(ex);
                Console.ReadLine();
            }

            #endregion 添加用户到租

            #region 新建ou和groupdemo

            //Test create ou
            Console.WriteLine("Create OU Start...");
            try
            {
                adHelp.CreateOU("NewOU01");
                Console.WriteLine("Create OU Finish...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Create OU Error...");
                Console.WriteLine(ex);
                Console.ReadLine();
            }

            //Test create group
            Console.WriteLine("Create Group Start...");
            try
            {
                adHelp.CreateGroup("NewGroup01");
                Console.WriteLine("Create Group Finish...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Create Group Error...");
                Console.WriteLine(ex);
                Console.ReadLine();
            }

            #endregion 新建ou和groupdemo

            #endregion ad 测试
        }
    }
}
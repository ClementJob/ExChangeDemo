using System;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;

namespace AD
{
    public class ExchangeMange
    {
        public void AddExchangeUser(string identity, string alias)
        {
            string runasUsername = "administrator";
            string runasPassword = "Quattro@DH204";
            SecureString ssRunasPassword = new SecureString();
            foreach (char x in runasPassword)
            {
                ssRunasPassword.AppendChar(x);
            }
            PSCredential credentials =
                new PSCredential(runasUsername, ssRunasPassword);
            var connInfo = new WSManConnectionInfo(new Uri("http://10.10.10.100/PowerShell"),
                "http://schemas.microsoft.com/powershell/Microsoft.Exchange",
                credentials);
            connInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            var runspace = RunspaceFactory.CreateRunspace(connInfo);
            var command = new Command("Enable-Mailbox");
            command.Parameters.Add("Identity", identity);
            command.Parameters.Add("Alias", alias);
            runspace.Open();
            var pipeline = runspace.CreatePipeline();
            pipeline.Commands.Add(command);
            var results = pipeline.Invoke();
            Console.WriteLine("通道错误数：" + pipeline.Error.Count);
            runspace.Dispose();
        }
    }
}
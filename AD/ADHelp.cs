using System;
using System.DirectoryServices;

namespace AD
{
    public class ADHelp
    {
        #region 构造函数

        private string _pathFront { get; set; } = "";
        private string _pathBehind { get; set; } = "";
        private string _path { get; set; }
        private string _userName { get; set; }
        private string _pwd { get; set; }

        private string _oU { get; set; }

        private string[] _dC { get; set; }

        public ADHelp(string serverIP, string userName, string pwd, string ou, string dc)
        {
            _userName = userName;
            _pwd = pwd;
            _oU = ou;
            _dC = dc.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            _pathFront = $"LDAP://{serverIP}";

            if (!string.IsNullOrEmpty(ou))
                _pathBehind += $"OU={ou}";
            if (_dC.Length == 2)
            {
                if (_pathBehind.Contains("OU"))
                    _pathBehind += ",";
                _pathBehind += $"DC={_dC[0]},DC={_dC[1]}";
            }
            if (!string.IsNullOrEmpty(_pathBehind))
                _path = _pathFront + "/" + _pathBehind;
            if (_path.Substring(_path.Length - 1, 1) == "/")
                _path = _path.Substring(0, _path.Length - 1);
        }

        /// <summary>
        /// 判断是否存在
        /// </summary>
        /// <param name="objectName"></param>
        /// <param name="catalog"></param>
        /// <returns></returns>
        public bool ObjectExists(string objectName, string catalog)
        {
            DirectoryEntry de = GetDirectoryEntry();
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;
            switch (catalog)
            {
                case "User": deSearch.Filter = "(&(objectClass=user) (cn=" + objectName + "))"; break;
                case "Group": deSearch.Filter = "(&(objectClass=group) (cn=" + objectName + "))"; break;
                case "OU": deSearch.Filter = "(&(objectClass=OrganizationalUnit) (OU=" + objectName + "))"; break;
                default: break;
            }
            SearchResultCollection results = deSearch.FindAll();
            if (results.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        #endregion 构造函数

        #region 连接

        /// <summary>
        /// 创建连接
        /// 其中ADPath是要查询组织单元的所在的LDAP，其格式为：LDAP:\\ OU=XX部门，OU=XX公司，DC=域名，DC=COM，
        /// 如果连接到的AD是在服务器上那么格式写成LDAP:\\XX.XX.XX.XX\ OU=XX部门，OU=XX公司，DC=域名，DC=COM（XX.XX.XX.XX为服务器IP）；
        /// ADAccount和ADPwd为AD用户的账户和密码，如果是管理员则可以进行任何操作，普通只能进行查询操作。
        /// </summary>
        /// <returns></returns>
        public DirectoryEntry GetDirectoryEntry()
        {
            DirectoryEntry de = new DirectoryEntry();
            de.Path = _path;
            de.Username = _userName;
            de.Password = _pwd;
            return de;
        }

        /// <summary>
        /// 带有一个参数的创建连接重载
        /// </summary>
        /// <param name="DomainReference"></param>
        /// <returns></returns>
        public DirectoryEntry GetDirectoryEntry(string DomainReference)
        {
            DirectoryEntry entry = new DirectoryEntry(DomainReference, _userName, _pwd, AuthenticationTypes.Secure);
            return entry;
        }

        #endregion 连接

        #region 组织

        /// <summary>
        /// 新建OU
        /// </summary>
        /// <param name="path"></param>
        public void CreateOU(string name)
        {
            if (!ObjectExists(name, "OU"))
            {
                DirectoryEntry dse = GetDirectoryEntry();
                DirectoryEntries ous = dse.Children;
                DirectoryEntry newou = ous.Add("OU=" + name, "OrganizationalUnit");
                newou.CommitChanges();
                newou.Close();
                dse.Close();
            }
            else
            {
                Console.WriteLine("OU已存在");
            }
        }

        /// <summary>
        /// 新建Security Group
        /// </summary>
        /// <param name="path"></param>
        public void CreateGroup(string name)
        {
            if (!ObjectExists(name, "Group"))
            {
                DirectoryEntry dse = GetDirectoryEntry();
                DirectoryEntries Groups = dse.Children;
                DirectoryEntry newgroup = Groups.Add("CN=" + name, "group");
                newgroup.CommitChanges();
                newgroup.Close();
                dse.Close();
            }
            else
            {
                Console.WriteLine("用户组已存在");
            }
        }

        #endregion 组织

        #region 创建用户

        /// <summary>
        /// 新建用户
        /// </summary>
        /// <param name="name"></param>
        /// <param name="login"></param>
        public void CreateUser(string name, string login)
        {
            if (ObjectExists(login, "User"))
            {
                Console.WriteLine("用户已存在");
                Console.ReadLine();
                return;
            }
            DirectoryEntry de = GetDirectoryEntry();
            DirectoryEntries users = de.Children;
            DirectoryEntry newuser = users.Add("CN=" + login, "user");
            SetProperty(newuser, "givenname", name);
            SetProperty(newuser, "SAMAccountName", login);
            SetProperty(newuser, "userPrincipalName", login + string.Join(".", _dC));
            newuser.CommitChanges();

            //SetPassword(newuser.Path);
            //newuser.CommitChanges();
            EnableAccount(newuser);
            newuser.Close();
            de.Close();
        }

        /// <summary>
        /// 属性设置
        /// </summary>
        /// <param name="de"></param>
        /// <param name="PropertyName"></param>
        /// <param name="PropertyValue"></param>
        public void SetProperty(DirectoryEntry de, string PropertyName, string PropertyValue)
        {
            if (PropertyValue != null)
            {
                if (de.Properties.Contains(PropertyName))
                {
                    de.Properties[PropertyName][0] = PropertyValue;
                }
                else
                {
                    de.Properties[PropertyName].Add(PropertyValue);
                }
            }
        }

        /// <summary>
        /// 密码设置
        /// </summary>
        /// <param name="path"></param>
        public void SetPassword(string path)
        {
            DirectoryEntry user = new DirectoryEntry();
            user.Path = path;
            user.AuthenticationType = AuthenticationTypes.Secure;
            object ret = user.Invoke("SetPassword", new object[] { "Password01!" });
            user.CommitChanges();
            user.Close();
        }

        #endregion 创建用户

        #region 用户组

        /// <summary>
        /// 添加用户到组 已存在 则移除
        /// </summary>
        /// <param name="de"></param>
        /// <param name="userDn"></param>
        /// <param name="GroupName"></param>
        public void AddUserToGroup(DirectoryEntry de, string userName, string GroupName)
        {
            var userDn = $"CN={userName}," + _pathBehind;
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;
            deSearch.Filter = "(&(objectClass=group) (cn=" + GroupName + "))";
            SearchResult Groupresult = deSearch.FindOne();
            if (Groupresult != null)
            {
                DirectoryEntry user = GetDirectoryEntry(_pathFront + $"/" + userDn);
                if (user != null)
                {
                    DirectoryEntry dirEntry = Groupresult.GetDirectoryEntry();
                    if (dirEntry.Properties["member"].Contains(userDn))
                    {
                        Console.WriteLine("用户组中已存在该用户，即将移除");
                        dirEntry.Properties["member"].Remove(userDn);
                        Console.WriteLine("用户已从组中移除");
                    }
                    else
                    {
                        dirEntry.Properties["member"].Add(userDn);
                        Console.WriteLine("添加成功，用户已添加到组");
                    }
                    dirEntry.CommitChanges();
                    dirEntry.Close();
                }
                else
                {
                    Console.WriteLine("用户不存在");
                }
                user.Close();
            }
            else
            {
                Console.WriteLine("用户组不存在");
            }
            return;
        }

        #endregion 用户组

        #region 用户信息更新

        /// <summary>
        /// 用户信息更新
        /// </summary>
        /// <param name="de"></param>
        /// <param name="UserName"></param>
        /// <param name="company"></param>
        public void ModifyUser(DirectoryEntry de, string UserName, string company)
        {
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;
            deSearch.Filter = "(&(objectClass=user) (cn=" + UserName + "))";
            SearchResult result = deSearch.FindOne();
            if (result != null)
            {
                DirectoryEntry dey = GetDirectoryEntry(result.Path);
                SetProperty(dey, "company", company);
                dey.CommitChanges();
                dey.Close();
            }
            de.Close();
        }

        #endregion 用户信息更新

        #region Enable/Disable用户账号

        /// <summary>
        /// 启用账号
        /// </summary>
        /// <param name="de"></param>
        public void EnableAccount(DirectoryEntry de)
        {
            //设置账号密码不过期
            int exp = (int)de.Properties["userAccountControl"].Value;
            de.Properties["userAccountControl"].Value = exp | 0x10000;
            de.CommitChanges();
            //启用账号
            int val = (int)de.Properties["userAccountControl"].Value;
            de.Properties["userAccountControl"].Value = val & ~0x0002;
            de.CommitChanges();
        }

        /// <summary>
        /// 停用账号
        /// </summary>
        /// <param name="de"></param>
        public void DisableAccount(DirectoryEntry de)
        {
            //启用账号
            int val = (int)de.Properties["userAccountControl"].Value;
            de.Properties["userAccountControl"].Value = val | 0x0002;
            de.CommitChanges();
        }

        #endregion Enable/Disable用户账号
    }
}
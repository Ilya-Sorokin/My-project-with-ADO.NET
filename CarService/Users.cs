using System;
using System.Collections.Generic;

namespace JobCentre
{
    [Serializable]
    public class Users
    {
        public List<string> Logins = new List<string>(); // Логин.
        public List<string> Passwords = new List<string>(); // Пароль.
    }
}

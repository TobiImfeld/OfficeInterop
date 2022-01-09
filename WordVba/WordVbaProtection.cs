using System;
using System.Security.Cryptography;
using System.Text;

namespace WordVba
{
    public class WordVbaProtection
    {
        private WordVbaProject project;

        internal WordVbaProtection(WordVbaProject project)
        {
            this.project = project;
            VisibilityState = true;
        }

        public bool UserProtected { get; internal set; }

        public bool HostProtected { get; internal set; }

        public bool VbeProtected { get; internal set; }

        public bool VisibilityState { get; internal set; }

        internal byte[] PasswordHash { get; set; }

        internal byte[] PasswordKey { get; set; }

        /// <summary>
        /// Password protect the VBA project.
        /// An empty string or null will remove the password protection
        /// </summary>
        /// <param name="Password">The password</param>
        public void SetPassword(string Password)
        {
            if (string.IsNullOrEmpty(Password))
            {
                PasswordHash = null;
                PasswordKey = null;
                VbeProtected = false;
                HostProtected = false;
                UserProtected = false;
                VisibilityState = true;
                this.project.ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
            }
            else
            {
                //Join Password and Key
                byte[] data;
                //Set the key
                PasswordKey = new byte[4];
                RandomNumberGenerator r = RandomNumberGenerator.Create();
                r.GetBytes(PasswordKey);

                data = new byte[Password.Length + 4];
                Array.Copy(Encoding.GetEncoding(this.project.CodePage).GetBytes(Password), data, Password.Length);
                VbeProtected = true;
                VisibilityState = false;
                Array.Copy(PasswordKey, 0, data, data.Length - 4, 4);

                //Calculate Hash
                var provider = SHA1.Create();
                PasswordHash = provider.ComputeHash(data);
                this.project.ProjectID = "{00000000-0000-0000-0000-000000000000}";
            }
        }      
    }
}

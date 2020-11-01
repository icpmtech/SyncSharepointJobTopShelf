namespace SyncUserProfilesToListContatos
{
    public class Configuration
    {
        public string ListTitle { get; }
        public string Password { get; }
        public string EmailAccount { get; }
        public string TenantUrl { get; }

        public Configuration(string listTitle, string password, string emailAccount, string tenantUrl)
        {
            this.ListTitle = listTitle;
            this.Password = password;
            this.EmailAccount = emailAccount;
            this.TenantUrl = tenantUrl;
        }
    }
}
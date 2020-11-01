using Quartz;

namespace SyncUserProfilesToListContatos
{
    /// <summary>
    /// This class is used for Get App Configuration from App.Config
    /// </summary>
    public interface IConfigurationProvider
    {
        Configuration GetConfiguration();
    }

    public class ConfigurationProvider : IConfigurationProvider
    {

        public Configuration Configuration { get; }
        public ConfigurationProvider(Configuration configuration )
        {
            Configuration = configuration;
        }

        public Configuration GetConfiguration()
        {
            return Configuration;
        }
    }
}
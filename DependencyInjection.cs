using System;
using System.Collections.Specialized;
using System.Configuration;
using Autofac;
using Autofac.Extras.Quartz;
using log4net;

namespace SyncUserProfilesToListContatos
{
    /// <summary>
    /// Bootstraps the Dependency Injection using Autofac
    /// </summary>
    public class DependencyInjection
    {
        public static IContainer Build()
        {
            var builder = new ContainerBuilder();

            // Registers SchedulerSharePointService with Autofac
            builder.RegisterType<SchedulerSharePointService>().AsSelf().InstancePerLifetimeScope();                       

            // Register Quartz with Autofac and reads quartz section in App.config. 
            builder.RegisterModule(new QuartzAutofacFactoryModule
            {
                ConfigurationProvider = context => (NameValueCollection)ConfigurationManager.GetSection("quartz")
            });
            builder.Register(ctx =>
            {
                var listTitle = ConfigurationManager.AppSettings["ListTitle"];
                var password = ConfigurationManager.AppSettings["PASSWORD"];
                var emailAccount = ConfigurationManager.AppSettings["EMAILACCOUNT"];
                var tenantUrl = ConfigurationManager.AppSettings["TenantUrl"];
                Configuration configuration = new Configuration(listTitle, password, emailAccount, tenantUrl);
                return new ConfigurationProvider(configuration);
            }).As<IConfigurationProvider>();
            // This line registers all Jobs in the current executing assembly
            builder.RegisterModule(new QuartzAutofacJobsModule(typeof(SharepointSyncJob).Assembly));

            // As per recent update GetLogger might have some performance issues. 
            builder.Register(c => LogManager.GetLogger(typeof(Object))).As<ILog>();            

            return builder.Build();
        }
    }
}
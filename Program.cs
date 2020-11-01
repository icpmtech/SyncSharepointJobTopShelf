using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;
using Topshelf.Autofac;
using Autofac;

namespace SyncUserProfilesToListContatos
{
    class Program
    {
        static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();

            IContainer container = DependencyInjection.Build();


            HostFactory.Run(hostConfigurator => 
            {
                // Set windows service properties
                hostConfigurator.SetServiceName("SyncSharepointList");
                hostConfigurator.SetDisplayName("Sync Sharepoint List Contacts");
                hostConfigurator.SetDescription("Job to sync Sharepoint the list Contacts in Sharepoint.");

                hostConfigurator.RunAsLocalSystem();
                // Configure Log4Net with Topself
                hostConfigurator.UseLog4Net();
                hostConfigurator.UseAutofacContainer(container);
                hostConfigurator.Service<SchedulerSharePointService>(serviceConfigurator => 
                {
                    serviceConfigurator.ConstructUsing(hostSettings => container.Resolve<SchedulerSharePointService>());
                    serviceConfigurator.WhenStarted(s => s.Start());
                    serviceConfigurator.WhenStopped(s => s.Shutdown());
                });
            });            
        }
    }
}

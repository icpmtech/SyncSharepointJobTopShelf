using Quartz;

namespace SyncUserProfilesToListContatos
{
    /// <summary>
    /// This class is used for starting Quartz job
    /// </summary>
    public class SchedulerSharePointService : ISchedulerSharePointService
    {
        private readonly IScheduler _scheduler;

        /// <summary>
        /// Injects IScheduler from Quartz library
        /// </summary>
        /// <param name="scheduler"></param>
        public SchedulerSharePointService(IScheduler scheduler)
        {
            _scheduler = scheduler;
        }

        public void Start()
        {
            _scheduler.Start();
        }
        
        public void Shutdown()
        {
            _scheduler.Shutdown();
        }
    }
}
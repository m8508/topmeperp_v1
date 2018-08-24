using Hangfire;
using Hangfire.Common;
using Hangfire.Logging;
using Hangfire.States;
using Hangfire.Storage;
using System;

namespace topmeperp.Schedule
{
    /// <summary>
    /// 排程作業，啟動方法!!
    /// </summary>
    public class BackgroundService
    {
        static log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        [AutomaticRetry(Attempts = 20)]
        public  void SendMailSchedule()
        {
            logger.Info("SendMailSchedule start !!" + DateTime.Now);
        }
    }
    //HangFire Task Failure Event Sample
    public class LogFailureAttribute : JobFilterAttribute, IApplyStateFilter
    {
        private static readonly ILog Logger = LogProvider.GetCurrentClassLogger();

        public void OnStateApplied(ApplyStateContext context, IWriteOnlyTransaction transaction)
        {
            var failedState = context.NewState as FailedState;
            if (failedState != null)
            {
                Logger.ErrorException(String.Format("Background job #{0} was failed with an exception.", context.JobId),failedState.Exception);
            }
        }

        public void OnStateUnapplied(ApplyStateContext context, IWriteOnlyTransaction transaction)
        {
        }
    }
}
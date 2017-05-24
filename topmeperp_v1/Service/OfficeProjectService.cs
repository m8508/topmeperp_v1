using System.Collections;
using java.util;
using net.sf.mpxj;
using System;
using System.Collections.Generic;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class OfficeProjectService
    {
        List<PLAN_TASK> allTasks = new List<PLAN_TASK>();

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(typeof(OfficeProjectService));
        public void convertProjectFile(string prjfile)
        {
            net.sf.mpxj.mpp.MPPReader reader = new net.sf.mpxj.mpp.MPPReader();
            ProjectFile projectObj = reader.read(prjfile);

            int i = 1;
            foreach (net.sf.mpxj.Task task in ToEnumerable(projectObj.AllTasks))
            {
                DateTime dtStart = new DateTime();
                DateTime dtFinish = new DateTime();
                //ToString("yyyyMMddHHmmss")
                if (null != task.Start)
                {
                    dtStart = new DateTime((task.Start.getYear() + 1900), task.Start.getMonth() + 1, task.Start.getDate());
                    logger.Debug("start date Year =" + (task.Start.getYear() + 1900) + ",Month=" + (task.Start.getMonth() + 1) + ",Date=" + task.Start.getDate());
                    dtFinish = new DateTime((task.Finish.getYear() + 1900), task.Finish.getMonth() + 1, task.Finish.getDate());
                    logger.Debug("start date Year =" + (task.Finish.getYear() + 1900) + ",Month=" + (task.Finish.getMonth() + 1) + ",Date=" + task.Finish.getDate());
                }
                logger.Debug("DURATION=" + task.Duration + ",Task: " + i + "=" + task.Name + ",StartDate=" + dtStart.ToString("yyyy/MM/dd") + ",EndDate=" + dtFinish.ToString("yyyy/MM/dd") + " ID=" + task.ID + " Unique ID=" + task.UniqueID);

                i++;
                foreach (net.sf.mpxj.Task child in ToEnumerable(task.ChildTasks))
                {
                    Console.WriteLine(child.ParentTask.Name + ",Task: " + child.Name);
                    logger.Debug(child.ParentTask.Name + ",Task: " + child.Name);
                }
            }
        }
        private static OfficeProjectService ToEnumerable(Collection javaCollection)
        {
            return new OfficeProjectService(javaCollection);
        }
        public OfficeProjectService(Collection collection)
        {
            m_collection = collection;
        }

        public IEnumerator GetEnumerator()
        {
            return new Enumerator(m_collection);
        }

        private Collection m_collection;
    }

    public class Enumerator : IEnumerator
    {
        public Enumerator(Collection collection)
        {
            m_collection = collection;
            m_iterator = m_collection.iterator();
        }

        public object Current
        {
            get
            {
                return m_iterator.next();
            }
        }

        public bool MoveNext()
        {
            return m_iterator.hasNext();
        }

        public void Reset()
        {
            m_iterator = m_collection.iterator();
        }

        private Collection m_collection;
        private Iterator m_iterator;
    }
}
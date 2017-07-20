using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web;

namespace topmeperp.Service
{
    public class EMailService
    {
        ILog log = log4net.LogManager.GetLogger(typeof(EMailService));
        public bool SendMailByGmail(string fromAddress, string MailList, string bccMailList, string Subject, string Body)
        {
            MailMessage msg = new MailMessage();
            //收件者，以逗號分隔不同收件者 ex "test@gmail.com,test2@gmail.com"
            if (null != MailList && "" != MailList)
            {
                msg.To.Add(MailList);
            }
            //密件副本清單
            if (null != bccMailList && "" != bccMailList)
            {
                msg.Bcc.Add(bccMailList);
            }

            msg.From = new MailAddress(fromAddress, "測試郵件", System.Text.Encoding.UTF8);
            //郵件標題 
            msg.Subject = Subject;
            //郵件標題編碼  
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            //郵件內容
            msg.Body = Body;
            msg.IsBodyHtml = true;
            msg.BodyEncoding = System.Text.Encoding.UTF8;//郵件內容編碼 
            msg.Priority = MailPriority.Normal;//郵件優先級 
                                               //建立 SmtpClient 物件 並設定 Gmail的smtp主機及Port 
            #region 其它 Host
            /*
             *  outlook.com smtp.live.com port:25
             *  yahoo smtp.mail.yahoo.com.tw port:465
            */
            #endregion
            SmtpClient MySmtp = new SmtpClient("smtp.gmail.com", 587);
            //設定你的帳號密碼
            MySmtp.Credentials = new System.Net.NetworkCredential("chuph.chuph@gmail.com", "chuph1234");
            //Gmial 的 smtp 使用 SSL
            MySmtp.EnableSsl = true;
            log.Info("Email Send!!");
            try
            {
                MySmtp.Send(msg);
                return true;
            }
            catch (Exception ex)
            {
                log.Error(ex.StackTrace);
                return false;
            }
        }
    }
}
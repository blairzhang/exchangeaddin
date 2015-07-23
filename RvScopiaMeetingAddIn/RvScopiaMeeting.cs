using System;
using System.Linq;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Mime;
using Radvision.Scopia.ExchangeMeetingAddIn.icm;
using System.Net;
using System.Security.AccessControl;
using Microsoft.Exchange.Data.ContentTypes.Tnef;
using Microsoft.Exchange.Data.TextConverters;
using webServiceData = Microsoft.Exchange.WebServices.Data;
using System.Threading;
using RvScopiaMeetingAddIn;
using System.Security.Cryptography;


namespace Radvision.Scopia.ExchangeMeetingAddIn
{
   

    public class RvScopiaMeetingFactory : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new RvScopiaMeeting();
        }
    }

    class TraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
    {
        #region ITraceListener Members
        public void Trace(string traceType, string traceMessage)
        {

            RvLogger.DebugWrite("Trace: " + traceMessage.ToString());
        }
        #endregion

       }


    public enum RequestType { 
        CreateMeeting = 0,
        ModifyMeeting = 1,
        CancelMeeting = 2,
        Other = 3
    }

    public enum MeetingType
    {
        Normal = 0,
        Reccurence = 1,
        Ocurrence = 2
    }

    public class RvScopiaMeeting : RoutingAgent
    {
        private static Dictionary<string, string> settingsInfo = null;

        private static Dictionary<string, string> messages = null;

        private Adapter adapter = null;

        private EmailMessage emailMessage = null;

        public static string CONFERENCE_ID = "CONFERENCE_ID";

        public static string HASH_MARK = "HASH:";

        //public static string ERROR_MARK = "Error:";

        //public static string WARNING_MARK = "Warning:";

        private static string SERVER_ACCOUNT = "sa@zilladom12.extest.microsoft.com";

        public static string FAIL_TO_CREATE = "Fail to create meeting : ";

        public static string FAIL_TO_UPDATE = "Fail to update meeting : ";

        public static string FAIL_TO_DELETE = "Fail to delete meeting : ";

        public static List<string> MARK_OF_SCOPIA_LOCATION = new List<string>();
        public static List<string> MARK_OF_SCOPIA_SUBJECT = new List<string>();
        public static List<string> MARK_OF_SCOPIA_RECEIPENT = new List<string>();
        public static bool IS_DELETE_ERROR = true;
        public static bool IS_SEND_SUCCESS_MAIL = false;
        public static bool IS_RESCHEDULE = true;
        public static int AFTER_HOWMANYSECONDS_TO_GETAPPOINTMENT = 25;
        public static int HOWMANYSECONDS_TO_HANDLE_TASKS = 5;
        public static int HOW_MANY_SECONDS_WAIT_FOR_FOLLOWING_REQUEST = 6;
        public static int HOW_MANY_MILLISECONDS_TO_CALL_EWS = 100;
        public static double LOG_FILE_SIEZ_MB = 10;
        public static double HOW_MANY_MINUTES_TO_CHECK_LOG = 100;

        public static bool USE_HTML = true;

        private static string errorMessage = "";
        private static LinkedList<object[]> ewsTasks = new LinkedList<object[]>();
        private static object serviceLock = new object();

        private AgentAsyncContext agentAsyncContext;

        static RvScopiaMeeting(){
            try
            {
                RvLogger.OpenLog();
                RvScopiaMeeting.parseSettingFile();
                string messagesPropertiesFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\messages.properties";
                messages = RvScopiaMeeting.getPropertiesFromPropertiesFile(messagesPropertiesFile);

                startThreadForEWSInvoking();
                System.Timers.Timer t = new System.Timers.Timer(HOW_MANY_MINUTES_TO_CHECK_LOG * 60 * 1000);  
                t.Elapsed += new System.Timers.ElapsedEventHandler(RvScopiaMeeting.checkLog); 
                t.AutoReset = true;
                t.Enabled = true;  
                RvScopiaMeeting.getWebService(null);
            } 
            catch (Exception ex)
            {
                errorMessage = ScopiaMeetingAddInException.ERROR_MESSAGE_GETUSER_ERROR + ex.Message;
                RvLogger.DebugWrite(ex.Message);
                RvLogger.DebugWrite(ex.StackTrace);
            }
        }

        private static void parseSettingFile() {
            string settingsFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\settings.properties";
            settingsInfo = RvScopiaMeeting.getPropertiesFromPropertiesFile(settingsFile);
            string markStr = "";
            markStr = settingsInfo["ScopiaMeeting.symbol.location"];
            string[] marks = null;
            if (!string.IsNullOrEmpty(markStr))
            {
                marks = markStr.Split(';');
                foreach (string mark in marks)
                    if (!string.IsNullOrEmpty(mark))
                        MARK_OF_SCOPIA_LOCATION.Add(mark.Trim().ToLower());
            }
            markStr = settingsInfo["ScopiaMeeting.symbol.subject"];
            if (!string.IsNullOrEmpty(markStr))
            {
                marks = markStr.Split(';');
                foreach (string mark in marks)
                    if (!string.IsNullOrEmpty(mark))
                        MARK_OF_SCOPIA_SUBJECT.Add(mark.Trim().ToLower());
            }
            markStr = settingsInfo["ScopiaMeeting.symbol.recipient"];
            if (!string.IsNullOrEmpty(markStr))
            {
                marks = markStr.Split(';');
                foreach (string mark in marks)
                    if (!string.IsNullOrEmpty(mark))
                        MARK_OF_SCOPIA_RECEIPENT.Add(mark.Trim().ToLower());
            }

            string deleteError = settingsInfo["scopia.clientless.scheduling.deleteAppointmentIncaseError"];
            IS_DELETE_ERROR = string.IsNullOrEmpty(deleteError) ? IS_DELETE_ERROR : ("false".Equals(deleteError.ToLower()) ? false : true);
            string sendSuccess = settingsInfo["scopia.clientless.scheduling.sendSuccessMail"];
            IS_SEND_SUCCESS_MAIL = string.IsNullOrEmpty(sendSuccess) ? IS_SEND_SUCCESS_MAIL : "true".Equals(sendSuccess.ToLower());
            string strict = settingsInfo["scopia.clientless.scheduling.strict.mode"];
            IS_RESCHEDULE = string.IsNullOrEmpty(strict) ? IS_RESCHEDULE : !"true".Equals(strict.ToLower());
            if (settingsInfo.ContainsKey("ExchangeServer.service.account"))
                SERVER_ACCOUNT = settingsInfo["ExchangeServer.service.account"];
            if(settingsInfo.ContainsKey("ExchangeServer.user.email"))
                SERVER_ACCOUNT = settingsInfo["ExchangeServer.user.email"];
            string useHtml = settingsInfo["scopia.clientless.scheduling.use.html"];
            USE_HTML = string.IsNullOrEmpty(useHtml) ? USE_HTML : "true".Equals(useHtml.ToLower());
            string afterHowManySecondsToGetAppointment = settingsInfo["scopia.clientless.scheduling.getAppointment.delaySeconds"];
            if (!string.IsNullOrEmpty(afterHowManySecondsToGetAppointment))
            {
                try
                {
                    int senconds = int.Parse(afterHowManySecondsToGetAppointment);
                    if (senconds > 0 && senconds < 200)
                        AFTER_HOWMANYSECONDS_TO_GETAPPOINTMENT = senconds;
                }
                catch (Exception) { }
            }
            string howManySecondsToHandleTasks = settingsInfo["scopia.clientless.scheduling.intervalOfHandlingTasks"];
            if (!string.IsNullOrEmpty(howManySecondsToHandleTasks))
            {
                try
                {
                    int senconds = int.Parse(howManySecondsToHandleTasks);
                    if (senconds > 0 && senconds < 100)
                        HOWMANYSECONDS_TO_HANDLE_TASKS = senconds;
                }
                catch (Exception) { }
            }

            string howManySecondsWait4FollowingRequest = settingsInfo["scopia.clientless.scheduling.intervalOfWaitingFurtherRequest"];
            if (!string.IsNullOrEmpty(howManySecondsWait4FollowingRequest))
            {
                try
                {
                    int senconds = int.Parse(howManySecondsWait4FollowingRequest);
                    if (senconds > 0 && senconds < 100)
                        HOW_MANY_SECONDS_WAIT_FOR_FOLLOWING_REQUEST = senconds;
                }
                catch (Exception) { }
            }

            string howManyMilliseconds2CallEws = settingsInfo["scopia.clientless.scheduling.intervalOfInvokingEWS"];
            if (!string.IsNullOrEmpty(howManyMilliseconds2CallEws))
            {
                try
                {
                    int senconds = int.Parse(howManyMilliseconds2CallEws);
                    if (senconds > 0 && senconds < 10000)
                        HOW_MANY_MILLISECONDS_TO_CALL_EWS = senconds;
                }
                catch (Exception) { }
            }

            string logFileSizeMB = settingsInfo["scopia.clientless.log.fileSize"];
            if (!string.IsNullOrEmpty(logFileSizeMB))
            {
                try
                {
                    double size = double.Parse(logFileSizeMB);
                    if (size > 0 && size < 1024)
                        LOG_FILE_SIEZ_MB = size;
                }
                catch (Exception) { }
            }

            string howManyMinutes2CheckLog = settingsInfo["scopia.clientless.log.invervalOfCheckingLogSize"];
            if (!string.IsNullOrEmpty(howManyMinutes2CheckLog))
            {
                try
                {
                    double minutes = double.Parse(howManyMinutes2CheckLog);
                    if (minutes > 0 && minutes < 60 * 24 * 10)
                        HOW_MANY_MINUTES_TO_CHECK_LOG = minutes;
                }
                catch (Exception) { }
            }
            
        }

        public RvScopiaMeeting()
        {
            //this.OnCategorizedMessage += new CategorizedMessageEventHandler(RvScopiaMeeting_OnCategorizedMessage);
            this.OnResolvedMessage += new ResolvedMessageEventHandler(RvScopiaMeeting_OnResolvedMessage);
        }

        private static string GetFQDN()
        {
            string domainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            string hostName = Dns.GetHostName();
            string fqdn = "";
            if (!hostName.Contains(domainName))
                fqdn = hostName + "." + domainName;
            else
                fqdn = hostName;

            return fqdn;
        }

        void RvScopiaMeeting_OnResolvedMessage(ResolvedMessageEventSource source, QueuedMessageEventArgs args){
            this.emailMessage = args.MailItem.Message;

            if (this.emailMessage == null || this.emailMessage.TnefPart == null) {
                return;
            }

            long now = DateTime.UtcNow.Ticks;
            SchedulingInfo schedulingInfo = new SchedulingInfo();
            schedulingInfo.subject = args.MailItem.Message.Subject;
            schedulingInfo.delegatorEmailAddr = args.MailItem.Message.From.NativeAddress;
            RvLogger.DebugWrite("Enter transport agent, from: " + schedulingInfo.delegatorEmailAddr + ", subject: " + schedulingInfo.subject);

            try
            {
                this.agentAsyncContext = this.GetAgentAsyncContext();
                schedulingInfo.requestType = this.getRequestType(this.emailMessage.MapiMessageClass);
                //Reject all meeting type except request and cancel.
                if (RequestType.Other == schedulingInfo.requestType)
                {
                    RvLogger.DebugWrite("Reject other request type: " + this.emailMessage.MapiMessageClass);
                    return;
                }
                RvMailParser parser = new RvMailParser(this);
                try
                {
                    parser.parseTnefSimple(args.MailItem, schedulingInfo);
                }
                catch (Exception exceptionParseMail)
                {
                    RvLogger.DebugWrite("Fail to parse mail.");
                    RvLogger.DebugWrite(exceptionParseMail.Message);
                    RvLogger.DebugWrite(exceptionParseMail.StackTrace);
                    return;
                }
                
                //Reject forwarded appointment
                if (!string.IsNullOrEmpty(schedulingInfo.subjectPrefix) && "FW:".Equals(schedulingInfo.subjectPrefix))
                {
                    RvLogger.DebugWrite("Reject forward request type");
                    return;
                }

                if (schedulingInfo.recurrencePattern != null){
                    schedulingInfo.recurrenceHashInfo = RvScopiaMeeting.getHashString4Str(schedulingInfo.recurrencePattern.getStringForHash());
                    schedulingInfo.recurrencePattern.startDate = schedulingInfo.startDate;
                    schedulingInfo.recurrencePattern.endDate = schedulingInfo.endDate;
                }

                if (null == schedulingInfo.emailMessage) {
                    RvLogger.DebugWrite("null == schedulingInfo.emailMessage================================================");
                    return; 
                }

                if (!isScopia(schedulingInfo))
                {
                    RvLogger.DebugWrite("This is not a SCOPIA meeting");
                    return;
                }

                parseRecipentsChanged(schedulingInfo);
                if (schedulingInfo.isRecipentsChanged)
                    if (schedulingInfo.requestType == RequestType.CancelMeeting)
                    {
                        schedulingInfo.requestType = RequestType.CreateMeeting;
                        schedulingInfo.isAddRecipents = false;
                        schedulingInfo.subject = schedulingInfo.subject.Substring(schedulingInfo.subjectPrefix.Length + 1);
                        Thread.Sleep(HOW_MANY_SECONDS_WAIT_FOR_FOLLOWING_REQUEST * 1000);
                    }else
                        schedulingInfo.isAddRecipents = true;

                if (RvScopiaMeeting.SERVER_ACCOUNT.Equals(schedulingInfo.senderEmailAddr))
                {
                    RvLogger.DebugWrite("Send a email back to notify the sender this mail is failed to send out.");
                    return;
                }

                //when modify a recurrence, to make sure the modified ocurrence request is later than the recurrence request.
                if (schedulingInfo.meetingType == MeetingType.Ocurrence)
                    Thread.Sleep(HOW_MANY_SECONDS_WAIT_FOR_FOLLOWING_REQUEST * 1000);

                icm.XmlApi.scheduleReportType result = changeMail(source, args, schedulingInfo, now);
                if (null != result && result.Success && isCreateMeetingRequest(schedulingInfo))
                {
                    Dictionary<string, byte[]> attachmentsdata = null;
                    if (this.emailMessage.Attachments.Count > 0)
                    {
                        attachmentsdata = new Dictionary<string, byte[]>(this.emailMessage.Attachments.Count);
                        for (int i = 0; i < this.emailMessage.Attachments.Count; i++)
                        {
                            Attachment attachment = this.emailMessage.Attachments[i];
                            Stream readStream = attachment.GetContentReadStream();
                            byte[] bytes = null;
                            if (readStream.Length > 0) {
                                bytes = new byte[readStream.Length];
                                readStream.Read(bytes, 0, bytes.Length);
                            } else
                                bytes = Encoding.ASCII.GetBytes(" ");
                            attachmentsdata.Add(attachment.FileName, bytes);
                        }
                    }
                    
                    parser.changeBodyOfTnef(args.MailItem.Message.TnefPart, schedulingInfo);

                    if (attachmentsdata != null)
                    {
                        foreach (KeyValuePair<string, byte[]> attachmentdata in attachmentsdata)
                        {
                            Attachment attachment = this.emailMessage.Attachments.Add(attachmentdata.Key);
                            Stream attachmentStream = attachment.GetContentWriteStream();
                            attachmentStream.Write(attachmentdata.Value, 0, attachmentdata.Value.Length);
                            attachmentStream.Flush();
                            attachmentStream.Close();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                RvLogger.DebugWrite(ex.Message);
                RvLogger.DebugWrite(ex.StackTrace);
                string baseFailCode = ex.Message;
                sendBackAMail(source, schedulingInfo, baseFailCode);
            }
            finally
            {
                RvLogger.DebugWrite("Start to agentAsyncContext.Complete()================================================");
                agentAsyncContext.Complete();
                RvLogger.DebugWrite("Complete agentAsyncContext.Complete()================================================");
                RvLogger.DebugWrite("Leave transport agent, from: " + schedulingInfo.delegatorEmailAddr + ", subject: " + schedulingInfo.subject);
            }
        }

        private string convertToUnicode(String source)
        {
            StringBuilder result = new StringBuilder();
            for (int i = 0, len = source.Length; i < len; i++)
            {
                char aChar = source[i];
                if (aChar >= 0x4e00 && aChar <= 0x9fbb)
                {
                    result.Append("&#").Append("" + (int)aChar).Append(";");
                }
                else
                    result.Append(aChar);
            }

            return result.ToString();
        }

        private bool isScopia(SchedulingInfo schedulingInfo)
        {
            if (!string.IsNullOrEmpty(schedulingInfo.location))
                foreach(string mark in MARK_OF_SCOPIA_LOCATION) {
                    if (schedulingInfo.location.ToLower().IndexOf(mark) != -1)
                        return true;
                }

            if (!string.IsNullOrEmpty(schedulingInfo.subject))
                foreach (string mark in MARK_OF_SCOPIA_SUBJECT)
                {
                    if (schedulingInfo.subject.ToLower().IndexOf(mark) != -1)
                        return true;
                }

            if (!string.IsNullOrEmpty(schedulingInfo.recipents))
                foreach (string mark in MARK_OF_SCOPIA_RECEIPENT)
                {
                    if (schedulingInfo.recipents.ToLower().IndexOf(mark) != -1)
                        return true;
                }

            if (!string.IsNullOrEmpty(schedulingInfo.conferenceID))
                return true;

            return false;
        }

        private RequestType getRequestType(string messageClass)
        {
            RequestType requestType = RequestType.CreateMeeting;
            if ("IPM.Schedule.Meeting.Request".Equals(messageClass))
                requestType = RequestType.CreateMeeting;
            else if ("IPM.Schedule.Meeting.Canceled".Equals(messageClass))
                requestType = RequestType.CancelMeeting;
            else
                requestType = RequestType.Other;

            return requestType;
        }

        private void parseRecipentsChanged(SchedulingInfo schedulingInfo)
        {
            RvLogger.DebugWrite("parseRecipentsChanged schedulingInfo.displayNamesFromTnef=============" + schedulingInfo.displayNamesFromTnef);
            if (string.IsNullOrEmpty(schedulingInfo.displayNamesFromTnef.Trim()))
                return;

            string[] names = schedulingInfo.displayNamesFromTnef.Split(new string[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
            List<string> list = names.ToList<string>();
            list.Sort();
            StringBuilder displayNamesStrb = new StringBuilder("");
            foreach (string name in list)
            {
                displayNamesStrb.Append(name).Append("; ");
            }
            string displayNamesStr = displayNamesStrb.ToString().Trim();
            RvLogger.DebugWrite("parseRecipentsChanged displayNamesStr=============" + displayNamesStr);
            if (displayNamesStr.Length < 2)
                return;
            schedulingInfo.displayNamesFromTnef = displayNamesStr.Substring(0, displayNamesStr.Length - 1);
            RvLogger.DebugWrite("DisplayNames emails are displayNamesFromTnef: " + schedulingInfo.displayNamesFromTnef);
            if (!schedulingInfo.displayNames.Equals(schedulingInfo.displayNamesFromTnef))
                schedulingInfo.isRecipentsChanged = true;
        }

        private void sendBackAMail(QueuedMessageEventSource source, SchedulingInfo schedulingInfo, string baseFailCode)
        {
            try
            {
                RvLogger.DebugWrite("Send back a mail to the sender.");
                RvLogger.DebugWrite("Notification mail code: " + baseFailCode);
                DirectoryInfo directoryInfo = new DirectoryInfo(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
                string PickupPath = directoryInfo.Parent.Parent.FullName + @"\Pickup\";
                if (settingsInfo["ExchangeServer.PickupDirectoryPath"] != null)
                {
                    PickupPath = settingsInfo["ExchangeServer.PickupDirectoryPath"];
                    if(!PickupPath.EndsWith(@"\"))
                         PickupPath = PickupPath + @"\";
                }
                

                if (!Directory.Exists(PickupPath))
                {
                    Directory.CreateDirectory(PickupPath);
                }
                EmailMessage origMsg = this.emailMessage;

                EmailMessage newMsg = EmailMessage.Create();
                //newMsg.MessageId = "11111111112222222222";
 
                string subjectFailCode = getFailCode(schedulingInfo, "subject.fail");

                if ("SUCCESS_NOTIFICATION".Equals(baseFailCode))
                    subjectFailCode = getFailCode(schedulingInfo, "subject.success");
                else if ("RESOURCE_SHORTAGE".Equals(baseFailCode))
                    subjectFailCode = getFailCode(schedulingInfo, "subject.resource.shortage");
                else if (!notResourceProblem(baseFailCode))
                    baseFailCode = "ERROR_RESOURCE_SHORTAGE";

                string failCode = getFailCode(schedulingInfo, baseFailCode);

                newMsg.Subject = messages[subjectFailCode];
                newMsg.To.Clear();
                newMsg.To.Add(new EmailRecipient(
                    origMsg.Sender.DisplayName,
                    origMsg.Sender.SmtpAddress));

                newMsg.From = new EmailRecipient("Service Account",
                    SERVER_ACCOUNT);
                Stream writer = newMsg.Body.GetContentWriteStream();
                StreamWriter sw = new StreamWriter(writer, Encoding.Unicode);
                if (schedulingInfo != null)
                {
                    sw.WriteLine("Subject : " + schedulingInfo.subject);
                    sw.WriteLine();
                }
                string message = messages.ContainsKey(failCode) ? messages[failCode] : null;
                if (string.IsNullOrEmpty(message))
                    message = messages["common.message"];

                RvLogger.DebugWrite("Notification mail message: " + message);
                sw.WriteLine(message);
                sw.Close();
                writer.Close();
                SaveMessage(newMsg, String.Concat(PickupPath, schedulingInfo.delegatorEmailAddr + "-" + DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm-ss-ms") + ".eml"));
                RvLogger.DebugWrite("Notification mail is sent.");
                // Cancel the original message
                if (!"SUCCESS_NOTIFICATION".Equals(baseFailCode) && !"RESOURCE_SHORTAGE".Equals(baseFailCode)) 
                    source.Delete();
            }
            catch (Exception ex)
            {
                RvLogger.DebugWrite(ex.StackTrace);
                RvLogger.DebugWrite(ex.Message);
            }
            RvLogger.DebugWrite("Send back a mail to the sender completed.");
        }

        private bool notResourceProblem(string baseFailCode) {
            return !"ERROR_USER_CONFLICT".Equals(baseFailCode)
                    && !"ERROR_GK_RESOURCE_SHORTAGE".Equals(baseFailCode)
                    && !"ERROR_GW_NOT_AVAILABLE".Equals(baseFailCode)
                    && !"ERROR_GW_OVER_MEMBER_CALL_LIMIT".Equals(baseFailCode)
                    && !"ERROR_GW_RESOURCE_SHORTAGE".Equals(baseFailCode)
                    && !"ERROR_GW_PARTY_RATE_NOT_MATCH".Equals(baseFailCode)
                    && !"ERROR_GW_SERVICE_NOT_SUPPORT".Equals(baseFailCode)
                    && !"ERROR_LICENSE_CONCURRENT_CALLS_LIMIT".Equals(baseFailCode)
                    && !"ERROR_LICENSE_CONCURRENT_GUEST_MOBILE_CALL_LIMIT".Equals(baseFailCode)
                    && !"ERROR_LICENSE_CONCURRENT_REGISTED_MOBILE_CALL_LIMIT".Equals(baseFailCode)
                    && !"ERROR_MASTER_MCU_RESOURCE_SHORTAGE".Equals(baseFailCode)
                    && !"ERROR_MCU_OVER_MEMBER_CALL_LIMIT".Equals(baseFailCode)
                    && !"ERROR_MCU_RESOURCE_SHORTAGE".Equals(baseFailCode)
                    && !"ERROR_NETWORK_CONGESTION".Equals(baseFailCode)
                    && !"ERROR_PARTY_NOT_AVAILABLE".Equals(baseFailCode)
                    && !"ERROR_NETWORK_NOT_ARRIVE".Equals(baseFailCode);
        }

        public static void SaveMessage(EmailMessage msg, string filePath)
        {
            try
            {
                FileStream file = new FileStream(filePath, System.IO.FileMode.Create);
                msg.MimeDocument.WriteTo(file);
                file.Close();
            }
            catch (Exception ex)
            {
                RvLogger.DebugWrite(Environment.NewLine + "3-------------------------------------------------------------------------------" + ex.StackTrace);
            }
        }

        private static Dictionary<string, string> getPropertiesFromPropertiesFile(string filePath)
        {
            StreamReader reader = File.OpenText(filePath);
            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            string line = null;
            while ((line = reader.ReadLine()) != null)
            {
                int index = line.IndexOf("=");
                if (-1 == index)
                    continue;
                if(line.StartsWith("-"))
                    continue;
                if (line.StartsWith("/"))
                    continue;
                string key = line.Substring(0, index);
                string value = line.Substring(index + 1);
                keyValues.Add(key.Trim(), value.Trim());
            }
            reader.Close();

            return keyValues;
        }

        private static string EncodeBase64(string code_type, string code)
        {
            string encode = "";
            byte[] bytes = Encoding.GetEncoding(code_type).GetBytes(code);
            try
            {
                encode = System.Convert.ToBase64String(bytes);
            }
            catch
            {
                encode = code;
            }
            return "outlook-addin:" + encode;
        }

        private static webServiceData.ExchangeService getWebService(SchedulingInfo schedulingInfo){

            long now = DateTime.UtcNow.Ticks;
            RvLogger.DebugWrite("start to create webservice======================================");
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

            webServiceData.ExchangeService service = new webServiceData.ExchangeService(webServiceData.ExchangeVersion.Exchange2010_SP2);

            String userAccount = null;
            if (settingsInfo.ContainsKey("ExchangeServer.email.account"))
                userAccount = settingsInfo["ExchangeServer.email.account"];

            if(settingsInfo.ContainsKey("ExchangeServer.user.account"))
                userAccount = settingsInfo["ExchangeServer.user.account"];
            String userPassword = null;
            if (settingsInfo.ContainsKey("ExchangeServer.email.password"))
                userPassword = settingsInfo["ExchangeServer.email.password"];
            if (settingsInfo.ContainsKey("ExchangeServer.user.password"))
                userPassword = settingsInfo["ExchangeServer.user.password"];

            String userEmail = null;
            if (settingsInfo.ContainsKey("ExchangeServer.email.account"))
                userEmail = settingsInfo["ExchangeServer.email.account"];
            if (settingsInfo.ContainsKey("ExchangeServer.user.email"))
                userEmail = settingsInfo["ExchangeServer.user.email"];

            RvLogger.DebugWrite("EXCHANGE_USER: " + userAccount);
            RvLogger.DebugWrite("EXCHANGE_EMAIL: " + userEmail);
            RvLogger.DebugWrite("EXCHANGE_password: " + userPassword);

            service.Credentials = new NetworkCredential(userAccount, userPassword);

            try
            {
                service.AutodiscoverUrl(userEmail, RedirectionUrlValidationCallback);
            }
            catch (Exception e)
            {
                RvLogger.InfoWrite("AutodiscoverUrl Error: " + e.Message);
                RvLogger.InfoWrite(e.StackTrace);
                service.TraceEnabled = true;
                service.TraceListener = new TraceListener();
                service.TraceFlags = webServiceData.TraceFlags.All;
                try
                {
                    service.AutodiscoverUrl(userEmail, RedirectionUrlValidationCallback);
                }
                catch (Exception e1)
                {
                    RvLogger.InfoWrite("AutodiscoverUrl Error1: " + e1.Message);
                    RvLogger.InfoWrite(e1.StackTrace);
                    if (settingsInfo.ContainsKey("ExchangeServer.ewsurl"))
                    {
                        String ewsurl = settingsInfo["ExchangeServer.ewsurl"];
                        service.Url = new Uri(ewsurl);
                    }
                    if (settingsInfo.ContainsKey("ExchangeServer.trace.all"))
                    {
                        String traceAll = settingsInfo["ExchangeServer.trace.all"];
                        if (traceAll != null && !traceAll.Equals("true"))
                        {
                            service.TraceEnabled = false;
                            service.TraceListener = null;
                        }
                    }

                }
            }
            RvLogger.DebugWrite("Discovered service URL:"+service.Url);
            if (null != schedulingInfo)
                service.ImpersonatedUserId = new webServiceData.ImpersonatedUserId(webServiceData.ConnectingIdType.SmtpAddress, schedulingInfo.delegatorEmailAddr);

            RvLogger.DebugWrite("Finished to create webservice======================================time:" + (DateTime.UtcNow.Ticks - now));

            return service;
        }

  
        

        private icm.XmlApi.scheduleReportType changeMail(QueuedMessageEventSource source, QueuedMessageEventArgs args, SchedulingInfo schedulingInfo, long now)
        {
            icm.XmlApi.scheduleReportType result = null;
            icm.XmlApi.userType userInfo = null;
            string dialingInfoStr = null;
            bool isCreate = isCreateMeetingRequest(schedulingInfo);
            webServiceData.ExchangeService service = null;

            try
            {
                userInfo = getUserInfo(schedulingInfo.delegatorEmailAddr);
                if (null == userInfo)
                    throw new ScopiaMeetingAddInException(ScopiaMeetingAddInException.ERROR_MESSAGE_USER_NOT_FOUND);
                else if (!userInfo.Schedulable)
                    throw new ScopiaMeetingAddInException(ScopiaMeetingAddInException.ERROR_MESSAGE_HAVE_NO_PERMISSION);
                else if (isCreate)
                {
                    service = getWebService(schedulingInfo);
                    if(null == service)
                        throw new ScopiaMeetingAddInException(ScopiaMeetingAddInException.ERROR_MESSAGE_INCORRECT_EWS_CONFIGURATION);
                }
            }
            catch (Exception ex) {
                sendBackAMail(source, schedulingInfo, ex.Message);
                if (IS_DELETE_ERROR) {
                    if (null == service)
                        service = getWebService(schedulingInfo);
                    if (null == service)
                        return result;
                    result = new icm.XmlApi.scheduleReportType();
                    result.Success = false;
                    result.Detail = "";
                    //Do not add a cancel request into the task list,
                    if (schedulingInfo.requestType == RequestType.CreateMeeting)
                    {
                        object[] parameters = new object[5];
                        parameters[0] = schedulingInfo;
                        parameters[1] = null;
                        parameters[2] = result;
                        parameters[3] = now;
                        parameters[4] = service;
                        lock (ewsTasks)
                        {
                            RvLogger.DebugWrite("add the task in=============" + schedulingInfo.subject);
                            ewsTasks.AddLast(parameters);
                        }
                    }
                }
                return result;
            }

            bool successfullyRescheduled = false;
            if (schedulingInfo.requestType == RequestType.CreateMeeting)
            {
                try {
                    icm.XmlApi.virtualRoomType[] virtualRooms = this.getVirtualRoom(userInfo);
                    if (null == virtualRooms && isCreate)
                        throw new ScopiaMeetingAddInException(ScopiaMeetingAddInException.ERROR_MESSAGE_NO_VIRTUALROOM);
                    if (isCreate) {
                        icm.XmlApi.virtualRoomType defaultVirtualRoom = null;
                        foreach (icm.XmlApi.virtualRoomType virtualRoom in virtualRooms) {
                            if (virtualRoom.Default) {
                                defaultVirtualRoom = virtualRoom;
                                break;
                            }
                        }
                        if (null == defaultVirtualRoom)
                            defaultVirtualRoom = virtualRooms[0];

                        if (defaultVirtualRoom.OneTimePINRequired) {
                            if (string.IsNullOrEmpty(schedulingInfo.meetingPin))
                                schedulingInfo.meetingPin = CreateRandomPWD(6);
                        }

                        icm.XmlApi.dialingInfoType dialingInfo = this.getDialingInfo(schedulingInfo, userInfo, defaultVirtualRoom);
                        if (USE_HTML)
                            dialingInfoStr = dialingInfo.DescriptionOfHTML;
                        else
                            dialingInfoStr = dialingInfo.Description;

                        if (!USE_HTML && null != dialingInfoStr)
                        {
                            dialingInfoStr = dialingInfoStr.Replace("\n", "<br>");
                            dialingInfoStr = dialingInfoStr.Replace("\r", "");
                            dialingInfoStr = dialingInfoStr.Replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;");
                        }
                        schedulingInfo.dialingInfo = dialingInfoStr;
                    }
                    result = this.scheduleMeeting(schedulingInfo, schedulingInfo.subject, false);
                    if (!result.Success && IS_RESCHEDULE && !notResourceProblem(string.IsNullOrEmpty(result.ErrorCode) ? (result.Detail == null ? "" : result.Detail) : result.ErrorCode))
                    {
                        result = this.scheduleMeeting(schedulingInfo, schedulingInfo.subject, true);
                        if (result.Success)
                        {
                            successfullyRescheduled = true;
                        } 
                    }
                } catch (Exception ex) {
                    result = new icm.XmlApi.scheduleReportType();
                    result.Success = false;
                    result.ErrorCode = ex.Message;
                    result.Detail = ex.Message;
                }

                result.Detail = result.Detail == null ? "" : result.Detail;

                if ((isCreate && result.Success) || (!result.Success && IS_DELETE_ERROR))
                {
                    if (null == service)
                        service = getWebService(schedulingInfo);
                    if (null == service)
                        return result;

                    object[] parameters = new object[5];
                    parameters[0] = schedulingInfo;
                    parameters[1] = dialingInfoStr;
                    parameters[2] = result;
                    parameters[3] = now;
                    parameters[4] = service;
                    lock (ewsTasks)
                    {
                        RvLogger.DebugWrite("add the task in=============" + schedulingInfo.subject);
                        ewsTasks.AddLast(parameters);
                    }
                }
            }
            else
            {
                result = this.deleteMeeting(schedulingInfo, true);
                result.Detail = result.Detail == null ? "" : result.Detail;
            }

            if (false == result.Success && ("CONF_NOT_FOUND".Equals(result.Detail) || "CANCELLED".Equals(result.Detail)))
                result.Success = true;

            if (false == result.Success)
            {
                sendBackAMail(source, schedulingInfo, string.IsNullOrEmpty(result.ErrorCode) ? result.Detail : result.ErrorCode);
            }
            else if (successfullyRescheduled)
                sendBackAMail(source, schedulingInfo, "RESOURCE_SHORTAGE");
            else if (IS_SEND_SUCCESS_MAIL)
            {
                sendBackAMail(source, schedulingInfo, "SUCCESS_NOTIFICATION");
            }

            return result;
        }

        
        private string CreateRandomPWD(int codeCount)
        {
            Random rand = new Random();
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < codeCount; i++)
            {
                result.Append((char)rand.Next('0', '9' + 1));
            }
            return result.ToString();
        }

        private static void startThreadForEWSInvoking() {
            Thread t = new Thread(new ThreadStart(delegate {
                LinkedList<object[]> needToHandle = new LinkedList<object[]>();
                int milliSecondsForSleeping = HOWMANYSECONDS_TO_HANDLE_TASKS * 1000;
                while (true)
                {
                    Thread.Sleep(milliSecondsForSleeping);
                    lock (ewsTasks)
                    {
                        long now = DateTime.UtcNow.Ticks;
                        foreach (object[] parameters in ewsTasks)
                        {
                            long value = (long) parameters[3];
                            if ((now - value) / 10000000 > AFTER_HOWMANYSECONDS_TO_GETAPPOINTMENT) {
                                needToHandle.AddLast(parameters);
                            }
                            else {
                                break;
                            }
                        }

                        foreach (object[] parameters in needToHandle)
                        {
                            Thread.Sleep(HOW_MANY_MILLISECONDS_TO_CALL_EWS);
                            ewsTasks.Remove(parameters);
                            ThreadPool.QueueUserWorkItem(new WaitCallback(ThreadMethod), parameters);
                        }
                        needToHandle.Clear();
                    }

                }
            }));

            t.Start();
        }

        public static void ThreadMethod(object paraList)
        {
            object[] obj = (object[]) paraList;
            SchedulingInfo schedulingInfo = (SchedulingInfo) obj[0];
            RvLogger.DebugWrite("handle task===================" + schedulingInfo.subject);
            string dialingInfoStr= (string) obj[1];
            icm.XmlApi.scheduleReportType result = (icm.XmlApi.scheduleReportType) obj[2];
            long now = (long) obj[3];
            webServiceData.ExchangeService service = (webServiceData.ExchangeService) obj[4];
            RvLogger.DebugWrite("start searching appointment===================" + schedulingInfo.subject);
            try
            {
                long nowlong = DateTime.UtcNow.Ticks;
                RvLogger.DebugWrite("start to search appointment======================================");
                webServiceData.Appointment theCurrentAppointment = getCurrentAppointment(schedulingInfo, service);
                if (null == theCurrentAppointment)
                {
                    if (1 != now)
                    {
                        obj[3] = 1L;
                        lock (ewsTasks)
                        {
                            ewsTasks.AddLast(obj);
                        }
                        RvLogger.InfoWrite("Re-add the task into the queue: " + schedulingInfo.subject);
                    }else
                        RvLogger.InfoWrite("searched appointment===================null");
                    return;
                }
                RvLogger.DebugWrite("searched appointment===================" + theCurrentAppointment);
                updateAppointment(theCurrentAppointment, schedulingInfo, service, dialingInfoStr, result);
                RvLogger.DebugWrite("Finished to update appointment======================================time:" + (DateTime.UtcNow.Ticks - nowlong));
            }
            catch (Exception ex) {
                RvLogger.DebugWrite(ex.Message);
                RvLogger.DebugWrite(ex.StackTrace);
            }
        }

        private static void checkLog(object source, System.Timers.ElapsedEventArgs e)
        {
            RvLogger.DebugWrite("checklog=========================================");
            try
            {
                FileInfo log = new FileInfo(RvLogger.DefaultFileName);
                if (log.Length > 1024 * 1000 * LOG_FILE_SIEZ_MB)
                {
                    RvLogger.InfoWrite("Start to open a new log.");
                    RvLogger.Open(RvLogger.DefaultFileName, true);
                    RvLogger.InfoWrite("Finish to open a new log.");
                }
            } catch (Exception ex){
                RvLogger.InfoWrite(ex.Message);
                RvLogger.InfoWrite(ex.StackTrace);
            }
        }

        private static webServiceData.Appointment getCurrentAppointment(SchedulingInfo schedulingInfo, webServiceData.ExchangeService service)
        {
            webServiceData.Appointment theCurrentAppointment = null;
            //service.ImpersonatedUserId = new webServiceData.ImpersonatedUserId(webServiceData.ConnectingIdType.SmtpAddress, schedulingInfo.senderEmailAddr);

            //long startSearchTime = DateTime.UtcNow.Ticks;
            //do
            //{
                if (schedulingInfo.meetingType == MeetingType.Ocurrence)
                {
                    webServiceData.Appointment appointment = searchAppointment(schedulingInfo, service);
                    if (null != appointment)
                    {
                        theCurrentAppointment = appointment;
                    }
                }
                else
                {
                    webServiceData.Appointment appointment = searchAppointmentByFilter(schedulingInfo, service, null);
                    if (null != appointment)
                    {
                        theCurrentAppointment = appointment;
                    }
                }

            //    if (null == theCurrentAppointment) Thread.Sleep(5000);
            //} while (null == theCurrentAppointment && (DateTime.UtcNow.Ticks - startSearchTime) / 10000000 <= 60);

            if (null == theCurrentAppointment)
            {
                //do
                //{
                    webServiceData.Appointment appointment = searchAppointment(schedulingInfo, service);
                    //if (null != appointment)
                    //{
                        theCurrentAppointment = appointment;
                    //}
                    //else
                    //{
                    //    Thread.Sleep(5000);
                    //}
                //} while (null == theCurrentAppointment && (DateTime.UtcNow.Ticks - startSearchTime) / 10000000 <= 10);
            }

            return theCurrentAppointment;
        }

        private static webServiceData.Appointment searchAppointment(SchedulingInfo schedulingInfo, webServiceData.ExchangeService service)
        {
            RvLogger.DebugWrite("enter searchAppointment==============");
            webServiceData.Mailbox mailBox = new webServiceData.Mailbox(schedulingInfo.delegatorEmailAddr);
            webServiceData.FolderId folderID = new webServiceData.FolderId(webServiceData.WellKnownFolderName.Calendar, mailBox);
            webServiceData.CalendarFolder folder = webServiceData.CalendarFolder.Bind(service, folderID);
            webServiceData.CalendarView view = new webServiceData.CalendarView(schedulingInfo.startDate, schedulingInfo.endDate);
            webServiceData.PropertySet propertySet = new webServiceData.PropertySet(webServiceData.BasePropertySet.FirstClassProperties);
            view.PropertySet = propertySet;
            webServiceData.FindItemsResults<webServiceData.Appointment> results = folder.FindAppointments(view);

            RvLogger.DebugWrite("results==============" + (null == results.Items ? "0" : "" + results.Items.Count));

            foreach (webServiceData.Item item in results)
            {
                try
                {
                    webServiceData.Appointment appointment = (webServiceData.Appointment)item;
                    if (string.IsNullOrEmpty(schedulingInfo.location)) schedulingInfo.location = "";
                    if (string.IsNullOrEmpty(appointment.Location)) appointment.Location = "";
                    if (string.IsNullOrEmpty(schedulingInfo.subject)) schedulingInfo.subject = "";
                    if (string.IsNullOrEmpty(appointment.Subject)) appointment.Subject = "";
                    if (schedulingInfo.location == appointment.Location 
                        && appointment.Subject == schedulingInfo.subject
                        && 0 == appointment.Start.ToUniversalTime().CompareTo(schedulingInfo.startDate.ToUniversalTime())
                        && 0 == appointment.End.ToUniversalTime().CompareTo(schedulingInfo.endDate.ToUniversalTime()))
                    {
                        return appointment;
                    }
                }
                catch (ScopiaMeetingAddInException ex)
                {
                    throw ex;
                }
            }

            return null;
        }

        private static webServiceData.Appointment searchAppointmentByFilter(SchedulingInfo schedulingInfo, webServiceData.ExchangeService service, string appointmentID)
        {
            RvLogger.DebugWrite("enter searchAppointmentByFilter==============");
            List<webServiceData.SearchFilter> searchORFilterCollection = new List<webServiceData.SearchFilter>();
            if (null != appointmentID)
                searchORFilterCollection.Add(new webServiceData.SearchFilter.IsEqualTo(webServiceData.EmailMessageSchema.Id, new webServiceData.ItemId(appointmentID)));
            else {
                searchORFilterCollection.Add(new webServiceData.SearchFilter.IsEqualTo(webServiceData.EmailMessageSchema.Subject, schedulingInfo.subject));
                searchORFilterCollection.Add(new webServiceData.SearchFilter.IsEqualTo(webServiceData.EmailMessageSchema.From, schedulingInfo.delegatorEmailAddr));
                searchORFilterCollection.Add(new webServiceData.SearchFilter.IsGreaterThan(webServiceData.EmailMessageSchema.LastModifiedTime, DateTime.UtcNow.AddHours(-25)));
                if (!string.IsNullOrEmpty(schedulingInfo.storeEntryId))
                    searchORFilterCollection.Add(new webServiceData.SearchFilter.IsEqualTo(webServiceData.EmailMessageSchema.StoreEntryId, schedulingInfo.storeEntryId));
                if (!string.IsNullOrEmpty(schedulingInfo.conversationKey))
                    searchORFilterCollection.Add(new webServiceData.SearchFilter.IsEqualTo(webServiceData.EmailMessageSchema.ConversationId, schedulingInfo.conversationKey));
            }
            RvLogger.DebugWrite("enter searchAppointmentByFilter==============1");
            webServiceData.SearchFilter searchFilter = new webServiceData.SearchFilter.SearchFilterCollection(webServiceData.LogicalOperator.And, searchORFilterCollection.ToArray());
            RvLogger.DebugWrite("enter searchAppointmentByFilter==============2 " + schedulingInfo.delegatorEmailAddr);
            webServiceData.Mailbox mailBox = new webServiceData.Mailbox(schedulingInfo.delegatorEmailAddr);
            RvLogger.DebugWrite("enter searchAppointmentByFilter==============3");
            webServiceData.FolderId folderID = new webServiceData.FolderId(webServiceData.WellKnownFolderName.Calendar, mailBox); //No need to set mail since the service already know it.
            //webServiceData.FolderId folderID = new webServiceData.FolderId(webServiceData.WellKnownFolderName.Calendar);
            RvLogger.DebugWrite("enter searchAppointmentByFilter==============4");
            webServiceData.FindItemsResults<webServiceData.Item> results = service.FindItems(
                                    folderID,
                                    searchFilter,
                                    new webServiceData.ItemView(100));
            RvLogger.DebugWrite("enter searchAppointmentByFilter==============5");
            RvLogger.DebugWrite("results searchAppointmentByFilter==============" + (null == results.Items ? "0" : "" + results.Items.Count));

            foreach (webServiceData.Item item in results) {
                try {
                    webServiceData.Appointment appointment = (webServiceData.Appointment)item;
                    if (string.IsNullOrEmpty(schedulingInfo.location)) schedulingInfo.location = "";
                    if (string.IsNullOrEmpty(appointment.Location)) appointment.Location = "";
                    if (schedulingInfo.location == appointment.Location
                        && schedulingInfo.startDate.ToUniversalTime().Equals(appointment.Start.ToUniversalTime())
                        && schedulingInfo.endDate.ToUniversalTime().Equals(appointment.End.ToUniversalTime()))
                    {
                        RvLogger.DebugWrite("lastModifiedTime1===================" + appointment.LastModifiedTime);
                        return appointment;
                    }
                }
                catch (ScopiaMeetingAddInException ex) {
                    throw ex;
                }
            }

            return null;
        }

        private icm.XmlApi.userType getUserInfo(string email){
            try
            {
                adapter = new Adapter(settingsInfo["ScopiaManagement.url"],
                                    new NetworkCredential(settingsInfo["ScopiaManagement.loginID"],
                                                            EncodeBase64("utf-8", settingsInfo["Scopiamanagement.password"]), ""));
                icm.XmlApi.userType userInfo = adapter.GetUserInfo(false, email);
                return userInfo;
            }
            catch (WebException ex) {
                if(ex.Message.IndexOf("401") != -1)
                    throw new ScopiaMeetingAddInException(ScopiaMeetingAddInException.ERROR_MESSAGE_WRONG_USERNAME_PASSWORD);
                else
                    throw new ScopiaMeetingAddInException(ScopiaMeetingAddInException.ERROR_MESSAGE_CANNOT_CONNECT2SCOPIA);
            }
            catch (Exception ex)
            {
                throw new ScopiaMeetingAddInException(ScopiaMeetingAddInException.ERROR_MESSAGE_GETUSER_ERROR + ex.Message);
            }

        }

        private static void updateAppointment(webServiceData.Appointment appointment, SchedulingInfo schedulingInfo, webServiceData.ExchangeService service, string dialingInfoStr, icm.XmlApi.scheduleReportType result)
        {
                if (null == appointment) 
                    return;

                string uniqueId = null;
                if (schedulingInfo.meetingType == MeetingType.Ocurrence)
                {
                    appointment = webServiceData.Appointment.Bind(service, appointment.Id, new webServiceData.PropertySet(webServiceData.BasePropertySet.FirstClassProperties) { RequestedBodyType = webServiceData.BodyType.HTML });
                }
                else if (appointment.AppointmentType == webServiceData.AppointmentType.RecurringMaster)
                {
                    uniqueId = appointment.Id.UniqueId;
                    appointment = webServiceData.Appointment.Bind(service, appointment.Id);
                }
                else
                    appointment = webServiceData.Appointment.Bind(service, appointment.Id);

                saveConferenceInfoIntoAppointment(result, appointment, schedulingInfo, service, dialingInfoStr, uniqueId);
        }

        private static void saveConferenceInfoIntoAppointment(icm.XmlApi.scheduleReportType result, webServiceData.Appointment appointment, SchedulingInfo schedulingInfo, webServiceData.ExchangeService service, string dialingInfoStr, string uniqueId)
        {
                string hashString = RvScopiaMeeting.GetHashString(schedulingInfo, schedulingInfo.subject);
                if (schedulingInfo.requestType != RequestType.CancelMeeting)
                {
                    string conferenceID = null == result || !result.Success ? "empty" : result.ConferenceId;
                    if (isCreateMeetingRequest(schedulingInfo))
                    {
                        if (!result.Success && IS_DELETE_ERROR) {
                            appointment.Delete(webServiceData.DeleteMode.HardDelete, webServiceData.SendCancellationsMode.SendToNone);
                            RvLogger.DebugWrite("delete appointment: " + schedulingInfo.subject);
                            return;
                        }
                        appointment.MeetingWorkspaceUrl = CONFERENCE_ID + ":" + conferenceID + ":" + hashString + "$" + schedulingInfo.recurrenceHashInfo;
                        if (USE_HTML)
                            appointment.Body.BodyType = webServiceData.BodyType.HTML;

                        List<webServiceData.Attachment> attachments = new List<webServiceData.Attachment>(appointment.Attachments.Count);
                        foreach (webServiceData.Attachment attachment in appointment.Attachments)
                            appointment.Body.Text = appointment.Body.Text.Replace(attachment.Name, "");
                        appointment.Body.Text = appointment.Body + Environment.NewLine + "<br>" + dialingInfoStr;
                        if (schedulingInfo.meetingType == MeetingType.Reccurence)
                            appointment.MeetingWorkspaceUrl = "#" + appointment.MeetingWorkspaceUrl;
                    }
                    else
                    {
                        if (!result.Success && IS_DELETE_ERROR) {
                            appointment.Delete(webServiceData.DeleteMode.HardDelete, webServiceData.SendCancellationsMode.SendToNone);
                            RvLogger.DebugWrite("delete appointment: " + schedulingInfo.subject);
                            return;
                        }
                    }                    
                }

                updateAppointment(appointment);
        }

        private static void updateAppointment(webServiceData.Appointment appointment)
        {
            int j = 3;
            do
            {
                try
                {
                    appointment.Update(webServiceData.ConflictResolutionMode.AlwaysOverwrite, webServiceData.SendInvitationsOrCancellationsMode.SendToNone);
                    break;
                }
                catch (Exception ex)
                {
                    RvLogger.DebugWrite("Exception==============" + ex.Message);
                    RvLogger.DebugWrite(ex.StackTrace);
                }
                j--;
            } while (j >= 0);
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        private icm.XmlApi.dialingInfoType getDialingInfo(SchedulingInfo schedulingInfo, icm.XmlApi.userType userInfo, icm.XmlApi.virtualRoomType defaultVirtualRoom)
        {
            bool isCreate = isCreateMeetingRequest(schedulingInfo);
            return adapter.GetDialingInfo(userInfo, schedulingInfo, defaultVirtualRoom, isCreate, schedulingInfo.requestType == RequestType.CancelMeeting);
        }

        private icm.XmlApi.virtualRoomType[] getVirtualRoom(icm.XmlApi.userType userInfo)
        {
            icm.XmlApi.getVirtualRoomResponseType virtualRoomResponse = adapter.GetVirtualRooms(userInfo);
            return virtualRoomResponse.VirtualRoom;
        }

        private icm.XmlApi.scheduleReportType scheduleMeeting(SchedulingInfo schedulingInfo, string subject, bool removeResource)
        {
            icm.XmlApi.conferenceType conferenceType = new icm.XmlApi.conferenceType();
            conferenceType.Subject = subject; 
            conferenceType.Duration = "PT" + schedulingInfo.endDate.Subtract(schedulingInfo.startDate).TotalMinutes + "M";
            conferenceType.UserEmail = schedulingInfo.delegatorEmailAddr;
            conferenceType.ConferenceId = schedulingInfo.conferenceID;
            if (!removeResource)
                conferenceType.OutlookResources = schedulingInfo.outlookResources;
            conferenceType.StartTime = schedulingInfo.startDate;
            conferenceType.StartTimeSpecified = true;
            conferenceType.TimeZoneId = schedulingInfo.timeZoneID;
            conferenceType.RecipentsChanged = schedulingInfo.isRecipentsChanged;
            conferenceType.AddRecipents = schedulingInfo.isAddRecipents;
            conferenceType.RemovePersonalTerminals = removeResource;
            if (!string.IsNullOrEmpty(schedulingInfo.meetingPin))
                conferenceType.AccessPIN = Encoding.UTF8.GetBytes(schedulingInfo.meetingPin);

            if (schedulingInfo.meetingType == MeetingType.Reccurence)
            {
                this.AddRecurrenceInfo(conferenceType, schedulingInfo);
            }
            else if (schedulingInfo.meetingType == MeetingType.Ocurrence)
            {
                if (schedulingInfo.hasPreStartDate)
                    conferenceType.OccurrenceOldStartTime = schedulingInfo.preStartDate;
                else
                    conferenceType.OccurrenceOldStartTime = schedulingInfo.startDate;
                conferenceType.OccurrenceOldStartTimeSpecified = true;
            }

            if (!string.IsNullOrEmpty(schedulingInfo.recipents))
                conferenceType.OutlookRecipents = schedulingInfo.recipents;
            if (!string.IsNullOrEmpty(schedulingInfo.outlookResources))
                conferenceType.OutlookResources = schedulingInfo.outlookResources;

            RvLogger.DebugWrite("conferenceType.outlookResources=====================" + conferenceType.OutlookResources);
            RvLogger.DebugWrite("Start scheduling this SCOPIA meeting, the subject is : " + subject);

            icm.XmlApi.scheduleReportType report = null;
            if (isCreateMeetingRequest(schedulingInfo)) {
                conferenceType.Description = "";
                report = adapter.ScheduleConference(conferenceType);
            } else {
                if (schedulingInfo.meetingType == MeetingType.Reccurence)
                {
                    report = adapter.ModifyConference(conferenceType);
                    report.Detail = report.Detail == null ? "" : report.Detail;
                    if (!report.Success && "PATTERN_CHANGED".Equals(report.Detail)) {
                        this.deleteMeeting(schedulingInfo, false);
                        conferenceType.ReccurencePatternChanged = true;
                        report = adapter.ScheduleConference(conferenceType);
                    } else if (!report.Success && !"PATTERN_CHANGED".Equals(report.Detail)) {
                        if (notResourceProblem(report.ErrorCode) || !IS_RESCHEDULE)
                        {
                            this.deleteMeeting(schedulingInfo, true);
                            //throw new ScopiaMeetingAddInException(string.IsNullOrEmpty(report.ErrorCode) ? report.Detail : report.ErrorCode);
                        }
                    }
                }
                else
                { 
                    report = adapter.ModifyConference(conferenceType);
                    report.Detail = report.Detail == null ? "" : report.Detail;
                    if (!report.Success && schedulingInfo.meetingType == MeetingType.Ocurrence && schedulingInfo.isRecipentsChanged
                        && !schedulingInfo.isAddRecipents && "CONF_NOT_FOUND".Equals(report.Detail))
                        report.Success = true;
                    if (!report.Success && (notResourceProblem(report.ErrorCode) || !IS_RESCHEDULE))
                    {
                        this.deleteMeeting(schedulingInfo, false);
                        //throw new ScopiaMeetingAddInException(string.IsNullOrEmpty(report.ErrorCode) ? report.Detail : report.ErrorCode);
                    }
                }
            }
            return report;
        }

        private icm.XmlApi.scheduleReportType deleteMeeting(SchedulingInfo schedulingInfo, bool deleteClientlessMapping) 
        {
            return adapter.CancelConference(schedulingInfo, deleteClientlessMapping);
        }

        private static bool isCreateMeetingRequest(SchedulingInfo schedulingInfo)
        {
            if (schedulingInfo.requestType == RequestType.CreateMeeting)
                return string.IsNullOrEmpty(schedulingInfo.conferenceID) || schedulingInfo.conferenceID.Equals("empty");
            return false;
        }

        private string getFailCode(SchedulingInfo schedulingInfo, string baseFailCode)
        {
            string code = "";
            if (null != schedulingInfo)
                if (schedulingInfo.requestType == RequestType.CreateMeeting) {
                    if(isCreateMeetingRequest(schedulingInfo))
                        code = baseFailCode + ".create";
                    else
                        code = baseFailCode + ".update";
                }else
                    code = baseFailCode + ".delete";

            return code;
        }

        private void AddRecurrenceInfo(icm.XmlApi.conferenceType conference, SchedulingInfo schedulingInfo)
        {
            if (conference == null)
                throw new ArgumentNullException("conference");

            RecurrencePattern recurrencePattern = schedulingInfo.recurrencePattern;
            //StringBuilder patternString = new StringBuilder("");
            //patternString.Append(appointment.Start.ToUniversalTime()).Append(appointment.End.ToUniversalTime());
            if (recurrencePattern.hasEnd)
            {
                conference.RecurrenceEnd = new icm.XmlApi.conferenceTypeRecurrenceEnd();
                conference.RecurrenceEnd.Item = recurrencePattern.numberOfOccurrences;
                //patternString.Append(recurrence.NumberOfOccurrences).Append(recurrence.StartDate.ToUniversalTime());
                RvLogger.InfoWrite("Recurrence will be scheduled with end date " + recurrencePattern.endDate + " - Occurrences [" + recurrencePattern.numberOfOccurrences + "].");
            }
            else
                RvLogger.InfoWrite("Recurrence will be scheduled without end date.");

            if (recurrencePattern.patternType == PatternType.DAILY)
            {
                icm.XmlApi.recurrenceDailyType recurrenceDailyType = new icm.XmlApi.recurrenceDailyType();
                if (recurrencePattern.dailyInterval != -1)
                    recurrenceDailyType.Item = recurrencePattern.dailyInterval;
                conference.Item = recurrenceDailyType;
                //patternString.Append("daily").Append(dailyPattern.Interval);
            }
            else if (recurrencePattern.patternType == PatternType.WEEKLY)
            {
                //patternString.Append("weekly").Append(weeklyPattern.Interval);
                if (recurrencePattern.weeklyInterval == 0)
                {
                    icm.XmlApi.recurrenceDailyType
                        recurrenceDailyType = new icm.XmlApi.recurrenceDailyType();
                    recurrenceDailyType.Item = true;

                    conference.Item = recurrenceDailyType;
                }
                else
                {
                    icm.XmlApi.recurrenceWeeklyType
                        recurrenceWeeklyType = new icm.XmlApi.recurrenceWeeklyType();
                    recurrenceWeeklyType.NumberOfEveryWeek = recurrencePattern.weeklyInterval;
                    recurrenceWeeklyType.DayOfWeek = recurrencePattern.daysOfWeek;
                    conference.Item = recurrenceWeeklyType;
                }

            }
            else if (recurrencePattern.patternType == PatternType.MONTHLY)
            {
                icm.XmlApi.recurrenceMonthlyType recurrenceMonthlyType = new icm.XmlApi.recurrenceMonthlyType();
                recurrenceMonthlyType.NumberOfEveryMonth = recurrencePattern.monthlyInterval;
                recurrenceMonthlyType.Item = recurrencePattern.dayOfMonth;
                conference.Item = recurrenceMonthlyType;

                    //patternString.Append("monthly").Append(monthlyPattern.Interval).Append(monthlyPattern.DayOfMonth);
            }
            else if (recurrencePattern.patternType == PatternType.RELATIVE_MONTHLY)
            {                  
                icm.XmlApi.recurrenceMonthlyType
                    recurrenceMonthlyType = new icm.XmlApi.recurrenceMonthlyType();

                icm.XmlApi.recurrenceMonthlyTypeDayOfNumberOfEveryMonth
                    dayOfNumberOfEveryMonth = new icm.XmlApi.recurrenceMonthlyTypeDayOfNumberOfEveryMonth();

                dayOfNumberOfEveryMonth.WeekOfMonth = recurrencePattern.weekOfMonth;

                dayOfNumberOfEveryMonth.DayOfWeek = recurrencePattern.dayOfWeekOfMonthly;


                recurrenceMonthlyType.NumberOfEveryMonth = recurrencePattern.monthlyInterval;
                recurrenceMonthlyType.Item = dayOfNumberOfEveryMonth;

                conference.Item = recurrenceMonthlyType;

                //patternString.Append("relativeMonthly").Append(monthlyPattern.DayOfTheWeekIndex)
                //    .Append(monthlyPattern.DayOfTheWeek)
                //    .Append(monthlyPattern.Interval);
            }
            else if (recurrencePattern.patternType == PatternType.YEARLY ||
                recurrencePattern.patternType == PatternType.RELATIVE_YEARLY)
            {
                RvLogger.FatalWrite("Trying to add NOT SUPPORTED yearly recurrence.");
            }else
                RvLogger.FatalWrite("Trying to add NOT SUPPORTED UNKNOWN recurrence.");

            //return patternString.ToString();
        }

        private static icm.XmlApi.dayOfWeekType convertToDayOfWeekType(webServiceData.DayOfTheWeek whichDay)
        {
            icm.XmlApi.dayOfWeekType result = icm.XmlApi.dayOfWeekType.MON;
            if (webServiceData.DayOfTheWeek.Tuesday == whichDay)
                result = icm.XmlApi.dayOfWeekType.TUE;
            else if (webServiceData.DayOfTheWeek.Wednesday == whichDay)
                result = icm.XmlApi.dayOfWeekType.WED;
            else if (webServiceData.DayOfTheWeek.Thursday == whichDay)
                result = icm.XmlApi.dayOfWeekType.THU;
            else if (webServiceData.DayOfTheWeek.Friday == whichDay)
                result = icm.XmlApi.dayOfWeekType.FRI;
            else if (webServiceData.DayOfTheWeek.Saturday == whichDay)
                result = icm.XmlApi.dayOfWeekType.SAT;
            else if (webServiceData.DayOfTheWeek.Sunday == whichDay)
                result = icm.XmlApi.dayOfWeekType.SUN;

            return result;
        }

        private static icm.XmlApi.dayRecurrenceType convertToDayRecurrenceType(webServiceData.DayOfTheWeek whichDay)
        {
            icm.XmlApi.dayRecurrenceType result = icm.XmlApi.dayRecurrenceType.MON;
            if ( webServiceData.DayOfTheWeek.Day == whichDay )
                result = icm.XmlApi.dayRecurrenceType.ANYDAY;
            else if ( webServiceData.DayOfTheWeek.Weekday == whichDay )
                result = icm.XmlApi.dayRecurrenceType.WEEKDAY;
            else if ( webServiceData.DayOfTheWeek.WeekendDay == whichDay )
                result = icm.XmlApi.dayRecurrenceType.WEEKENDDAY;
            else if (webServiceData.DayOfTheWeek.Tuesday == whichDay)
                result = icm.XmlApi.dayRecurrenceType.TUE;
            else if (webServiceData.DayOfTheWeek.Wednesday == whichDay)
                result = icm.XmlApi.dayRecurrenceType.WED;
            else if (webServiceData.DayOfTheWeek.Thursday == whichDay)
                result = icm.XmlApi.dayRecurrenceType.THU;
            else if (webServiceData.DayOfTheWeek.Friday == whichDay)
                result = icm.XmlApi.dayRecurrenceType.FRI;
            else if (webServiceData.DayOfTheWeek.Saturday == whichDay)
                result = icm.XmlApi.dayRecurrenceType.SAT;
            else if (webServiceData.DayOfTheWeek.Sunday == whichDay)
                result = icm.XmlApi.dayRecurrenceType.SUN;

            return result;
        }

        private static bool CertificateValidationCallBack(
            object sender,
            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
            System.Security.Cryptography.X509Certificates.X509Chain chain,
            System.Net.Security.SslPolicyErrors sslPolicyErrors)
            {
               // If the certificate is a valid, signed certificate, return true.
               if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
               {
                  return true;
               }

               // If there are errors in the certificate chain, look at each error to determine the cause.
               if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
               {
                  if (chain != null && chain.ChainStatus != null)
                  {
                     foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                     {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                           // Self-signed certificates with an untrusted root are valid. 
                           continue;
                        }
                        else
                        {
                           if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                           {
                              // If there are any other errors in the certificate chain, the certificate is invalid,
                              // so the method returns false.
                              return false;
                           }
                        }
                     }
                  }

                  // When processing reaches this line, the only errors in the certificate chain are 
                  // untrusted root errors for self-signed certificates. These certificates are valid
                  // for default Exchange server installations, so return true.
                  return true;
               }
               else
               {
                  // In all other cases, return false.
                  return false;
               }
            }


        private static string GetHashString(SchedulingInfo schedulingInfo, string subject)
        {
            string startDate = schedulingInfo.startDate.ToUniversalTime().ToString();
            string endDate = schedulingInfo.endDate.ToUniversalTime().ToString();
            string location = schedulingInfo.location;
            string resources = schedulingInfo.outlookResources;
            string recipients = schedulingInfo.recipents;
            string str = subject + startDate + endDate + location + resources + recipients;
            return RvScopiaMeeting.getHashString4Str(str);
        }

        private static string getHashString4Str(string str) {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            // Calculate the hash for the files. 
            using (HashAlgorithm hashAlg = HashAlgorithm.Create())
            {
                byte[] hashBytesA = hashAlg.ComputeHash(bytes);
                return BitConverter.ToString(hashBytesA);
            }
        }

        private static bool isChangedByEWS(SchedulingInfo schedulingInfo)
        {
            String lastInfo = schedulingInfo.lastHashInfo;
            String lastRecurrenceHashInfo = schedulingInfo.lastRecurrenceHashInfo;
            using (HashAlgorithm hashAlg = HashAlgorithm.Create())
            {
                // Compare the hashes.
                if (lastInfo.Equals(GetHashString(schedulingInfo, schedulingInfo.subject))
                    && lastRecurrenceHashInfo.Equals(schedulingInfo.recurrenceHashInfo))
                {
                    return true;
                }
            }
            return false;
        }

        public static string parseMeetingPin(string source)
        {
            string meetingPin = "";
            if (source != null && !string.IsNullOrEmpty(source.Trim()))
            {
                source = source.Trim();
                int lastIndexOfPIN = source.ToLower().LastIndexOf("pin");
                if (-1 == lastIndexOfPIN) return meetingPin;
                int indexOfPinEqualSign = source.IndexOf("=", lastIndexOfPIN);
                if (-1 == indexOfPinEqualSign) return meetingPin;
                if (!string.IsNullOrEmpty(source.Substring(lastIndexOfPIN + 3, indexOfPinEqualSign - lastIndexOfPIN - 3).Trim()))
                    return meetingPin;

                if (source.Length >= indexOfPinEqualSign)
                {
                    string endStr = source.Substring(indexOfPinEqualSign + 1).Trim();
                    if (string.IsNullOrEmpty(endStr))
                        return meetingPin;

                    int indexOfFirstSpace = endStr.IndexOf(" ");
                    int indexOfFirstN = endStr.IndexOf("\n");
                    int indexOfFirstRN = endStr.IndexOf("\r\n");

                    indexOfFirstSpace = indexOfFirstSpace == -1 || (indexOfFirstN != -1 && indexOfFirstN < indexOfFirstSpace) ? indexOfFirstN : indexOfFirstSpace;
                    indexOfFirstSpace = indexOfFirstSpace == -1 || (indexOfFirstRN != -1 && indexOfFirstRN < indexOfFirstSpace) ? indexOfFirstRN : indexOfFirstSpace;

                    if (indexOfFirstSpace == -1)
                        meetingPin = endStr;
                    else
                        meetingPin = endStr.Substring(0, indexOfFirstSpace);

                    return meetingPin;
                }
                else
                {
                    return "";
                }
            }
            return "";
        }

    }
   
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
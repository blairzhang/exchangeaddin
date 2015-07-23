using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.TextConverters;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.ContentTypes.iCalendar;
using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.ContentTypes.Tnef;
using Radvision.Scopia.ExchangeMeetingAddIn.icm.XmlApi;
using RvScopiaMeetingAddIn;

namespace Radvision.Scopia.ExchangeMeetingAddIn
{
    public class RvMailParser
    {
        private RvScopiaMeeting rvScopiaMeeting = null;
        public RvMailParser(RvScopiaMeeting rvScopiaMeeting)
        {
            this.rvScopiaMeeting = rvScopiaMeeting;
        }

        public void parseTnefSimple(MailItem mailItem, SchedulingInfo schedulingInfo)
        {
            try
            {

                TnefReader tnefreader = new TnefReader(mailItem.Message.TnefPart.GetContentReadStream(), 0, TnefComplianceMode.Loose);
                EnvelopeRecipientCollection.Enumerator enumerator = mailItem.Recipients.GetEnumerator();
                StringBuilder recipents = new StringBuilder("");
                StringBuilder resources = new StringBuilder("");
                ArrayList displayNames = new ArrayList();
                while (enumerator.MoveNext())
                {
                    string emailAddress = enumerator.Current.Address.ToString();
                    int recipientP2Type = (int)enumerator.Current.Properties["Microsoft.Exchange.Transport.RecipientP2Type"];
                    String displayName = emailAddress;
                    try
                    {
                        displayName = (string)enumerator.Current.Properties["Microsoft.Exchange.MapiDisplayName"];
                    }
                    catch (Exception ex1)
                    {
                        RvLogger.DebugWrite("Error: " + ex1.Message);
                        if(emailAddress != null )
                        {
                               int lastIndex = emailAddress.LastIndexOf("@");
                                if (lastIndex != -1)
                                {
                                        displayName = emailAddress.Substring(0, lastIndex);
                                }
                        }
                    }
                    displayNames.Add(displayName);
                    if (recipientP2Type == 3)
                    {
                        resources.Append(emailAddress).Append(" ");
                    }
                    else if (recipientP2Type == 1 || recipientP2Type == 2)
                    {
                        recipents.Append(emailAddress).Append(" ");
                    }
                }
                schedulingInfo.outlookResources = resources.ToString();
                schedulingInfo.recipents = recipents.ToString();
                displayNames.Sort();
                StringBuilder displayNamesStrb = new StringBuilder("");
                foreach (object displayName in displayNames)
                {
                    displayNamesStrb.Append((string)displayName).Append("; ");
                }
                string displayNamesStr = displayNamesStrb.ToString().Trim();
                schedulingInfo.displayNames = displayNamesStr.Substring(0, displayNamesStr.Length - 1);
                RvLogger.DebugWrite("Recipents emails are :" + schedulingInfo.recipents);
                RvLogger.DebugWrite("Resources emails are :" + schedulingInfo.outlookResources);
                RvLogger.DebugWrite("DisplayNames emails are :" + displayNamesStr);

                bool isRecurrence = false;
                bool isOcurrence = false;
                while (tnefreader.ReadNextAttribute())
                {
                    if (tnefreader.AttributeTag != TnefAttributeTag.MapiProperties)
                        continue;

                    while (tnefreader.PropertyReader.ReadNextProperty())
                    {
                        try
                        {
                            TnefPropertyTag tag = tnefreader.PropertyReader.PropertyTag;
                            //RvLogger.DebugWrite("PropertyTagID:" + tag.Id);
                            //RvLogger.DebugWrite("PropertyTagToString:" + tag.ToString());
                            //RvLogger.DebugWrite("ValueType==:" + tnefreader.PropertyReader.ValueType);
                            /*try
                            {
                                RvLogger.DebugWrite("PropertyID:" + tnefreader.PropertyReader.PropertyNameId.Id);
                                RvLogger.DebugWrite("PropertyName:" + tnefreader.PropertyReader.PropertyNameId.Name);
                                RvLogger.DebugWrite("PropertySetGUID:" + tnefreader.PropertyReader.PropertyNameId.PropertySetGuid);
                            }
                            catch (Exception) {
                                RvLogger.DebugWrite("***********************************************0");
                            }
                            */
                            if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 33302)
                            {
                                byte[] recurrencePatternByte = tnefreader.PropertyReader.ReadValueAsBytes();
                                RecurrencePattern recurrencePattern = RvMailParser.parseRecurrenceMeeting(recurrencePatternByte);
                                schedulingInfo.recurrencePattern = recurrencePattern;
                            }
                            else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 33374)
                            {
                                byte[] timezonebytes = tnefreader.PropertyReader.ReadValueAsBytes();
                                byte[] lengthOfZone = { timezonebytes[6], timezonebytes[7] };
                                int length = BitConverter.ToInt16(lengthOfZone, 0);
                                byte[] timezonekeyNameBytes = new byte[length * 2];
                                Array.Copy(timezonebytes, 8, timezonekeyNameBytes, 0, length * 2 - 1);
                                String timezonekeyName = Encoding.Unicode.GetString(timezonekeyNameBytes);
                                schedulingInfo.timeZoneID = timezonekeyName;
                                //RvLogger.DebugWrite("timezonekeyName:" + timezonekeyName);
                            }
                            if (!tnefreader.PropertyReader.ValueType.IsArray)
                            {
                                object propValue = tnefreader.PropertyReader.ReadValue();

                                //RvLogger.DebugWrite("PropertyValue:" + propValue);

                                if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 33315)
                                {
                                    if ((Boolean)propValue)
                                        isRecurrence = (Boolean)propValue;
                                    isOcurrence = !(Boolean)propValue;
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 33330)
                                {
                                    schedulingInfo.recurrenceInfo = (string)propValue;
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 33332)
                                {
                                    schedulingInfo.timeZoneDisplayName = (string)propValue;
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 33336)
                                {
                                    schedulingInfo.displayNamesFromTnef = ((string)propValue).Trim();
                                }

                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 41)
                                {
                                    schedulingInfo.preStartDate = (DateTime)propValue;
                                    RvLogger.DebugWrite("schedulingInfo.preStartDate===================" + ((DateTime)propValue));
                                    schedulingInfo.hasPreStartDate = true;
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 4115)
                                {
                                    RvLogger.DebugWrite("Parsed html body===================" + ((string)propValue));
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 4118)
                                {
                                    RvLogger.DebugWrite("Parsed native body===================" + ((string)propValue));
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 4096)
                                {
                                    RvLogger.DebugWrite("Parsed text body===================" + ((string)propValue));
                                }
                                else if (tag.Id == TnefPropertyId.StartDate)
                                {
                                    schedulingInfo.startDate = ((DateTime)propValue);
                                    RvLogger.DebugWrite("schedulingInfo.startDate===================" + ((DateTime)propValue));

                                }
                                else if (tag.Id == TnefPropertyId.EndDate)
                                {
                                    schedulingInfo.endDate = ((DateTime)propValue);
                                }
                                else if (tag.Id == TnefPropertyId.LastModificationTime)
                                {
                                    RvLogger.DebugWrite("lastModifiedTime===================" + ((DateTime)propValue));
                                }
                                else if (tag.Id == TnefPropertyId.CreationTime)
                                {
                                    RvLogger.DebugWrite("CreationTime===================" + ((DateTime)propValue));
                                }
                                else if (tag.Id == TnefPropertyId.SenderEmailAddress)
                                {
                                    schedulingInfo.senderEmailAddr = (string)propValue;
                                }
                                else if (tag.Id == TnefPropertyId.SubjectPrefix)
                                {
                                    if (null != propValue)
                                        schedulingInfo.subjectPrefix = ((string)propValue).Trim();
                                }
                                else if (tag.Id == TnefPropertyId.StoreEntryId)
                                {
                                    schedulingInfo.storeEntryId = (string)propValue;
                                }
                                else if (tag.Id == TnefPropertyId.INetMailOverrideCharset)
                                {
                                    object a = propValue;
                                }
                                else if (tag.Id == TnefPropertyId.ConversationKey)
                                {
                                    schedulingInfo.conversationKey = (string)propValue;
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 2)
                                {
                                    schedulingInfo.location = (string)propValue;
                                    string meetingPin = RvScopiaMeeting.parseMeetingPin(schedulingInfo.location);
                                    if (!string.IsNullOrEmpty(meetingPin))
                                        schedulingInfo.meetingPin = meetingPin;
                                }
                                else if (tnefreader.PropertyReader.IsNamedProperty && tnefreader.PropertyReader.PropertyNameId.Id == 33289)
                                {
                                    string value = (string)propValue;
                                    RvLogger.DebugWrite("infomation======================================" + value);
                                    int firstIndex = value.IndexOf(":");
                                    int lastIndex = value.LastIndexOf(":");
                                    if (firstIndex != -1)
                                    {
                                        schedulingInfo.conferenceID = value.Substring(firstIndex + 1, lastIndex - firstIndex - 1);
                                        string hashString = value.Substring(lastIndex + 1);
                                        int indexHashSplit = hashString.IndexOf("$");
                                        schedulingInfo.lastHashInfo = hashString.Substring(0, indexHashSplit);
                                        schedulingInfo.lastRecurrenceHashInfo = hashString.Substring(indexHashSplit + 1);
                                        if (value.StartsWith("#"))
                                        {
                                            isRecurrence = true;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (tag == TnefPropertyTag.RtfCompressed)
                                {
                                    schedulingInfo.emailMessage = "";
                                    Stream stream = tnefreader.PropertyReader.GetRawValueReadStream();
                                    Stream decompressedRtfStream = new ConverterStream(stream, new RtfCompressedToRtf(), ConverterStreamAccess.Read);
                                    RtfToHtml rtfToHtml = new RtfToHtml();
                                    rtfToHtml.OutputEncoding = System.Text.Encoding.UTF8;
                                    rtfToHtml.EnableHtmlDeencapsulation = true;
                                    Stream text = new ConverterStream(decompressedRtfStream, rtfToHtml, ConverterStreamAccess.Read);
                                    StreamReader sr = new StreamReader(text);
                                    schedulingInfo.emailMessage = sr.ReadToEnd();
                                    //RvLogger.DebugWrite("schedulingInfo.emailMessage: " + schedulingInfo.emailMessage);
                                    int indexOfSystemInfoSign = schedulingInfo.emailMessage.IndexOf("*~*~*~*~*~*~*~*~*~*");
                                    int startIndex = -1;
                                    int endIndex = -1;
                                    if (indexOfSystemInfoSign > -1)
                                    {
                                        startIndex = schedulingInfo.emailMessage.Substring(0, indexOfSystemInfoSign).LastIndexOf("<div>");
                                        endIndex = schedulingInfo.emailMessage.IndexOf("</div>", indexOfSystemInfoSign) + 6;
                                        if (startIndex > -1 && endIndex > -1)
                                            schedulingInfo.emailMessage = schedulingInfo.emailMessage.Substring(0, startIndex) + schedulingInfo.emailMessage.Substring(endIndex);
                                    }

                                    schedulingInfo.emailMessage = schedulingInfo.emailMessage.Replace("<img src=\"objattph://\">", "");
                                    sr.Close();
                                    text.Close();
                                    decompressedRtfStream.Close();
                                    stream.Close();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            RvLogger.DebugWrite("ex.Message======================" + ex.Message);
                            RvLogger.DebugWrite(ex.StackTrace);
                        }
                    }
                }

                if (isRecurrence && !isOcurrence)
                    schedulingInfo.meetingType = MeetingType.Reccurence;
                else if (isRecurrence && isOcurrence)
                    schedulingInfo.meetingType = MeetingType.Ocurrence;
                else
                    schedulingInfo.meetingType = MeetingType.Normal;

                tnefreader.Close();
            }
            catch (Exception e)
            {
                RvLogger.DebugWrite("Fail to parse mail-- " + e.Message);
                RvLogger.DebugWrite(e.StackTrace);
            }
        }

        public void changeBodyOfTnef(MimePart tnefPart, SchedulingInfo schedulingInfo)
        {
            TnefReader tnefreader = new TnefReader(tnefPart.GetContentReadStream(), 0, TnefComplianceMode.Loose);
            while (tnefreader.ReadNextAttribute())
            {
                if (tnefreader.AttributeTag != TnefAttributeTag.MapiProperties)
                    continue;
                string dialingInfo = schedulingInfo.dialingInfo;
                TnefWriter writer = new TnefWriter(
                tnefPart.GetContentWriteStream(tnefPart.ContentTransferEncoding), tnefreader.AttachmentKey, 0, TnefWriterFlags.NoStandardAttributes);

                writer.StartAttribute(TnefAttributeTag.MapiProperties, TnefAttributeLevel.Message);
                writer.WriteAllProperties(tnefreader.PropertyReader);
                writer.StartProperty(TnefPropertyTag.RtfCompressed);
                if (null != dialingInfo)
                {
                    dialingInfo = convertToUnicode(dialingInfo);
                }
                string body = null == dialingInfo ? schedulingInfo.emailMessage : schedulingInfo.emailMessage + "<br><br>" + dialingInfo;
                Stream stream = new MemoryStream(Encoding.UTF8.GetBytes(body));
                HtmlToRtf htmlToRtf = new HtmlToRtf();
                RtfToRtfCompressed rtfToRtfCompressed = new RtfToRtfCompressed();
                htmlToRtf.InputEncoding = System.Text.Encoding.UTF8;//GetEncoding("ISO-8859-1");
                stream = new ConverterStream(stream, htmlToRtf, ConverterStreamAccess.Read);
                stream = new ConverterStream(stream, rtfToRtfCompressed, ConverterStreamAccess.Read);
                writer.WritePropertyValue(stream);

                if (null != writer)
                {
                    writer.Close();
                }
            }
            tnefreader.Close();
            RvLogger.DebugWrite("ok**************************************");
        }

        private string convertToUnicode(String source) {
            StringBuilder result = new StringBuilder();
            for (int i = 0, len = source.Length; i < len; i++) {
                char aChar = source[i];
                if ((aChar >= 0x4e00 && aChar <= 0x9fbb) 
                    || (aChar >= 0x2e80 && aChar <= 0x2fdf) 
                    || (aChar >= 0x3400 && aChar <= 0x4dbf)
                    || (aChar >= 0xf900 && aChar <= 0xfaff)
                    || (aChar >= 0x31a0 && aChar <= 0x31bf)
                    || (aChar >= 0x3040 && aChar <= 0x30ff)
                    || (aChar >= 0x31f0 && aChar <= 0x31ff)
                    || (aChar >= 0x1100 && aChar <= 0x11ff)
                    || (aChar >= 0x3130 && aChar <= 0x318f)
                    || (aChar >= 0xac00 && aChar <= 0xd7af)
                    || (aChar >= 0xa960 && aChar <= 0xa97f)
                    || (aChar >= 0xd7b0 && aChar <= 0xd7ff)
                    || (aChar >= 0xff00 && aChar <= 0xffef)
                    || (aChar >= 0x0400 && aChar <= 0x052f)
                    || (aChar >= 0x3000 && aChar <= 0x303f)
                    || (aChar >= 0x2000 && aChar <= 0x206f)
                    )
                {
                    result.Append("&#").Append("" + (int)aChar).Append(";");
                } else
                    result.Append(aChar);
            }

            return result.ToString();
        }

        private string UnicodeString(string text)
        {
            return Encoding.UTF8.GetString(Encoding.ASCII.GetBytes(text));
        }

        private static String byteArrayToString(byte[] array, Encoding econding)
        {
            econding.GetString(array);
            return System.Text.Encoding.UTF8.GetString(array);
        }

        private static RecurrencePattern parseRecurrenceMeeting(byte[] recurrence)
        {

            RecurrencePattern recurrencePattern = new RecurrencePattern();

            int startMinutes = 0;
            int endMinutes = 0;
            for (int i = 40; i < recurrence.Length; i++)
            {
                if (recurrence[i] == 6 && recurrence[i + 1] == 48 && recurrence[i + 2] == 0 && recurrence[i + 3] == 0)
                {
                    string a = System.Convert.ToString(recurrence[i - 1], 16);
                    string b = System.Convert.ToString(recurrence[i - 2], 16);
                    String c = System.Convert.ToString(recurrence[i - 3], 16);
                    String d = System.Convert.ToString(recurrence[i - 4], 16);
                    endMinutes = System.Convert.ToInt32(a + b + c + d, 16);

                    a = System.Convert.ToString(recurrence[i - 5], 16);
                    b = System.Convert.ToString(recurrence[i - 6], 16);
                    c = System.Convert.ToString(recurrence[i - 7], 16);
                    d = System.Convert.ToString(recurrence[i - 8], 16);
                    startMinutes = System.Convert.ToInt32(a + b + c + d, 16);

                    break;
                }
            }
            DateTime startDate = Convert.ToDateTime("1601-1-1");
            startDate = startDate.AddMinutes(startMinutes);
            DateTime endDate = Convert.ToDateTime("1601-1-1");
            endDate = endDate.AddMinutes(endMinutes);


            recurrencePattern.startDate = startDate;
            recurrencePattern.endDate = endDate;

            if (recurrence[4] == 10 && recurrence[5] == 32)
            {
                //its a daily meeting
                recurrencePattern.patternType = PatternType.DAILY;

                string a = System.Convert.ToString(recurrence[15], 16);
                string b = System.Convert.ToString(recurrence[14], 16);
                int c = System.Convert.ToInt32(a + b, 16);
                c = c / 60 / 24;//its recurrence number. for example c=1 means every 1 day

                recurrencePattern.dailyInterval = c;

                if (recurrence[6] == 1)
                {
                    //weekday meeting.
                    recurrencePattern.dailyInterval = -1;//-1 means weekday.                   
                }
                recurrencePattern = parseEndType(recurrencePattern, recurrence);
            }
            else if (recurrence[4] == 11 && recurrence[5] == 32)
            {
                //weekly meeting
                recurrencePattern.patternType = PatternType.WEEKLY;

                string a = System.Convert.ToString(recurrence[15], 16);
                string b = System.Convert.ToString(recurrence[14], 16);
                int c = System.Convert.ToInt32(a + b, 16);//for example, c=1 means every 1 week
                recurrencePattern.weeklyInterval = c;

                if (recurrence[6] == 1)
                {
                    int weeks = recurrence[22];
                    byte[] list = BitConverter.GetBytes(weeks);
                    System.Collections.BitArray arr = new System.Collections.BitArray(list);
                    bool sbit = arr[0];//Sunday
                    bool mbit = arr[1];//Monday
                    bool tubit = arr[2];//Tuesday
                    bool wbit = arr[3];//Wednesday
                    bool thbit = arr[4];//Thursday
                    bool fbit = arr[5];//Friday
                    bool sabit = arr[6];//Saturday

                    List<dayOfWeekType> daysOfWeeklist = new List<dayOfWeekType>();
                    if (sbit)
                    {
                        daysOfWeeklist.Add(dayOfWeekType.SUN);
                    }
                    if (mbit)
                    {
                        daysOfWeeklist.Add(dayOfWeekType.MON);
                    }
                    if (tubit)
                    {
                        daysOfWeeklist.Add(dayOfWeekType.TUE);
                    }
                    if (wbit)
                    {
                        daysOfWeeklist.Add(dayOfWeekType.WED);
                    }
                    if (thbit)
                    {
                        daysOfWeeklist.Add(dayOfWeekType.THU);
                    }
                    if (fbit)
                    {
                        daysOfWeeklist.Add(dayOfWeekType.FRI);
                    }
                    if (sabit)
                    {
                        daysOfWeeklist.Add(dayOfWeekType.SAT);
                    }
                    dayOfWeekType[] daysOfWeek = daysOfWeeklist.ToArray();
                    recurrencePattern.daysOfWeek = daysOfWeek;
                }
                recurrencePattern = parseEndType(recurrencePattern, recurrence);
            }
            else if (recurrence[4] == 12 && recurrence[5] == 32)
            {
                //monthly meeting                

                string a = System.Convert.ToString(recurrence[15], 16);
                string b = System.Convert.ToString(recurrence[14], 16);
                int c = System.Convert.ToInt32(a + b, 10);//for example, c=1 means every 1 month

                recurrencePattern.monthlyInterval = c;

                if (recurrence[6] == 2)// the first radio box
                {
                    recurrencePattern.patternType = PatternType.MONTHLY;

                    int dayth = recurrence[22];//fox example, dayth=2 means the second day every month
                    recurrencePattern.dayOfMonth = dayth;
                }
                else if (recurrence[6] == 3)// the second radio box
                {
                    recurrencePattern.patternType = PatternType.RELATIVE_MONTHLY;

                    int ordernumber = recurrence[26];//the first/second/third/fourth/last weeks

                    switch (ordernumber)
                    {
                        case 1:
                            recurrencePattern.weekOfMonth = WeekOfMonthType.FIRST;
                            break;
                        case 2:
                            recurrencePattern.weekOfMonth = WeekOfMonthType.SECOND;
                            break;
                        case 3:
                            recurrencePattern.weekOfMonth = WeekOfMonthType.THIRD;
                            break;
                        case 4:
                            recurrencePattern.weekOfMonth = WeekOfMonthType.FOURTH;
                            break;
                        case 5:
                            recurrencePattern.weekOfMonth = WeekOfMonthType.LAST;
                            break;
                    }


                    int weekth = recurrence[22];//
                    switch (weekth)
                    {
                        case 1:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.SUN;
                            break;
                        case 2:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.MON;
                            break;
                        case 4:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.TUE;
                            break;
                        case 8:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.WED;
                            break;
                        case 16:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.THU;
                            break;
                        case 64:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.SAT;
                            break;
                        case 127:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.ANYDAY;
                            break;
                        case 62:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.WEEKDAY;
                            break;
                        case 65:
                            recurrencePattern.dayOfWeekOfMonthly = dayRecurrenceType.WEEKENDDAY;
                            break;
                    }
                }
                recurrencePattern = parseEndType(recurrencePattern, recurrence);
            }
            else if (recurrence[4] == 13 && recurrence[5] == 32)
            {
                //yearly meeting
                recurrencePattern.patternType = PatternType.YEARLY;
                recurrencePattern = parseEndType(recurrencePattern, recurrence);
            }

            return recurrencePattern;

        }

        private static RecurrencePattern parseEndType(RecurrencePattern recurrencePattern, byte[] recurrence)
        {
            int n = 4;
            if (recurrence[6] == 0)
            {
                n = 0;
            }
            else if (recurrence[6] == 3)
            {
                n = 8;
            }


            if (recurrence[22 + n] == 35 && recurrence[23 + n] == 32)
            {
                //no end
                recurrencePattern.hasEnd = false;
            }
            if (recurrence[22 + n] == 34 && recurrence[23 + n] == 32)
            {
                //end after n occurence
                string _a = System.Convert.ToString(recurrence[23 + n + 4], 16);
                string _b = System.Convert.ToString(recurrence[23 + n + 3], 16);
                int _c = System.Convert.ToInt32(_a + _b, 16);
                recurrencePattern.hasEnd = true;
                recurrencePattern.numberOfOccurrences = _c;
            }
            if (recurrence[22 + n] == 33 && recurrence[23 + n] == 32)
            {
                //end after date
                string _a = System.Convert.ToString(recurrence[23 + n + 4], 16);
                string _b = System.Convert.ToString(recurrence[23 + n + 3], 16);
                int _c = System.Convert.ToInt32(_a + _b, 16);
                recurrencePattern.hasEnd = true;
                recurrencePattern.numberOfOccurrences = _c;
            }

            return recurrencePattern;
        }

    }  
   

}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
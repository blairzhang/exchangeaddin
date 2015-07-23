using System;
using System.Collections.Generic;
using System.Text;

namespace Radvision.Scopia.ExchangeMeetingAddIn
{
    public class SchedulingInfo
    {
        public string subject { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public string senderEmailAddr { get; set; }
        public string location { get; set; }
        public RequestType requestType { get; set; }
        public string conferenceID { get; set; }
        public string outlookResources { get; set; }
        public string recipents { get; set; }
        public string lastHashInfo { get; set; }
        public string lastRecurrenceHashInfo { get; set; }
        public string recurrenceHashInfo { get; set; }
        public string meetingPin { get; set; }
        public DateTime preStartDate { get; set; }
        public bool hasPreStartDate { get; set; }
        public RecurrencePattern recurrencePattern { get; set; }
        public string delegatorEmailAddr { get; set; }
        public string emailMessage { get; set; }
        public string recurrenceInfo { get; set;}
        public string subjectPrefix { get; set; }
        public MeetingType meetingType { get; set; }
        public string timeZoneDisplayName { get; set; }
        public string timeZoneID { get; set; }
        public bool isRecipentsChanged { get; set; }
        public bool isAddRecipents { get; set; }
        public string displayNames { get; set; }
        public string displayNamesFromTnef { get; set; }
        public string storeEntryId { get; set; }
        public string conversationKey { get; set; }
        public string dialingInfo { get; set; }
        public long bodyLength { get; set; }

        public SchedulingInfo()
        {
            subject = "";
            senderEmailAddr = "";
            location = "";
            requestType = RequestType.CreateMeeting;
            conferenceID = null;
            recipents = "";
            lastHashInfo = "";
            lastRecurrenceHashInfo = "empty";
            recurrenceHashInfo = "empty";
            meetingPin = "";
            hasPreStartDate = false;
            recurrencePattern = null;
            outlookResources = "";
            recurrenceInfo = "";
            subjectPrefix = "";
            isRecipentsChanged = false;
            isAddRecipents = false;
            displayNames = "";
            displayNamesFromTnef = "";
            emailMessage = "";
            dialingInfo = null;
        }
        
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(" subject=").Append(this.subject)
            .Append(",start=").Append(this.startDate)
            .Append(",end=").Append(this.endDate)
            .Append(",preStartDate=").Append(this.preStartDate)
            .Append(",location=").Append(this.location)
            .Append(",RequestType=").Append(requestType)
            .Append(",conferenceID=").Append(conferenceID == null ? "null" : conferenceID)
            .Append(",outlookResources=").Append(outlookResources)
            .Append(", recurrenceHashInfo=").Append(recurrenceHashInfo)
            .Append(", recurrencePattern=").Append(recurrencePattern)
            .Append(", delegatorEmailAddr=").Append(delegatorEmailAddr);
               
            return sb.ToString();
        }

    }
}

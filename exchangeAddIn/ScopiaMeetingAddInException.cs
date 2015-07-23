using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RvScopiaMeetingAddIn
{
    public class ScopiaMeetingAddInException : Exception
    {
        public static string ERROR_MESSAGE_GETUSER_ERROR = "ERROR_MESSAGE_GETUSER_ERROR";
        public static string ERROR_MESSAGE_CONF_NOT_FOUND = "ERROR_MESSAGE_CONF_NOT_FOUND";
        public static string ERROR_MESSAGE_NO_VIRTUALROOM = "ERROR_MESSAGE_NO_VIRTUALROOM";
        public static string ERROR_MESSAGE_USER_NOT_FOUND = "ERROR_MESSAGE_USER_NOT_FOUND";
        public static string ERROR_MESSAGE_HAVE_NO_PERMISSION = "ERROR_MESSAGE_HAVE_NO_PERMISSION";
        public static string ERROR_MESSAGE_CANNOT_CONNECT2SCOPIA = "ERROR_MESSAGE_CANNOT_CONNECT2SCOPIA";
        public static string ERROR_MESSAGE_WRONG_USERNAME_PASSWORD = "ERROR_MESSAGE_WRONG_USERNAME_PASSWORD";
        //public static string ERROR_MESSAGE_CANNOT_CHANGE_NORMAL_OCCURRENCE2SCOPIA = "ERROR_MESSAGE_CANNOT_CHANGE_NORMAL_OCCURRENCE2SCOPIA";
        //public static string REJECT_REPEATING = "Reject repeating.";
        public static string MEETING_TYPE_NOT_AVAILABLE = "ServiceTemplateId";
        public static string ERROR_FAIL_TO_MODITY = "ERROR_FAIL_TO_MODITY";
        public static string ERROR_MESSAGE_INCORRECT_EWS_CONFIGURATION = "ERROR_MESSAGE_INCORRECT_EWS_CONFIGURATION";

        public ScopiaMeetingAddInException() { }
        public ScopiaMeetingAddInException(string message)  
            : base(message) {
        }
        public ScopiaMeetingAddInException(string message, Exception inner)  
            : base(message, inner) {
        }  
    }
}

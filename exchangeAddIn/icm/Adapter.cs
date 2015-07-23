using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Radvision.Scopia.ExchangeMeetingAddIn.icm
{
	public class Adapter
	{
		#region ... Members section...

		public string Organization {
			get { return string.IsNullOrEmpty( this._Organization ) ? null : this._Organization; }
		}

		public string LoginID 		{
			get { return string.IsNullOrEmpty( this._LoginID ) ? null : this._LoginID; }
		}

		#endregion

		#region ... Members section...

		/// <summary>
		/// ICM server connector
		/// </summary>
		private Connector _Connector;

		/// <summary>
		/// Organization name
		/// </summary>
		private string _Organization;

		/// <summary>
		/// ICM LoginID name
		/// </summary>
		private string _LoginID;

		#endregion

		#region ...Constructors & Destructor section...

		private Adapter()
		{
		}

		public Adapter( string url, string loginID )
		{
			if( url == null )
				throw new ArgumentNullException( "url" );

			if( loginID == null )
				throw new ArgumentNullException( "loginID" );

			this._LoginID	= loginID;
			this._Connector	= new Connector( url, true );

			RvLogger.Write( LogModule.ICM, LogLevel.Info, "Created new ICM adapter [Url '{0}'; Credential : default; LoginID {1}", url, this._LoginID );
		}

		public Adapter( string url, string loginID, string organization )
		{
			if( url == null )
				throw new ArgumentNullException( "url" );

			if( loginID == null )
				throw new ArgumentNullException( "loginID" );

			this._Connector		= new Connector( url, true );
			this._LoginID		= loginID;
			this._Organization	= organization;
			RvLogger.Write( LogModule.ICM, LogLevel.Info, "Created new ICM adapter [Url '{0}'; Credential : default; LoginID '{1}', Organization '{2}'", url, this._LoginID, this._Organization );
		}

		public Adapter( string url, NetworkCredential networkCredential )
		{
			if( url == null )
				throw new ArgumentNullException( "url" );

			if( networkCredential == null )
				throw new ArgumentNullException( "networkCredential" );

			this._Connector		= new Connector( url, this.BuildAuthNetworkCredential( networkCredential ) );
			this._LoginID		= networkCredential.UserName;
			this._Organization	= networkCredential.Domain;

			RvLogger.Write( LogModule.ICM, LogLevel.Info, "Created new ICM adapter [Url '{0}'; Credential : Username '{1}'; , Organization {2}",
				url, networkCredential.UserName, networkCredential.Domain );
		}

		public Adapter( string url, NetworkCredential networkCredential, AccountCredential accountCredential )
		{
			if( url == null )
				throw new ArgumentNullException( "url" );

			if( networkCredential == null )
				throw new ArgumentNullException( "networkCredential" );

			this._Connector		= new Connector( url, this.BuildAuthNetworkCredential( networkCredential ), accountCredential );
			this._LoginID		= networkCredential.UserName;
			this._Organization	= networkCredential.Domain;

			RvLogger.Write( LogModule.ICM, LogLevel.Info, "Created new ICM adapter [Url '{0}'; NetworkCredential : Username '{1}', Organization {2}; AccountCredential : Username '{1}'",
				url, networkCredential.UserName, networkCredential.Domain,
				accountCredential.Username);
		}

		private NetworkCredential BuildAuthNetworkCredential( NetworkCredential networkCredential )
		{
			if( networkCredential == null )
				throw new ArgumentNullException();

			NetworkCredential	newNetworkCredential			= new NetworkCredential();
								newNetworkCredential.UserName	= networkCredential.UserName;
								newNetworkCredential.Password	= networkCredential.Password;
								newNetworkCredential.Domain		= networkCredential.Domain;

			if( !string.IsNullOrEmpty( networkCredential.Domain ) )
			//	if( networkCredential.UserName.IndexOf( '@' ) == -1 ) 
					newNetworkCredential.UserName = networkCredential.UserName.Trim() + "@" + networkCredential.Domain.Trim();

			RvLogger.Write( LogModule.ICM, LogLevel.Debug, "BuildAuthNetworkCredential [Original NetworkCredential : Username '{0}', Organization {1};",
				networkCredential.UserName, networkCredential.Domain );

			RvLogger.Write( LogModule.ICM, LogLevel.Debug, "BuildAuthNetworkCredential [Auth NetworkCredential : Username '{0}', Organization {1};",
				newNetworkCredential.UserName, newNetworkCredential.Domain );


			return newNetworkCredential;
		}

		#endregion

		#region ...Requests methods section...

		public XmlApi.userType GetUserInfo(bool needDetail, string userEmail)
		{
			XmlApi.getUserRequestType	request						= new icm.XmlApi.getUserRequestType();
            if (null == userEmail){
				request.Item				= this.Organization;
				request.ItemElementName		= XmlApi.ItemChoiceType.MemberName;
				request.Items				= new object[] { this.LoginID };
				request.ItemsElementName	= new icm.XmlApi.ItemsChoiceType[] { icm.XmlApi.ItemsChoiceType.LoginId };
            } else
                request.UserEmail = userEmail;

            if (needDetail)
            {
                request.DetailedSpecified = true;
                request.Detailed = true;
            }
			XmlApi.getUserResponseType	response					= this._Connector.Request( request ) as XmlApi.getUserResponseType;

			return response != null &&  
						response.User != null && 
							response.User.Length != 0 ?  response.User[ 0 ] : null;
		}

        public XmlApi.getServerInfoResponseType GetAvayaLicenseInfo()
        {
            XmlApi.getServerInfoRequestType request = new icm.XmlApi.getServerInfoRequestType();
            XmlApi.getServerInfoResponseType response = this._Connector.Request(request) as XmlApi.getServerInfoResponseType;
            return response;
        }

		public XmlApi.getVirtualRoomResponseType GetVirtualRooms( icm.XmlApi.userType userInfo )
		{
			if( userInfo == null )
				throw new ArgumentNullException( "userInfo" );

			XmlApi.getVirtualRoomRequestType	request						= new icm.XmlApi.getVirtualRoomRequestType();
												request.MemberId			= userInfo.MemberId;
												request.Item				= userInfo.LoginID;
                                                request.ItemElementName     = icm.XmlApi.ItemChoiceType3.LoginId;

			XmlApi.getVirtualRoomResponseType	response					= this._Connector.Request( request ) as XmlApi.getVirtualRoomResponseType;
            
			//return response != null ? response.VirtualRoom : null;
            return response;
		}

        public XmlApi.getVirtualRoomResponseType searchVirtualRoomsByUserName(string partOfUserName, string memberID)
        {
            XmlApi.getVirtualRoomRequestType request = new icm.XmlApi.getVirtualRoomRequestType();
            request.MemberId = memberID;
            request.Item = partOfUserName;
            request.ItemElementName = icm.XmlApi.ItemChoiceType3.PartOfUserName;

            XmlApi.getVirtualRoomResponseType response = this._Connector.Request(request) as XmlApi.getVirtualRoomResponseType;

            return response;
        }

		public XmlApi.meetingServiceType[] GetMeetingTypes( icm.XmlApi.userType userInfo )
		{
			if( userInfo == null )
				throw new ArgumentNullException( "userInfo" );

			XmlApi.getMeetingServiceRequestType		request						= new icm.XmlApi.getMeetingServiceRequestType();
													request.MemberId			= userInfo.MemberId;
													request.Item				= userInfo.LoginID;
													request.ItemElementName 	= XmlApi.ItemChoiceType2.LoginId;
			XmlApi.getMeetingServiceResponseType	response					= this._Connector.Request( request ) as XmlApi.getMeetingServiceResponseType;

			return response != null ? response.MeetingService : null;
		}

		public XmlApi.organizationType[] GetMembers( icm.XmlApi.userType userInfo )
		{
			if( userInfo == null )
				throw new ArgumentNullException( "userInfo" );

            XmlApi.getOrganizationRequestType request = new icm.XmlApi.getOrganizationRequestType();
											request.Item				= userInfo.MemberId;
											request.ItemElementName 	= XmlApi.ItemChoiceType1.MemberId ;

            XmlApi.getOrganizationResponseType response = this._Connector.Request(request) as XmlApi.getOrganizationResponseType;

			return response != null ? response.Organization : null;
		}

		public XmlApi.terminalType[] GetTerminals( icm.XmlApi.userType userInfo, string name )
		{
			if( userInfo == null )
				throw new ArgumentNullException( "userInfo" );

			XmlApi.getTerminalRequestType	request				= new icm.XmlApi.getTerminalRequestType();
											request.MemberId	= userInfo.MemberId;
                                            request.SortBy = XmlApi.terminalSortByType.NAME;
                                            request.SortBySpecified = true;
                                            request.Ascending = true;
                                            request.AscendingSpecified = true;
                                            if (name == null)
                                                request.Length = "20";
                                            else
                                            {
                                                icm.XmlApi.ItemsChoiceType2[] choice2s = new XmlApi.ItemsChoiceType2[1];
                                                choice2s[0] = XmlApi.ItemsChoiceType2.Name;
                                                string[] items = new String[1];
                                                items[0] = name;
                                                request.ItemsElementName = choice2s;
                                                request.Items = items;
                                                request.Length = "20";
                                            }
                                            
			XmlApi.getTerminalResponseType	response			= this._Connector.Request( request ) as XmlApi.getTerminalResponseType;

			return response != null ? response.Terminal : null;
		}

		public XmlApi.netLocationType[] GetLocations( icm.XmlApi.userType userInfo )
		{
			if( userInfo == null )
				throw new ArgumentNullException( "userInfo" );

			XmlApi.getLocationRequestType	request		= new icm.XmlApi.getLocationRequestType();
			XmlApi.getLocationResponseType	response	= this._Connector.Request( request ) as XmlApi.getLocationResponseType;

			return response != null ? response.Location : null;
		}

		public XmlApi.conferenceType GetConference( string id, DateTime ? time )
		{
            if (string.IsNullOrEmpty(id))
				throw new ArgumentException( "Conference ID cannot be null or empty." );

			XmlApi.getConferenceRequestType		request					= new icm.XmlApi.getConferenceRequestType();
												request.ItemElementName = XmlApi.ItemChoiceType4.ConferenceId;
												request.Item			= id;

			if( time != null )
			{
				request.StartTime			= ( ( DateTime ) time ).ToUniversalTime();
				request.StartTimeSpecified	= true;
			}

			XmlApi.getConferenceResponseType	response	= this._Connector.Request( request ) as XmlApi.getConferenceResponseType;

			return response != null &&
						response.Conference != null &&
							response.Conference.Length != 0 ? response.Conference[ 0 ] : null;
		}

		public XmlApi.scheduleReportType ScheduleConference( XmlApi.conferenceType conference )
		{
			if( conference == null )
				throw new ArgumentNullException( "conference" );
            if (conference.ConferenceId == null)
                conference.ConferenceId = "";
            if( conference.LocationId == null)
                conference.LocationId = "";

            if (conference.TimeZoneId == null) {
                conference.TimeZoneId = TimeZoneInfo.Local.Id;
            }

            conference.clientSpecified = true;
            conference.client = XmlApi.clientType.OUTLOOK_CLIENTLESS;

			XmlApi.scheduleConferenceRequestType	request				= new XmlApi.scheduleConferenceRequestType();
													request.Conference	= conference;

			XmlApi.scheduleConferenceResponseType	response			= this._Connector.Request( request ) as XmlApi.scheduleConferenceResponseType;

			return response.Report;
		}

		public XmlApi.scheduleReportType ModifyConference( XmlApi.conferenceType conference )
		{
			if( conference == null )
				throw new ArgumentNullException( "conference" );

			XmlApi.modifyConferenceRequestType	request				= new XmlApi.modifyConferenceRequestType();
												request.Conference	= conference;

                                                conference.clientSpecified = true;
                                                conference.client = XmlApi.clientType.OUTLOOK_CLIENTLESS;
			XmlApi.modifyConferenceResponseType	response			= this._Connector.Request( request ) as XmlApi.modifyConferenceResponseType;

			return response.Report;
		}

        public XmlApi.scheduleReportType CancelConference(SchedulingInfo schedulingInfo, bool deleteClientlessMapping)
		{
            if (string.IsNullOrEmpty(schedulingInfo.conferenceID))
				throw new ArgumentException( "Conference ID cannot be null or empty." );

			XmlApi.cancelConferenceRequestType	request					= new XmlApi.cancelConferenceRequestType();
            request.ConferenceId = schedulingInfo.conferenceID;
            request.Reccuring = schedulingInfo.meetingType == MeetingType.Reccurence;
            request.ClientSpecified = true;
            request.Client = XmlApi.clientType.OUTLOOK_CLIENTLESS;
            request.DeleteClientLessConfIDMapping = deleteClientlessMapping;

            if (schedulingInfo.meetingType == MeetingType.Ocurrence)
			{
                request.StartTime = ((DateTime)schedulingInfo.startDate).ToUniversalTime();
				request.StartTimeSpecified	= true;
			}

			XmlApi.cancelConferenceResponseType	response				= this._Connector.Request( request ) as XmlApi.cancelConferenceResponseType;

			return response != null ? response.Report : null;
		}

        public XmlApi.dialingInfoType GetDialingInfo(icm.XmlApi.userType userInfo, SchedulingInfo schedulingInfo, icm.XmlApi.virtualRoomType defaultVirtualRoom, bool isCreate, bool isDelete)
        {
            if (userInfo == null)
                throw new ArgumentNullException("userInfo");

            if (userInfo.UserId == null)
                throw new ArgumentNullException("number");

            string accessPIN = schedulingInfo.meetingPin;
            XmlApi.getDialingInfoRequestType request = new XmlApi.getDialingInfoRequestType();
            request.MemberId = userInfo.MemberId;
            request.ItemElementName = XmlApi.ItemChoiceType5.UserId;
            request.Item = userInfo.UserId;
            if (!string.IsNullOrEmpty(accessPIN))
                request.AccessPIN = Encoding.UTF8.GetBytes(accessPIN);
            else
                request.AccessPIN = defaultVirtualRoom.AccessPIN;
            if (null != defaultVirtualRoom.ServicePrefix)
                request.servicePrefix = defaultVirtualRoom.ServicePrefix;
            if (null != defaultVirtualRoom.ServiceTemplateId)
                request.ServiceTemplateId = defaultVirtualRoom.ServiceTemplateId;
            if (isCreate)
            {
                request.operation = XmlApi.operationPolicyType.CREATE;
            }
            else
            {
                request.operation = XmlApi.operationPolicyType.MODIFY;
            }

            if (isDelete)
            {
                request.operation = XmlApi.operationPolicyType.CANCEL;
            }
            request.operationSpecified = true;

            XmlApi.getDialingInfoResponseType response = this._Connector.Request(request) as XmlApi.getDialingInfoResponseType;

            return response != null ? response.DialingInfo : null;
        }

        public XmlApi.userProfileType GetUserProfile(icm.XmlApi.userType userInfo)
        {
            if (userInfo == null)
                throw new ArgumentNullException("userInfo");

            XmlApi.getUserProfileRequestType request = new XmlApi.getUserProfileRequestType();
            request.MemberId = userInfo.MemberId;
            request.ProfileId = new String[1]{userInfo.UserProfileId};

            XmlApi.getUserProfileResponseType response = this._Connector.Request(request) as XmlApi.getUserProfileResponseType;

            return response != null ? (response.UserProfile != null && response.UserProfile.Length > 0 ? response.UserProfile[0] : null) : null;
        }

		#endregion
	}
}


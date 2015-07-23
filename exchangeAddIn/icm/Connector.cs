using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Radvision.Scopia.ExchangeMeetingAddIn.icm
{
	/// <summary>
	/// Provide ICM account information
	/// </summary>
	public class AccountCredential
	{
		#region ...Properties section...

		/// <summary>
		/// Provide ICM account username 
		/// </summary>
		public string Username
		{
			get { return this._Username; }
			set
			{
				if( value == null )
					throw new ArgumentNullException( "Username" );

				this._Username = value;
			}
		}

		/// <summary>
		/// Provide ICM account password
		/// </summary>
		public string Password
		{
			get { return this._Password; }
			set
			{
				if( value == null )
					throw new ArgumentNullException( "Password" );

				this._Password = value;
			}
		}

		#endregion

		#region ...Members section...

		/// <summary>
		/// Provide ICM account username 
		/// </summary>
		private string _Username;

		/// <summary>
		/// Provide ICM account password
		/// </summary>
		private string _Password;

		#endregion

		#region ...Constructor section...

		/// <summary>
		/// Closed constructor for AccountCredential class
		/// </summary>
		private AccountCredential()
		{
		}

		/// <summary>
		/// Constructor for AccountCredential class
		/// </summary>
		/// <param name="username">ICM Account Username</param>
		/// <param name="password">ICM Account Password</param>
		public AccountCredential( string username, string password )
		{
			if( username == null )
				throw new ArgumentNullException( "username" );

			if( password == null )
				throw new ArgumentNullException( "username" );

			this._Username = username;
			this._Password = password;
		}

		#endregion
	}

    public class ConformWriter : XmlTextWriter
    {
        public ConformWriter(Stream w, Encoding encoding):base(w, encoding){}
        public ConformWriter(String fileName, Encoding encoding) : base(fileName, encoding) { }

        internal string CheckUnicodeString(String value)
        {
            StringBuilder xml = new StringBuilder();
            for (int i = 0; i < value.Length; ++i)
            {
                if (value[i] > 0xFFFD)
                {
                    RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "Invalid Unicode - {0}", i);

                }
                else if (value[i] < 0x20 && value[i] != '\t' & value[i] != '\n' & value[i] != '\r')
                {
                    RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "Invalid Xml Characters - {0}", i);

                }
                else
                    xml.Append(value[i]);

            }
            return xml.ToString();
            /*
            for (int i = 0; i < value.Length; ++i)
            {
                if (value[i] > 0xFFFD)
                {
                    RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "Invalid Unicode - {0}", i);
                    throw new Exception("Invalid Unicode");
                }
                else if (value[i] < 0x20 && value[i] != '\t' & value[i] != '\n' & value[i] != '\r')
                {
                    RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "Invalid Xml Characters - {0}", i);
                    throw new Exception("Invalid Xml Characters");
                }else
                    RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "Good Xml Characters - {0}", i);
            }
             * */
        }

       

        public override void WriteString(String value)
        {
            //RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "before - total Xml Characters {0}", value.Length);
            string xml = CheckUnicodeString(value);
            //RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "after - total Xml Characters {0}", xml.Length);
            base.WriteString(xml);
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            base.WriteStartElement(prefix, XmlConvert.EncodeLocalName(localName), ns);
        }
    }
	/// <summary>
	/// Provide connection to ICM server
	/// </summary>
	public class Connector
	{
		#region ...Nested types...

		/// <summary>
		/// Web client class with cookies support.
		/// </summary>
		/// <remarks>By default microsoft WebClient class don't support cookies, we need override GetWebRequest method</remarks>
		public class WebClientWithCookieSupport : WebClient
		{
			/// <summary>
			/// Cookies container
			/// </summary>
			private readonly CookieContainer _CookieContainer = new CookieContainer();

			/// <summary>
			/// Returns a WebRequest object for the specified resource.
			/// </summary>
			/// <param name="address">A Uri that identifies the resource to request. </param>
			/// <returns>A new WebRequest object for the specified resource. </returns>
			protected override WebRequest GetWebRequest( Uri address )
			{
				HttpWebRequest	httpWebRequest	= base.GetWebRequest( address ) as HttpWebRequest;
				if( httpWebRequest != null )
					httpWebRequest.CookieContainer = this._CookieContainer;

                httpWebRequest.ContentType = "text/plain; charset=utf-8";
                httpWebRequest.Timeout = 600000;
				return httpWebRequest;
			}
		}		

		#endregion

		#region ...Properties section...

		/// <summary>
		/// Gets or sets a Boolean value that controls whether the DefaultCredentials are sent with requests. 
		/// </summary>
		public bool UseDefaultCredentials
		{
			get { return this._Client.UseDefaultCredentials;	}
			set { this._Client.UseDefaultCredentials = value;	}
		}

		/// <summary>
		/// Gets or sets the network credentials that are sent to the host and used to authenticate the request.
		/// </summary>
		public NetworkCredential NetworkCredentials
		{
			get { return this._NetworkCredential;	}
			set 
			{
				//
				// ICM don't support organization as domain, if it's will be fixed we can store NetworkCredentials only in WebClient object
				//
				this._NetworkCredential		= value;
				this._Client.Credentials	=
						this._NetworkCredential == null ? null : new NetworkCredential( this._NetworkCredential.UserName, this._NetworkCredential.Password );
			}
		}

		/// <summary>
		/// Gets or sets the account credentials that are sent to the ICM and used to authenticate the request.
		/// </summary>
		public AccountCredential AccountCredential
		{
			get { return this._AccountCredential;	}
			set { this._AccountCredential = value;	}
		}

		/// <summary>
		/// Gets or sets the base URI ICM server addres
		/// </summary>
		public string Address
		{
			get { return this._Client.BaseAddress;	}
			set { this._Client.BaseAddress = value; }
		}

		#endregion

		#region ...Members section...

		/// <summary>
		/// Account credentials that are sent to the ICM and used to authenticate the request
		/// </summary>
		private AccountCredential	_AccountCredential;

		/// <summary>
		/// Network credentials that are sent to the ICM and used to authenticate the request
		/// </summary>
		/// <remarks>ICM don't support organization as domain, if it's will be fixed we can store NetworkCredentials only in WebClient object</remarks>
		private NetworkCredential	_NetworkCredential;

		/// <summary>
		/// Web client provide network opperation
		/// </summary>
		private WebClient			_Client;

		/// <summary>
		/// Provide Serializer/Deserializer services
		/// </summary>
		private XmlSerializer		_XmlSerializer;

		/// <summary>
		/// UTF encoding for XML Serializer
		/// </summary>
		private Encoding			_XmlEncoding;

		#endregion

		#region ...Static members section...

		/// <summary>
		/// Xml API uniq Requests counter
		/// </summary>
		private static int			_RequestID;

        private string languageName;

		#endregion

		#region ...Constructors & Destructor section...

		/// <summary>
		/// Construct new instance of ICM connector
		/// </summary>
		private Connector()
		{
			this._Client		= new WebClientWithCookieSupport();
			this._XmlEncoding	= new UTF8Encoding( false,true );
			this._XmlSerializer = new XmlSerializer( typeof( XmlApi.mcuXmlApiType ) );
            this.languageName   = getAcceptLanguage();
			System.Net.ServicePointManager.ServerCertificateValidationCallback +=
				delegate( object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
											System.Security.Cryptography.X509Certificates.X509Chain chain,
											System.Net.Security.SslPolicyErrors sslPolicyErrors )
				{
					return true; 
				};
		}

		/// <summary>
		/// Construct new instance of ICM connector
		/// </summary>
		/// <param name="address">ICM server address</param>
		public Connector( string address ) : this() {
			this.Address = address;
		}

		/// <summary>
		/// Construct new instance of ICM connector
		/// </summary>
		/// <param name="address">ICM server Url base address</param>
		/// <param name="useDefaultCredentials">Use default network credentials for connect to the ICM server.</param>
		public Connector( string address, bool useDefaultCredentials ) : this( address ) {
			this.UseDefaultCredentials = useDefaultCredentials;
		}

		/// <summary>
		/// Construct new instance of ICM connector
		/// </summary>
		/// <param name="address">ICM server Url base address</param>
		/// <param name="networkCredentials">Network credentials for connect to the ICM server</param>
		public Connector( string address, NetworkCredential networkCredentials ) : this( address ) {
			this.NetworkCredentials = networkCredentials;
		}

		/// <summary>
		/// Construct new instance of ICM connector
		/// </summary>
		/// <param name="address">ICM server Url base address</param>
		/// <param name="networkCredential">Network credentials for connect to the ICM server</param>
		/// <param name="accountCredential">Account credentials for connect to the ICM server</param>
		public Connector( string address, NetworkCredential networkCredential, AccountCredential accountCredential )
			: this( address, networkCredential )
		{
			this.AccountCredential = accountCredential;
		}

		#endregion

		#region ...Public methods section...

		/// <summary>
		/// Send the specified request object to the ICM server and return response.
		/// </summary>
		/// <param name="request">ICM request object</param>
		/// <returns>ICM response object</returns>
		public XmlApi.MCUResponseType Request( XmlApi.MCURequestType request )
		{
			if( request == null )
				throw new ArgumentNullException( "request" );

			//
			// Serelize
			//
			string xmlRequest = this.SerializeMcuXmlApiTypeMessage( this.PrepareMcuXmlApiTypeMessage( request ) );
            /*string simpleRequest = xmlRequest;
            if (request is XmlApi.scheduleConferenceRequestType) { 
                XmlApi.scheduleConferenceRequestType requestType = (XmlApi.scheduleConferenceRequestType) request;
                if(requestType.Conference != null){
                    StringBuilder sb = new StringBuilder("");
                    sb.Append("Schedule_Conference_Request  ")
                        .Append("RequestID: ")
                        .Append(requestType.RequestID)
                        .Append("  ")
                        .Append("Subject: ")
                        .Append(requestType.Conference.Subject);
                    simpleRequest = sb.ToString();
                }
            }*/
            RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "OUTGOING MESSAGE --> ICM\r\n{0}", xmlRequest);

			//
			// Disable Expected 100 Continue, because iView http server don't support it.
			//
			ServicePointManager.Expect100Continue = false;

			//
			// Request
			//
            this._Client.Headers.Add("Accept-Language", getAcceptLanguage());
            this._Client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            string xmlResponse = this._XmlEncoding.GetString(this._Client.UploadData("/xmlservice/entry", "POST", this._XmlEncoding.GetBytes(xmlRequest)));

            if (request is XmlApi.getDialingInfoRequestType)
                RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "INCOMMING MESSAGE <-- ICM\r\n{0}", "GetDialingInfo Response returned.");
            else if (request is XmlApi.getVirtualRoomRequestType)
                RvLogger.Write(LogModule.ICM, LogLevel.Verbose, "INCOMMING MESSAGE <-- ICM\r\n{0}", "GetVirtualRoom Response returned.");
            else
                RvLogger.Write( LogModule.ICM, LogLevel.Verbose, "INCOMMING MESSAGE <-- ICM\r\n{0}", xmlResponse );
			//
			// Deserilize
			//

            object obj = this.DeserializeMcuXmlApiTypeMessage(xmlResponse).Item;

			return ( ( XmlApi.responseType ) this.DeserializeMcuXmlApiTypeMessage( xmlResponse ).Item ).Item;
		}

        private string getAcceptLanguage() { 
            //int languageID = LanguagePack.Instance.LanguageID;
            int languageID = 1033;
            string language = null;
            if (1028 == languageID) {
                language = "zh_TW";
            } else if (1031 == languageID)
            {
                language = "de_DE";
            }
            else if (1033 == languageID)
            {
                language = "en";
            }
            else if (1034 == languageID)
            {
                language = "es_ES";
            }
            else if (1036 == languageID)
            {
                language = "fr_FR";
            }
            else if (1040 == languageID)
            {
                language = "it_IT";
            }
            else if (1041 == languageID)
            {
                language = "ja_JP";
            }
            else if (1042 == languageID)
            {
                language = "ko";
            }
            else if (1046 == languageID)
            {
                language = "pt_BR";
            }
            else if (1049 == languageID)
            {
                language = "ru_RU";
            }
            else if (2052 == languageID)
            {
                language = "zh_CN";
            }
            else if (2070 == languageID)
            {
                language = "pt";
            }
            else {
                language = "en";
            }

            return language;
        }

		#endregion

		#region ...Private methods section...

		/// <summary>
		/// Format XML strings
		/// </summary>
		/// <param name="xmlString">XML string for formating</param>
		/// <returns>Formated XML string</returns>
		private string FromatXML( string xmlString )
		{
			//
			// Create document
			//
			XmlDocument xmlDocument = new XmlDocument();
						xmlDocument.Load(new System.IO.StringReader(xmlString) );

			//
			// Save the document to a file and auto-indent the output.
			//
			MemoryStream	stream				= new MemoryStream();
            XmlTextWriter   writer              = new XmlTextWriter(stream, this._XmlEncoding);
							writer.Formatting	= Formatting.Indented;
							xmlDocument.Save( writer ); 

			return this._XmlEncoding.GetString( stream.ToArray() );
		}

		/// <summary>
		/// Prepare full Xml API Request message
		/// </summary>
		/// <param name="request">ICM request object</param>
		/// <returns>Full Xml API Request message</returns>
		private XmlApi.mcuXmlApiType PrepareMcuXmlApiTypeMessage( XmlApi.MCURequestType request )
		{
			if( request == null )
				throw new ArgumentNullException( "request" );
		
			//
			//
			//
			XmlApi.requestType		requestType			= new XmlApi.requestType();
									requestType.Item	= request;

			XmlApi.mcuXmlApiType	mcuXmlApiType		= new XmlApi.mcuXmlApiType();
									mcuXmlApiType.Item	= requestType;

			//
			// Fill account info
			// 
			if( this._AccountCredential != null )
			{
				mcuXmlApiType.Account	= this._AccountCredential.Username;
				mcuXmlApiType.Password	= this._AccountCredential.Password;
			}

			//
			// Fill Request ID
			//
			request.RequestID = ( ++Connector._RequestID ).ToString();

			//
			//
			//
			return mcuXmlApiType;

		}

		/// <summary>
		/// Serelize mcuXmlApiType object to XML.
		/// </summary>
		/// <param name="mcuXmlApiType">Object for serelize to XML</param>
		/// <returns>XML string</returns>
		
        private string SerializeMcuXmlApiTypeMessage( XmlApi.mcuXmlApiType mcuXmlApiType )
		{
			using( MemoryStream memoryStream = new MemoryStream() )
			{
                using (ConformWriter xmlWriter = new ConformWriter(memoryStream, this._XmlEncoding))
				{
					//xmlWriter.Formatting = Formatting.Indented;
					this._XmlSerializer.Serialize( xmlWriter, mcuXmlApiType );
				}

				return this._XmlEncoding.GetString( memoryStream.ToArray() );
			}
		}
        
		/// <summary>
		/// Serelize XML to mcuXmlApiType.
		/// </summary>
		/// <param name="xmlMessage">XML message</param>
		/// <returns>mcuXmlApiType object</returns>
		private XmlApi.mcuXmlApiType DeserializeMcuXmlApiTypeMessage( string xmlMessage ) {
			return ( XmlApi.mcuXmlApiType ) this._XmlSerializer.Deserialize( new StringReader( xmlMessage ) );
		}

        

		#endregion
	}
}

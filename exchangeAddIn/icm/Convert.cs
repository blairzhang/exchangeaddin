using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Radvision.Scopia.ExchangeMeetingAddIn.icm
{
    class Convert
    {
        public static icm.XmlApi.attendeeType ToAttendee(icm.XmlApi.terminalType terminal)
        {
            icm.XmlApi.attendeeType attendee = new XmlApi.attendeeType();

			attendee.AddressBookEnabled				= terminal.AddressBookEnabled;
			attendee.AddressBookEnabledSpecified	= terminal.AddressBookEnabledSpecified;
			attendee.AreaCode						= terminal.AreaCode;
			attendee.CountryCode                    = terminal.CountryCode;
			attendee.Description					= terminal.Description;
			//attendee.DialIn							= terminal.DialIn;
            //attendee.Email							= terminal.Email;
            //attendee.FirstName						= terminal.FirstName;
            //attendee.LastName						= terminal.LastName;
			attendee.LocationId						= terminal.LocationId;
			attendee.MaxBandwidth					= terminal.MaxBandwidth;
			attendee.MaxBandwidthSpecified			= terminal.MaxBandwidthSpecified;
			attendee.MaxISDNBandwidth				= terminal.MaxISDNBandwidth;
			attendee.MaxISDNBandwidthSpecified		= terminal.MaxISDNBandwidthSpecified;
			attendee.MemberId						= terminal.MemberId;
			attendee.Protocol						= terminal.Protocol;
			attendee.RegisterGKId					= terminal.RegisterGKId;
			attendee.TelephoneNumber				= terminal.TelephoneNumber;
			attendee.Telepresence					= terminal.Telepresence;
			//attendee.TerminalEmail					= terminal.TerminalEmail;
			attendee.TerminalId						= terminal.TerminalId;
			attendee.TerminalName					= terminal.TerminalName;
			attendee.TerminalNumber					= terminal.TerminalNumber;
			attendee.ThreeG							= terminal.ThreeG;
			//attendee.UserId							= terminal.UserId;
			attendee.VideoProfile					= terminal.VideoProfile;
			attendee.VoiceOnly						= terminal.VoiceOnly;

            return attendee;
        }
    }
}

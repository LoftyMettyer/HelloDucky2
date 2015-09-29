using Nexus.Common.Interfaces;
using Nexus.Common.Models;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using Nexus.Common.Enums;
using OpenHRNexus.Common.Enums;
using System.Linq;

namespace Nexus.Common.Classes
{
    public class ProcessEmailTemplate : IProcessStep
    {
        public int Id { get; set; }
        public ProcessElementType Type
        {
            get
            {
                return ProcessElementType.Email;
            }
        }
        public EmailAddressCollection Destinations { get; set; }
        public string Body { get; set; }
        public string Subject { get; set; }
        public List<WebFormButtonModel> FollowOnActions { get; set; }

        public Dictionary<string, object> Variables { get; set; } = new Dictionary<string, object>();

        public ITranslation _translation;

        public void ConvertVariablesToDisplayText()
        {

            //// Lookup values need translating to values for display purposes
            //foreach (var lookupValue in Variables)
            //{

            //}

            //// translate lookup values (replace later to use ProcessVariable struct)
            //// Fettle request type

            var lookupValue = Variables.Where(v => v.Key == "we_22_8").First();

            if (_translation != null)
            {
                var translated = _translation.GetLookupValues(22).Where(v => v.value == (int)lookupValue.Value);
                Variables["we_22_8"] = translated.First().title;
            }
        }

        public MailMessage GenerateMailMessage()
        {
            var buttons = new Dictionary<string, object>();

            foreach (var button in FollowOnActions)
            {
                buttons.Add("button" + button.id.ToString(), button.TargetUrl);
            }

            ConvertVariablesToDisplayText();


            // Do clever stuff with dictionary

            var message = Body.FormatPlaceholder(buttons);

            message = message.FormatPlaceholder(Variables);

            var result = new MailMessage(Destinations.From, Destinations.To)
            {
                Body = message
            };

            return result;

        }

        public ProcessStepStatus Validate()
        {
            return ProcessStepStatus.Success;
        }
    }
}

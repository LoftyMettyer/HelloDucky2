using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Core.Test;

namespace Fusion.Test.SocialCare
{
    public class SocialCareGenericMetadataExtractor : GenericMetadataExtractor
    {
        const string NS = "http://advancedcomputersoftware.com/xml/fusion/socialCare";
        const string VersionNodeName = "version";


        public SocialCareGenericMetadataExtractor(string rootNodeName)
            : base(NS, rootNodeName, VersionNodeName, null, null)
        {

        }
        public SocialCareGenericMetadataExtractor(string rootNodeName, string entityRefNodeName)
            : base(NS, rootNodeName, VersionNodeName, entityRefNodeName, null)
        {

        }

        public SocialCareGenericMetadataExtractor(string rootNodeName, string entityRefNodeName, string primaryRefNodeName)
            : base(NS, rootNodeName, VersionNodeName, entityRefNodeName, primaryRefNodeName)
        {

        }

    }
}

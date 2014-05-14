using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Description;

namespace CrmCommonLib
{
    public class AuthConfiguration
    {

        public String ServerAddress;

        public String OrganizationName;

        public Uri DiscoveryUri;

        public Uri OrganizationUri;

        public Uri HomeRealmUri = null;

        public ClientCredentials DeviceCredentials = null;

        public ClientCredentials Credentials = null;

    }
}

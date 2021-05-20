using System;

namespace mloitzl.sharepoint.functions
{
    public class Settings
    {
        public Uri SiteUrl { get; set; }
        public Guid TenantId { get; set; }
        public Guid ClientId { get; set; }
        public string CertificateThumbprint { get; set; }
    }
}
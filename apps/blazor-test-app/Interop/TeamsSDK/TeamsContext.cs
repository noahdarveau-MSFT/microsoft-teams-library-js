using System.ComponentModel;
using System.Text.Json.Serialization;

namespace Blazor_Test_App.Interop.TeamsSDK;

public class TeamsContext
    {
        public App app { get; set; }
        public Page page { get; set; }
        public User user { get; set; }
        public Sharepointsite sharePointSite { get; set; }

        public class App
        {
            public string locale { get; set; }
            public string sessionId { get; set; }
            public string theme { get; set; }
            public int iconPositionVertical { get; set; }
            public string parentMessageId { get; set; }
            public long userClickTime { get; set; }
            public string userFileOpenPreference { get; set; }
            public Host host { get; set; }
            public string appLaunchId { get; set; }
        }

        public class Host
        {
            public string name { get; set; }
            public string clientType { get; set; }
            public string sessionId { get; set; }
            public string ringId { get; set; }
        }

        public class Page
        {
            public string id { get; set; }
            public string frameContext { get; set; }
            public string subPageId { get; set; }
            public bool isFullScreen { get; set; }
            public bool isMultiWindow { get; set; }
            public string sourceOrigin { get; set; }
        }

        public class User
        {
            public string id { get; set; }
            public string licenseType { get; set; }
            public string loginHint { get; set; }
            public string userPrincipalName { get; set; }
            public Tenant tenant { get; set; }
        }

        public class Tenant
        {
            public string id { get; set; }
            public string teamsSku { get; set; }
        }

        public class Sharepointsite
        {
            public string teamSiteUrl { get; set; }
            public string teamSiteDomain { get; set; }
            public string teamSitePath { get; set; }
            public string mySitePath { get; set; }
            public string mySiteDomain { get; set; }
        }
    }

public enum ChannelType
{
    [Description("Private")] Private,
    [Description("Regular")] Regular
}

public enum FrameContext
{
    [Description("authentication")] Authentication,
    [Description("content")] Content,
    [Description("remove")] Remove,
    [Description("settings")] Settings,
    [Description("sidePanel")] SidePanel,
    [Description("stage")] Stage,
    [Description("task")] Task
}

public enum HostClientType
{
    [Description("android")] Android,
    [Description("desktop")] Desktop,
    [Description("ios")] iOS,
    [Description("rigel")] Rigel,
    [Description("web")] Web
}

public enum Platform
{
    [Description("windows")] Windows,
    [Description("macos")] macOS
}

public enum TeamType
{
    Standard = 0,
    Edu = 1,
    Class = 2,
    Plc = 3,
    Staff = 4
}

public enum UserTeamRole
{
    Admin = 0,
    User = 1,
    Guest = 2
}

public class LocaleInfo
{
    public string LongDate { get; set; }
    public string LongTime { get; set; }
    [JsonConverter(typeof(EnumDescriptionConverter<Platform>))]
    public Platform Platform { get; set; }
    public string RegionalFormat { get; set; }
    public string ShortDate { get; set; }
    public string ShortTime { get; set; }
}

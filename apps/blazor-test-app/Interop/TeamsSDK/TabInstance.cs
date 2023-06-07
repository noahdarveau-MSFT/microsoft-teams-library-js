namespace Blazor_Test_App.Interop.TeamsSDK;

public class TabInstance {

  public string channelId { get; set; }

  public bool? ChannelIsFavorite { get; set; }

  public string ChannelName { get; set; }

  public string EntityId { get; set; }

  public string GroupId { get; set; }

  public string LastViewUnixEpochTime { get; set; }

  public string tabName { get; set; }

  public string TeamId { get; set; }

  public bool? TeamIsFavorite { get; set; }

  public string TeamName { get; set; }

  public string Url { get; set; }

  public string WebsiteUrl { get; set; }

};

public class TabInstanceParameters{

  public TabInstanceParameters(bool Channels, bool Teams) {
    this.FavoriteChannelsOnly = Channels;
    this.FavoriteTeamsOnly = Teams;
  }

  public bool? FavoriteChannelsOnly { get; set; }

  public bool? FavoriteTeamsOnly { get; set; }

}

public class TabInstanceArray {

  public TabInstance[] teamTabs { get; set; }

}
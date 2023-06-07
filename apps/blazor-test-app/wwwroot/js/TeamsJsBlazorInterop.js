export function initializeAsync() {
  return microsoftTeams.app.initialize();
}

export function getContextAsync() {
  const ret = microsoftTeams.app.getContext();
  console.log(ret);
  return ret;
}

export function setCurrentFrame(contentUrl, websiteUrl) {
  microsoftTeams.pages.setCurrentFrame(contentUrl, websiteUrl);
}

export function registerFullScreenHandler() {
  return microsoftTeams.pages.registerFullScreenHandler();
}

export function registerChangeConfigHandler() {
  microsoftTeams.pages.config.registerChangeConfigHandler();
}

export function getTabInstances(tabInstanceParameters) {
  const ret = microsoftTeams.pages.tabs.getTabInstances(tabInstanceParameters);
  console.log(ret);
  return ret;
}

export function getMruTabInstances(tabInstanceParameters) {
  return microsoftTeams.pages.tabs.getMruTabInstances(tabInstanceParameters);
}

export function shareDeepLink(deepLinkParameters) {
  microsoftTeams.pages.shareDeepLink(deepLinkParameters);
}

export function openLink(deepLink) {
  return microsoftTeams.app.openLink(deepLink);
}

export function navigateToTab(tabInstance) {
  return microsoftTeams.pages.tabs.navigateToTab(tabInstance);
}

// Settings module
export function registerOnSaveHandler(settings) {
  microsoftTeams.pages.config.registerOnSaveHandler((saveEvent) => {
    microsoftTeams.pages.config.setConfig(settings);
    saveEvent.notifySuccess();
  });

  microsoftTeams.pages.config.setValidityState(true);
}

export function isHostedInM365() {
  if (window.parent[0]) {
    return true;
  }
  return false;
}

export function notifySuccess() {
  microsoftTeams.app.notifySuccess();
}

using Microsoft.JSInterop;

namespace Blazor_Test_App.Interop.TeamsSDK;

public class MicrosoftTeams : InteropModuleBase
{
    protected override string ModulePath => "./js/TeamsJsBlazorInterop.js";

    public MicrosoftTeams(IJSRuntime jsRuntime) : base(jsRuntime) { }

    public Task InitializeAsync()
    {
        return InvokeVoidAsync("initializeAsync");
    }

    public Task<TabInstanceArray> GetTabInstances(TabInstanceParameters tabInstanceParameters) {
        return InvokeAsync<TabInstanceArray>("getTabInstances", tabInstanceParameters);
    }

    public Task NavigateToTab(TabInstance tabInstance) {
        return InvokeVoidAsync("navigateToTab", tabInstance);
    }

    public Task NavigateToPage(string pageId) {
        return InvokeVoidAsync("navigateToPage", pageId);
    }

    public Task<TabInstanceArray> GetMRUTabInstances(TabInstanceParameters tabInstanceParameters) {
        return InvokeAsync<TabInstanceArray>("getMruTabInstances", tabInstanceParameters);
    }

    public Task<TeamsContext> GetTeamsContextAsync()
    {
        return InvokeAsync<TeamsContext>("getContextAsync");
    }

    public Task RegisterOnSaveHandlerAsync(TeamsInstanceSettings settings)
    {
        return InvokeVoidAsync("registerOnSaveHandler", settings);
    }

    public Task<bool> IsHostedInM365()
    {
        try
        {
            return InvokeAsync<bool>("isHostedInM365");
        }
        catch (JSException)
        {
            return Task.FromResult(false);
        }
    }

    public Task notifySuccess() 
    {
        return InvokeVoidAsync("notifySuccess");
    }
}

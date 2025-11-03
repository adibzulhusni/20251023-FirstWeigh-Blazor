using Microsoft.JSInterop;
using System.Text.Json;

namespace FirstWeigh.Services
{
    public class BrowserStorageService
    {
        private readonly IJSRuntime _jsRuntime;

        public BrowserStorageService(IJSRuntime jsRuntime)
        {
            _jsRuntime = jsRuntime;
        }

        public async Task<T?> GetItemAsync<T>(string key)
        {
            var json = await _jsRuntime.InvokeAsync<string>("browserStorage.getItem", key);

            if (string.IsNullOrEmpty(json))
                return default;

            return JsonSerializer.Deserialize<T>(json);
        }

        public async Task SetItemAsync<T>(string key, T value)
        {
            var json = JsonSerializer.Serialize(value);
            await _jsRuntime.InvokeVoidAsync("browserStorage.setItem", key, json);
        }

        public async Task RemoveItemAsync(string key)
        {
            await _jsRuntime.InvokeVoidAsync("browserStorage.removeItem", key);
        }

        public async Task StartActivityTrackingAsync()
        {
            await _jsRuntime.InvokeVoidAsync("browserStorage.startActivityTracking");
        }

        public async Task<DateTime?> GetLastActivityTimeAsync()
        {
            var timeString = await _jsRuntime.InvokeAsync<string>("browserStorage.getLastActivityTime");

            if (string.IsNullOrEmpty(timeString))
                return null;

            if (DateTime.TryParse(timeString, out DateTime result))
                return result;

            return null;
        }
    }
}
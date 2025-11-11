using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Presentation;
using FirstWeigh.Models;
using Microsoft.AspNetCore.Components;
using System.Timers;

namespace FirstWeigh.Services
{
    public class AuthenticationService : IDisposable
    {
        private readonly UserService _userService;
        private readonly BrowserStorageService _storageService;
        private CurrentUserInfo? _currentUser;
        private System.Timers.Timer? _sessionTimer;
        private const string SESSION_KEY = "FirstWeigh_Session";
        private const int SESSION_TIMEOUT_MINUTES = 1;
        private readonly LoginAttemptService _loginAttemptService;
        private readonly AuditLogService _auditLogService;
        private readonly NavigationManager _navigationManager;
        public event Action<CurrentUserInfo?>? OnAuthenticationStateChanged;

        public CurrentUserInfo? CurrentUser => _currentUser;
        public bool IsAuthenticated => _currentUser != null;



        public AuthenticationService(UserService userService, BrowserStorageService storageService,LoginAttemptService loginAttemptService, AuditLogService auditLogService, NavigationManager navigationManager)
        {
            _userService = userService;
            _storageService = storageService;
            _loginAttemptService = loginAttemptService;
            _auditLogService = auditLogService;
            _navigationManager = navigationManager;  // ADD THIS
        }

        public async Task<bool> LoginAsync(string username, string password, bool rememberMe = false)
        {
            var user = await _userService.ValidateLoginAsync(username, password);

            if (user != null)
            {
                _currentUser = new CurrentUserInfo(user.Username, user.Role);

                var sessionData = new SessionData
                {
                    Username = _currentUser.Username,
                    Role = _currentUser.Role,
                    LoginTime = _currentUser.LoginTime,
                    RememberMe = rememberMe,
                    LastActivityTime = DateTime.Now
                };

                await _storageService.SetItemAsync(SESSION_KEY, sessionData);
                await _storageService.StartActivityTrackingAsync();
                StartSessionTimer();

                OnAuthenticationStateChanged?.Invoke(_currentUser);
                return true;
            }

            return false;
        }
        public async Task LogoutAsync()
        {
            var username = _currentUser?.Username ?? "Unknown";
            _currentUser = null;

            // Stop timer
            StopSessionTimer();

            // Clear session from storage
            await _storageService.RemoveItemAsync(SESSION_KEY);

            OnAuthenticationStateChanged?.Invoke(null);

            // Force page reload to login - ADD THIS
            _navigationManager.NavigateTo("/login", forceLoad: true);
        }

        public async Task<bool> RestoreSessionAsync()
        {
            var sessionData = await _storageService.GetItemAsync<SessionData>(SESSION_KEY);

            if (sessionData == null)
                return false;

            // Check if session has expired
            var lastActivity = sessionData.LastActivityTime;
            var timeSinceActivity = DateTime.Now - lastActivity;

            if (timeSinceActivity.TotalMinutes > SESSION_TIMEOUT_MINUTES)
            {
                // Session expired, clear it
                await _storageService.RemoveItemAsync(SESSION_KEY);
                return false;
            }

            // Restore session
            _currentUser = new CurrentUserInfo(sessionData.Username, sessionData.Role)
            {
                LoginTime = sessionData.LoginTime
            };

            // Update last activity time
            sessionData.LastActivityTime = DateTime.Now;
            await _storageService.SetItemAsync(SESSION_KEY, sessionData);

            // Start activity tracking
            await _storageService.StartActivityTrackingAsync();

            // Start session timer
            StartSessionTimer();

            OnAuthenticationStateChanged?.Invoke(_currentUser);
            return true;
        }

        private void StartSessionTimer()
        {
            StopSessionTimer();

            _sessionTimer = new System.Timers.Timer(60000); // Check every 1 minute
            _sessionTimer.Elapsed += OnSessionTimerElapsed;
            _sessionTimer.AutoReset = true;
            _sessionTimer.Start();
        }

        private void StopSessionTimer()
        {
            if (_sessionTimer != null)
            {
                _sessionTimer.Stop();
                _sessionTimer.Dispose();
                _sessionTimer = null;
            }
        }

        private async void OnSessionTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            try
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⏰ Session timer tick");

                var lastActivityTime = await _storageService.GetLastActivityTimeAsync();

                if (lastActivityTime.HasValue)
                {
                    var timeSinceActivity = DateTime.Now - lastActivityTime.Value;

                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Last activity: {lastActivityTime.Value:g}");
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Time since activity: {timeSinceActivity.TotalMinutes:F2} minutes");

                    if (timeSinceActivity.TotalMinutes > SESSION_TIMEOUT_MINUTES)
                    {
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ⚠️ SESSION TIMEOUT - Logging out!");

                        // Stop timer first to prevent multiple logout calls
                        StopSessionTimer();

                        // Logout on UI thread
                        await InvokeAsync(async () => await LogoutAsync());
                    }
                    else
                    {
                        // Update session data with current activity time
                        var sessionData = await _storageService.GetItemAsync<SessionData>(SESSION_KEY);
                        if (sessionData != null)
                        {
                            sessionData.LastActivityTime = lastActivityTime.Value;
                            await _storageService.SetItemAsync(SESSION_KEY, sessionData);

                            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ✅ Session updated - {SESSION_TIMEOUT_MINUTES - timeSinceActivity.TotalMinutes:F1} minutes remaining");
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ❌ No activity time found");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ❌ Timer error: {ex.Message}");
            }
        }

        // Add this helper method for UI thread invocation
        private Task InvokeAsync(Func<Task> action)
        {
            return Task.Run(action);
        }


        public UserPermissions GetCurrentUserPermissions()
        {
            if (_currentUser == null)
                return new UserPermissions();

            return UserPermissions.GetPermissionsForRole(_currentUser.Role);
        }

        public void Dispose()
        {
            StopSessionTimer();
        }

        // Internal class for session storage
        private class SessionData
        {
            public string Username { get; set; } = string.Empty;
            public string Role { get; set; } = string.Empty;
            public DateTime LoginTime { get; set; }
            public bool RememberMe { get; set; }
            public DateTime LastActivityTime { get; set; }
        }
    }
}
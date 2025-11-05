using FirstWeigh.Models;
using FirstWeigh.Services;
using Microsoft.AspNetCore.Components;
using static FirstWeigh.Models.UserPermissions;

namespace FirstWeigh.Components
{
    public class AuthorizedPageBase : ComponentBase, IDisposable
    {
        [Inject]
        protected AuthenticationService AuthService { get; set; } = default!;

        [Inject]
        protected NavigationManager Navigation { get; set; } = default!;

        protected CurrentUserInfo? CurrentUser => AuthService.CurrentUser;

        // Pages override this to specify which roles can access them
        protected virtual string[] AllowedRoles => Array.Empty<string>();

        protected bool IsAuthorized { get; private set; } = true;

        protected override void OnInitialized()
        {
            // Subscribe to authentication state changes
            AuthService.OnAuthenticationStateChanged += HandleAuthenticationStateChanged;

            // Check authorization
            CheckAuthorization();
        }

        private void CheckAuthorization()
        {
            // If no user is logged in, redirect to login
            if (CurrentUser == null)
            {
                Navigation.NavigateTo("/login");
                return;
            }

            // If page requires specific roles, check if user has one of them
            if (AllowedRoles.Length > 0)
            {
                IsAuthorized = AllowedRoles.Contains(CurrentUser.Role);
            }
            else
            {
                // No specific roles required, any authenticated user can access
                IsAuthorized = true;
            }
        }

        private void HandleAuthenticationStateChanged(CurrentUserInfo? user)
        {
            // If user logged out, redirect to login
            if (user == null)
            {
                Navigation.NavigateTo("/login");
            }
            else
            {
                // Re-check authorization with new user
                CheckAuthorization();
                StateHasChanged();
            }
        }

        public void Dispose()
        {
            AuthService.OnAuthenticationStateChanged -= HandleAuthenticationStateChanged;
        }
    }
}
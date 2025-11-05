namespace FirstWeigh.Models
{
    public class User
    {
        public string UserId { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Role { get; set; } = string.Empty;
        public bool IsActive { get; set; } = true;
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime LastModifiedDate { get; set; } = DateTime.Now;
        public string LastModifiedBy { get; set; } = string.Empty;
        public DateTime? LastLoginDate { get; set; }
    }

    public static class UserRoles
    {
        public const string Admin = "Admin";
        public const string Developer = "Developer";
        public const string Supervisor = "Supervisor";
        public const string Operator = "Operator";

        public static List<string> GetAllRoles()
        {
            return new List<string> { Admin, Developer, Supervisor, Operator };
        }
    }

    public class UserPermissions
    {
        public bool CanViewRecipes { get; set; }
        public bool CanEditRecipes { get; set; }
        public bool CanPerformWeighing { get; set; }
        public bool CanApproveAborts { get; set; }
        public bool CanManageUsers { get; set; }
        public bool CanManageBowls { get; set; }
        public bool CanViewBatchHistory { get; set; }
        public bool CanCreateBatches { get; set; }
        public bool CanManageIngredients { get; set; }

        public static UserPermissions GetPermissionsForRole(string role)
        {
            return role switch
            {
                UserRoles.Admin => new UserPermissions
                {
                    CanViewRecipes = true,
                    CanEditRecipes = true,
                    CanPerformWeighing = true,
                    CanApproveAborts = true,
                    CanManageUsers = true,
                    CanManageBowls = true,
                    CanViewBatchHistory = true,
                    CanCreateBatches = true,
                    CanManageIngredients = true
                },
                UserRoles.Developer => new UserPermissions
                {
                    CanViewRecipes = true,
                    CanEditRecipes = true,
                    CanPerformWeighing = true,
                    CanApproveAborts = true,
                    CanManageUsers = true,
                    CanManageBowls = true,
                    CanViewBatchHistory = true,
                    CanCreateBatches = true,
                    CanManageIngredients = true
                },
                UserRoles.Supervisor => new UserPermissions
                {
                    CanViewRecipes = false,
                    CanEditRecipes = false,
                    CanPerformWeighing = true,
                    CanApproveAborts = true,
                    CanManageUsers = false,
                    CanManageBowls = true,
                    CanViewBatchHistory = true,
                    CanCreateBatches = true,
                    CanManageIngredients = false
                },
                UserRoles.Operator => new UserPermissions
                {
                    CanViewRecipes = false,
                    CanEditRecipes = false,
                    CanPerformWeighing = true,
                    CanApproveAborts = false,
                    CanManageUsers = false,
                    CanManageBowls = true,
                    CanViewBatchHistory = true,
                    CanCreateBatches = false,
                    CanManageIngredients = false
                },
                _ => new UserPermissions()
            };
        }
        // Add this class at the end of your User.cs file

        /// <summary>
        /// Lightweight class to hold current authenticated user information
        /// </summary>
       
    }
    public class CurrentUserInfo
    {
        public string Username { get; set; } = string.Empty;
        public string Role { get; set; } = string.Empty;
        public DateTime LoginTime { get; set; } = DateTime.Now;

        public CurrentUserInfo() { }

        public CurrentUserInfo(string username, string role)
        {
            Username = username;
            Role = role;
            LoginTime = DateTime.Now;
        }
    }
    public class LoginAttempt
    {
        public string Username { get; set; } = string.Empty;
        public string IpAddress { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string Reason { get; set; } = string.Empty;
    }

    public class AuditLog
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string Username { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public string Details { get; set; } = string.Empty;
        public string IpAddress { get; set; } = string.Empty;
    }
}


using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.Owin;
using Owin;
using System;
using WebReport.Models;

[assembly: OwinStartupAttribute(typeof(WebReport.Startup))]
namespace WebReport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
            CreateRolesandUsers();
        }


        // In this method we will create default User roles and Admin user for login
        private void CreateRolesandUsers()
        {
            ApplicationDbContext context = new ApplicationDbContext();
            try
            {
                var roleManager = new RoleManager<IdentityRole>(new RoleStore<IdentityRole>(context));
                var UserManager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(context));




                // In Startup iam creating first Admin Role and creating a default Admin User 
                if (!roleManager.RoleExists("Admin"))
                {

                    // first we create Admin rool
                    var role = new Microsoft.AspNet.Identity.EntityFramework.IdentityRole();
                    role.Name = "Admin";
                    roleManager.Create(role);

                    //Here we create a Admin super user who will maintain the website				

                    var user = new ApplicationUser();
                    user.UserName = "quantri";
                    user.Email = "quantri@gmail.com";
                    string userPWD = "quantri";

                    var chkUser = UserManager.Create(user, userPWD);
                    //ADmin 2
                    var user1 = new ApplicationUser();
                    user1.UserName = "quantri1";
                    user1.Email = "quantri1@gmail.com";
                    string userPWD1 = "quantri1";

                    var chkUser1 = UserManager.Create(user1, userPWD1);
                    //Add default User to Role Admin
                    if (chkUser.Succeeded)
                    {
                        var result1 = UserManager.AddToRole(user.Id, "Admin");

                    }
                    if (chkUser1.Succeeded)
                    {
                        var result1 = UserManager.AddToRole(user1.Id, "Admin");

                    }
                }
            }
            catch (Exception ex) { }
        }


    }
}

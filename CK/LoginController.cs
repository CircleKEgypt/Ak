//using CK.Model;
//using CK.Models;
//using Microsoft.AspNetCore.Authentication;
//using Microsoft.AspNetCore.Authentication.Cookies;
//using Microsoft.AspNetCore.Authorization;
//using Microsoft.AspNetCore.Identity;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.AspNetCore.OutputCaching;
//using Microsoft.EntityFrameworkCore;
//using System.Diagnostics;
//using System.Security.Claims;

//namespace CK.Controllers
//{
//    public class LoginController : Controller
//    {
//        private readonly ILogger<LoginController> _logger;
//        private readonly CkhelperdbContext _dbContext;
//        private static readonly List<Helpersuser> Users = new List<Helpersuser>();
//        public LoginController(ILogger<LoginController> logger, CkhelperdbContext dbContext)
//        {
//            _logger = logger;
//            _dbContext = dbContext;
//        }
//        [HttpGet]
//        [ResponseCache(Location = ResponseCacheLocation.None, NoStore = true)]
//        public IActionResult Login()
//        {
//            ClaimsPrincipal claimUser = HttpContext.User;
//            if (claimUser.Identity.IsAuthenticated)
//            {
//                // Clear existing session when displaying the login page
//                HttpContext.Session.Clear();

//                return RedirectToAction("Index", "Home");
//            }

//            return View();
//        }
//        //[HttpPost]
//        //[ResponseCache(Location = ResponseCacheLocation.None, NoStore = true)]
//        //public async Task<IActionResult> Login(VMLogin modellogin)
//        //{
//        //    if (HttpContext.Request.Query.ContainsKey("preventBack"))
//        //    {
//        //        // Clear authentication cookies if any (extra measure)
//        //        await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

//        //        // Clear session on login
//        //        HttpContext.Session.Clear();
//        //    }

//        //    if (string.IsNullOrWhiteSpace(modellogin.username) || string.IsNullOrWhiteSpace(modellogin.Password))
//        //    {
//        //        TempData["ValidateMessage"] = "Username and password are required.";
//        //        return View("Login");
//        //    }

//        //    var authenticatedUser = _dbContext.Helpersusers
//        //        .FirstOrDefault(u => u.Username == modellogin.username && u.Password == modellogin.Password);

//        //    if (authenticatedUser != null)
//        //    {
//        //        List<Claim> claims = new List<Claim>
//        //{
//        //    new Claim(ClaimTypes.NameIdentifier, authenticatedUser.Username),
//        //    new Claim("OtherProperties", "example role")
//        //};

//        //        ClaimsIdentity claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
//        //        AuthenticationProperties properties = new AuthenticationProperties
//        //        {
//        //            AllowRefresh = true,
//        //        };

//        //        await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), properties);

//        //        // Store username in session for future use
//        //        HttpContext.Session.SetString("Username", authenticatedUser.Username);

//        //        return RedirectToAction("Index", "Home");
//        //    }

//        //    if (HttpContext.Session.GetString("LoggedOut") == "true")
//        //    {
//        //        TempData["PreventBack"] = true;
//        //        HttpContext.Session.SetString("LoggedOut", "false"); // Reset the session variable
//        //    }

//        //    return View();
//        //}
//        [HttpPost]
//        [ResponseCache(Location = ResponseCacheLocation.None, NoStore = true)]
//        public async Task<IActionResult> Login(VMLogin modellogin)
//        {
//            if (HttpContext.Request.Query.ContainsKey("preventBack"))
//            {
//                // Clear authentication cookies if any (extra measure)
//                await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

//                // Clear session on login
//                HttpContext.Session.Clear();
//            }

//            if (string.IsNullOrWhiteSpace(modellogin.username) || string.IsNullOrWhiteSpace(modellogin.Password))
//            {
//                TempData["ValidateMessage"] = "Username and password are required.";
//                return View("Login");
//            }

//            var authenticatedUser = _dbContext.Helpersusers
//                .FirstOrDefault(u => u.Username == modellogin.username && u.Password == modellogin.Password);

//            if (authenticatedUser != null)
//            {
//                // Create claims for the authenticated user
//                List<Claim> claims = new List<Claim>
//            {
//                new Claim(ClaimTypes.NameIdentifier, authenticatedUser.Username),
//                new Claim("OtherProperties", "example role")
//            };

//                ClaimsIdentity claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
//                AuthenticationProperties properties = new AuthenticationProperties
//                {
//                    AllowRefresh = true,
//                };

//                // Sign in the user
//                await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), properties);

//                // Store username in session for future use
//                HttpContext.Session.SetString("Username", authenticatedUser.Username);

//                return RedirectToAction("Index", "Home");
//            }

//            if (HttpContext.Session.GetString("LoggedOut") == "true")
//            {
//                TempData["PreventBack"] = true;
//                HttpContext.Session.SetString("LoggedOut", "false"); // Reset the session variable
//            }

//            return View();
//        }
//        [HttpGet]
//        public IActionResult CreateUser()
//        {
//            return View();
//        }
//        [HttpPost]
//        public IActionResult CreateUser(VMLogin model)
//        {
//            if (!ModelState.IsValid)
//            {
//                return View(model);
//            }

//            // Check if the username already exists
//            if (_dbContext.Helpersusers.Any(u => u.Username == model.username))
//            {
//                ModelState.AddModelError("username", "Username already exists.");
//                return View(model);
//            }

//            // Create new user
//            var newUser = new Helpersuser
//            {
//                Username = model.username,
//                Password = model.Password, // You should hash the password before saving it
//                Department = model.Department // Add Department value to the new user
//            };

//            _dbContext.Helpersusers.Add(newUser);
//            _dbContext.SaveChanges();

//            // Redirect to login page after user creation
//            return RedirectToAction("Login");
//        }
//        public IActionResult DisplayUsers()
//        {
//            var users = _dbContext.Helpersusers.ToList();
//            return View(users);
//        }

//        [HttpGet]
//        public async Task<IActionResult> EditUser(int? id)
//        {
//            //var user = _dbContext.Helpersusers.FirstOrDefault(u => u.Id == id);
//            //if (user == null)
//            //{
//            //    return NotFound();
//            //}
//            var user= await _dbContext.Helpersusers.FindAsync(id);
//            return View(user);
//        }

//        [HttpPost]
//        [ValidateAntiForgeryToken]
//        public async Task<IActionResult> EditUser(int id, [Bind ("Id,Username,Password,Department")] Helpersuser model)
//        {
//            if (!ModelState.IsValid)
//            {
//                return View(model);
//            }

//            //var userToUpdate = _dbContext.Helpersusers.FirstOrDefault(u => u.Id == id);

//            // Check if the userToUpdate is null
//           // if (userToUpdate != null)
//            //{
//                _dbContext.Update(model);
//                await _dbContext.SaveChangesAsync();
//                // Create a new user with the new username
//                //var newUser = new Helpersuser
//                //{
//                //    Username = model.Username,
//                //    Password = model.Password,
//                //    Department = model.Department
//                //};

//                //// Add the new user to the database context
//                //_dbContext.Helpersusers.Add(newUser);
//            //}
//            //else
//            //{
//            //    // Update the existing user with the provided data
//            //    if (model.Id>0)
//            //    {
//            //        // Check if the new username already exists in the database
//            //        var existingUserWithNewUsername = _dbContext.Helpersusers.FirstOrDefault(u => u.Id == model.Id);
//            //        if (existingUserWithNewUsername != null)
//            //        {
//            //            ModelState.AddModelError("Username", "The new username already exists. Please choose a different username.");
//            //            return View(model);
//            //        }

//            //        // Update the username
//            //        userToUpdate.Id = model.Id;
//            //    }

//                //// Update the password if it's not empty
//                //if (!string.IsNullOrEmpty(model.Password))
//                //{
//                //    userToUpdate.Password = model.Password;
//                //}

//                //// Update the department if it's not empty
//                //if (!string.IsNullOrEmpty(model.Department))
//                //{
//                //    userToUpdate.Department = model.Department;
//                //}
//            //}

//            // Save changes to the database
//           // _dbContext.SaveChanges();

//            return RedirectToAction("DisplayUsers");
//        }





//        [HttpPost]
//        public IActionResult DeleteUser(int? id)
//        {
//            // Retrieve the user details from the database based on the username
//            var user = _dbContext.Helpersusers.FirstOrDefault(u => u.Id == id);
//            if (user == null)
//            {
//                return NotFound(); // Return a 404 Not Found if user is not found
//            }

//            // Remove the user from the database
//            _dbContext.Helpersusers.Remove(user);
//            _dbContext.SaveChanges();

//            // Redirect to the display users page after the deletion is successful
//            return RedirectToAction("DisplayUsers");
//        }



//        public IActionResult Privacy()
//        {
//            return View();
//        }

//        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
//        public IActionResult Error()
//        {
//            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
//        }
//    }

//}


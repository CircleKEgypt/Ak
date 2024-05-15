using CK.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Security.Claims;
using Microsoft.AspNetCore.Cryptography.KeyDerivation;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
namespace CK.Controllers
{
    public class LoginController : Controller
    {
        private readonly ILogger<LoginController> _logger;
        private readonly CkproUsersContext _dbContext;
        private static readonly List<RptUser3> Users = new List<RptUser3>();
        public static string HashPassword(string password)
        {
            // Generate a salt
            var salt = new byte[16]; // 128 bits
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(salt);
            }

            // Hash the password with the salt
            var hashedPassword = KeyDerivation.Pbkdf2(
                password: password,
                salt: salt,
                prf: KeyDerivationPrf.HMACSHA256,
                iterationCount: 10000, // Recommended number of iterations
                numBytesRequested: 24); // 192 bits

            // Combine the salt and hashed password
            var hashBytes = new byte[salt.Length + hashedPassword.Length];
            Buffer.BlockCopy(salt, 0, hashBytes, 0, salt.Length);
            Buffer.BlockCopy(hashedPassword, 0, hashBytes, salt.Length, hashedPassword.Length);

            // Convert to base64 string
            return Convert.ToBase64String(hashBytes);
        }
        public static bool VerifyPassword(string hashedPassword, string providedPassword)
        {
            // Convert the hashed password back to bytes
            var hashBytes = Convert.FromBase64String(hashedPassword);

            // Extract the salt and hashed password
            var salt = new byte[16]; // Assuming 128 bits
            var storedHashedPassword = new byte[24]; // Assuming 192 bits
            Buffer.BlockCopy(hashBytes, 0, salt, 0, salt.Length);
            Buffer.BlockCopy(hashBytes, salt.Length, storedHashedPassword, 0, storedHashedPassword.Length);

            // Hash the provided password with the extracted salt
            var hashedProvidedPassword = KeyDerivation.Pbkdf2(
                password: providedPassword,
                salt: salt,
                prf: KeyDerivationPrf.HMACSHA256,
                iterationCount: 10000, // Must match the iteration count used when hashing
                numBytesRequested: 24); // Must match the number of bytes requested when hashing

            // Compare the hashed passwords
            return hashedProvidedPassword.SequenceEqual(storedHashedPassword);
        }
        public LoginController(ILogger<LoginController> logger, CkproUsersContext dbContext)
        {
            _logger = logger;
            _dbContext = dbContext;
        }
        [HttpGet]
        [ResponseCache(Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Login()
        {
            ClaimsPrincipal claimUser = HttpContext.User;
            if (claimUser.Identity.IsAuthenticated)
            {
                // Clear existing session when displaying the login page
                HttpContext.Session.Clear();

                return RedirectToAction("Home", "Home");
            }

            return View();
        }
        public class PasswordVerifier
        {
            public static bool VerifyPassword(string password, string encodedHash)
            {
                // Hash the password
                using (SHA256 sha256Hash = SHA256.Create())
                {
                    byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(password));

                    // Convert byte array to a Base64-encoded string
                    string base64Hash = Convert.ToBase64String(bytes);

                    // Compare the hashes
                    return base64Hash.Equals(encodedHash);
                }
            }
        }
        public class SessionCheckMiddleware
        {
            private readonly RequestDelegate _next;

            public SessionCheckMiddleware(RequestDelegate next)
            {
                _next = next;
            }

            public async Task InvokeAsync(HttpContext context)
            {
                // Check if the request is for the login page
                if (context.Request.Path.StartsWithSegments("/Login/Login"))
                {
                    await _next(context);
                    return;
                }

                // Check if the session contains a username
                var sessionUsername = context.Session.GetString("username");
                if (string.IsNullOrEmpty(sessionUsername))
                {
                    // Check if the response has already started
                    if (!context.Response.HasStarted)
                    {
                        // Redirect to the login page if the session is null or the user is not authenticated
                        context.Response.Redirect("/Login/Login");
                        return;
                    }
                }

                // If the session is valid or the response has already started, continue with the next middleware
                await _next(context);
            }
        }
        [HttpPost]
        [ResponseCache(Location = ResponseCacheLocation.None, NoStore = true)]
        [HttpPost]
        public async Task<IActionResult> Login(VMLogin modellogin)
        {
            // Debugging: Log the current session state
            var sessionUsername = HttpContext.Session.GetString("username"); // Corrected key
            Console.WriteLine($"SessionUsername: {sessionUsername}");
            //if (HttpContext.Request.Query.ContainsKey("preventBack"))
            //{
            // Clear authentication cookies if any (extra measure)
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

            // Clear session on login
            HttpContext.Session.Clear();
            //}

            if (string.IsNullOrWhiteSpace(modellogin.username) || string.IsNullOrWhiteSpace(modellogin.Password))
            {
                TempData["ValidateMessage"] = "Username and password are required.";
                return View("Login");
            }
            string hashedProvidedPassword = encrypt(modellogin.Password!);

            var authenticatedUser = _dbContext.RptUsers
      .FirstOrDefault(u => (u.Username == modellogin.username && u.Password == hashedProvidedPassword)|| (u.Username2 == modellogin.username && u.Password == hashedProvidedPassword));

            // Check if the user exists and the password is correct
            if (authenticatedUser != null)//&& VerifyPassword(authenticatedUser.Spassword, modellogin.Password))
            {
                // Create claims for the authenticated user
                List<Claim> claims = new List<Claim>
            {
                new Claim(ClaimTypes.NameIdentifier, authenticatedUser.Username),
                new Claim(ClaimTypes.NameIdentifier, authenticatedUser.Username2),
                new Claim("OtherProperties", "example role")
            };

                ClaimsIdentity claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
                AuthenticationProperties properties = new AuthenticationProperties
                {
                    AllowRefresh = true,
                    // Set the ExpiresUtc to a past date to clear the cookie upon closing the browser
                 //   ExpiresUtc = DateTimeOffset.UtcNow.AddMinutes(-1)
                };

                // Sign in the user
                await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), properties);

                // Store username in session for future use
                if (authenticatedUser.Username!="")
                {
                    HttpContext.Session.SetString("Username", authenticatedUser.Username);

                }
                else
                    HttpContext.Session.SetString("Username", authenticatedUser.Username2);
                HttpContext.Session.SetString("Role", authenticatedUser.Role);
                CkproUsersContext db2 = new CkproUsersContext();
                var username = HttpContext.Session.GetString("Username");
                bool isDmanager = db2.RptUsers.Any(s => s.Dmanager == username);
                bool isUsername = db2.RptUsers.Any(s => s.Username == username && (s.Storenumber != null || s.Storenumber != " "));
                IQueryable<Storeuser> query;
                if (isUsername)
                {
                    return RedirectToAction("HomeStore", "Home");

                }
                else
                {
                    return RedirectToAction("Home", "Home");
                }
            }

            if (HttpContext.Session.GetString("LoggedOut") == "true")
            {
                TempData["PreventBack"] = true;
                HttpContext.Session.SetString("LoggedOut", "false"); // Reset the session variable
            }

            return View();
        }
        [HttpGet]
        //[ValidateAntiForgeryToken]
        public IActionResult CreateUser()
        {
            var user = new User();
            var username = HttpContext.Session.GetString("Username");
            var Role = HttpContext.Session.GetString("Role");
            ViewBag.Username = username;
            ViewBag.Role = Role;
            return View();
        }
        public string encrypt(string clearText)
        {
            string EncryptionKey = "MAKV2SPBNI99212";
            byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return clearText;
        }
        public string Decrypt(string clearText)
        {
            string DecryptionKey = "MAKV2SPBNI99212";
            byte[] clearBytes = Convert.FromBase64String(clearText);
            using (Aes decryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(DecryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                decryptor.Key = pdb.GetBytes(32);
                decryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, decryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return clearText;
        }


        [HttpPost]
       //[ValidateAntiForgeryToken]
        public async Task<IActionResult> CreateUser ([Bind("Id,User1,Password,Role,Department")] User user)
        {
            // Encrypt the password before saving
            user.Role ??= "0";
            user.Department ??= "0";
            user.Password = encrypt(user.Password);

            _dbContext.Add(user);
            await _dbContext.SaveChangesAsync();
            return RedirectToAction("DisplayUsers");
        }
        public IActionResult DisplayUsers()
        {
            var username = HttpContext.Session.GetString("Username");
            var Role = HttpContext.Session.GetString("Role");
            ViewBag.Username = username;
            ViewBag.Role = Role;
            var users = _dbContext.Users.ToList();
            return View(users);
        }

        [HttpGet]
        public async Task<IActionResult> EditUser(int? id)
        {
            var username = HttpContext.Session.GetString("Username");
            var Role = HttpContext.Session.GetString("Role");
            ViewBag.Username = username;
            ViewBag.Role = Role;
            var user = await _dbContext.Users.FindAsync(id);
            user.Password = Decrypt(user.Password);
            return View(user);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> EditUser(int id, [Bind("Id,User1,Password,Role,Department")] User model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }
            model.Role ??= "0";
            model.Department ??= "0";
            model.Password = encrypt(model.Password);
            _dbContext.Update(model);
            await _dbContext.SaveChangesAsync();

            return RedirectToAction("DisplayUsers");
        }





        [HttpPost]
        public IActionResult DeleteUser(int? id)
        {
            // Retrieve the user details from the database based on the username
            var user = _dbContext.Users.FirstOrDefault(u => u.Id == id);
            if (user == null)
            {
                return NotFound(); // Return a 404 Not Found if user is not found
            }

            // Remove the user from the database
            _dbContext.Users.Remove(user);
            _dbContext.SaveChanges();

            // Redirect to the display users page after the deletion is successful
            return RedirectToAction("DisplayUsers");
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }

}


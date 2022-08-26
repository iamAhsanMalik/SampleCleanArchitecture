#region Packages
using Application.Contracts.Infrastructure;
using Application.DTOs.AccountDTOs;
using Application.Models.IdentityModels;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.WebUtilities;
using System.Security.Claims;
using System.Text;
#endregion

namespace Accounting.Controllers;

[Authorize]
public class AccountController : Controller
{
    private readonly IEmailService _emailSender;
    private readonly ISmsService _smsSender;
    private readonly SignInManager<ApplicationUser> _signInManager;
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly IFileHelpers _fileHelpers;

    public AccountController(IEmailService emailSender, ISmsService smsSender, SignInManager<ApplicationUser> signInManager, UserManager<ApplicationUser> userManager, IFileHelpers fileHelpers)
    {
        _emailSender = emailSender;
        _smsSender = smsSender;
        _signInManager = signInManager;
        _userManager = userManager;
        _fileHelpers = fileHelpers;
    }

    #region  Login
    //
    // GET: /Account/Login
    [HttpGet, AllowAnonymous]
    public IActionResult Login(string? returnUrl = null)
    {
        ViewData["ReturnUrl"] = returnUrl;
        return View();
    }

    //
    // POST: /Account/Login
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> Login(LoginDTO model, string? returnUrl = null)
    {

        returnUrl ??= Url.Content("/");

        returnUrl = returnUrl ?? Url.Content("/");

        ViewData["ReturnUrl"] = returnUrl;
        if (ModelState.IsValid)
        {
            // This doesn't count login failures towards account lockout
            // To enable password failures to trigger account lockout, set lockoutOnFailure: true
            var signInResult = await _signInManager.PasswordSignInAsync(model.Email, model.Password, model.RememberMe, lockoutOnFailure: false);
            if (signInResult?.Succeeded == true)
            {
                return RedirectToLocal(returnUrl);
            }
            if (signInResult?.RequiresTwoFactor == true)
            {
                return RedirectToAction(nameof(SendCode), new { ReturnUrl = returnUrl, model.RememberMe });
            }
            if (signInResult?.IsLockedOut == true)
            {
                return View("Lockout");
            }
            else
            {
                ModelState.AddModelError(string.Empty, "Invalid login attempt.");
                return View(model);
            }
        }
        // If we got this far, something failed, redisplay form
        return View(model);
    }
    #endregion

    #region Register
    //
    // GET: /Account/Register
    [HttpGet, AllowAnonymous]
    public IActionResult Register(string? returnUrl = null)
    {
        ViewData["ReturnUrl"] = returnUrl;
        return View();
    }

    //
    // POST: /Account/Register
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> Register(RegisterDTO model, string? returnUrl = null)
    {

        returnUrl ??= Url.Content(contentPath: "/");

        returnUrl = returnUrl ?? Url.Content("/");

        ViewData["ReturnUrl"] = returnUrl;
        if (ModelState.IsValid)
        {
            var user = new ApplicationUser { UserName = model.Email, Email = model.Email };

            var registerResult = await _userManager.CreateAsync(user, model.Password);

            if (registerResult?.Succeeded == true)
            {
                // Send an email with this link to verify newly register account
                var code = await _userManager.GenerateEmailConfirmationTokenAsync(user);

                var callbackUrl = Url.Action("ConfirmEmail", "Account", new { userId = user.Id, code }, protocol: HttpContext.Request.Scheme);

                var callbackUrl = Url.Action("ConfirmEmail", "Account", new { userId = user.Id, code = code }, protocol: HttpContext.Request.Scheme);

                if (!string.IsNullOrEmpty(user.Email))
                {
                    await _emailSender.SendEmailAsync(user.Email, "Confirm your email",
                    $"Please confirm your account by clicking here: <a href='{callbackUrl}'>Confirm Your Account</a>.");
                }
                await _signInManager.SignInAsync(user, isPersistent: false);
                _logger.LogInformation(3, "User created a new account with password.");
                return RedirectToLocal(returnUrl);
            }
            AddErrors(registerResult);
        }

        // If we got this far, something failed, redisplay form
        return View(model);
    }

    #endregion

    #region Logout
    //
    // POST: /Account/Logout
    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> Logout()
    {
        await _signInManager.SignOutAsync();
        return RedirectToAction(nameof(HomeController.Index), nameof(HomeController));
    }
    #endregion

    #region External Login
    //
    // POST: /Account/ExternalLogin
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public IActionResult ExternalLogin(string provider, string? returnUrl = null)
    {
        // Request a redirect to the external login provider.
        // Request a redirect to the external login provider.
        var redirectUrl = Url.Action("ExternalLoginCallback", "Account", new { ReturnUrl = returnUrl });
        var properties = _signInManager.ConfigureExternalAuthenticationProperties(provider, redirectUrl);
        return Challenge(properties, provider);
    }

    //
    // GET: /Account/ExternalLoginCallback
    [HttpGet]
    [AllowAnonymous]
    public async Task<IActionResult> ExternalProviderCallBack(string? returnUrl = null, string? remoteError = null)
    {
        if (remoteError != null)
        {
            ModelState.AddModelError(string.Empty, $"Error from external provider: {remoteError}");
            return View(nameof(Login));
        }
        var info = await _signInManager.GetExternalLoginInfoAsync();

        if (info == null)
        {
            return RedirectToAction(nameof(Login));
        }

        // Sign in the user with this external login provider if the user already has a login.
        var result = await _signInManager.ExternalLoginSignInAsync(info.LoginProvider, info.ProviderKey, isPersistent: false);
        if (result.Succeeded)
        {
            // Update any authentication tokens if login succeeded
            await _signInManager.UpdateExternalAuthenticationTokensAsync(info);

            return RedirectToLocal(returnUrl ?? nameof(HomeController.Index));
        }
        if (result.RequiresTwoFactor)
        {
            return RedirectToAction(nameof(SendCode), new { ReturnUrl = returnUrl });
        }
        if (result.IsLockedOut)
        {
            return View(nameof(HomeController.Lockout));
        }
        else
        {
            // If the user does not have an account, then ask the user to create an account.
            ViewData["ReturnUrl"] = returnUrl;
            ViewData["ProviderDisplayName"] = info.ProviderDisplayName;
            var email = info.Principal.FindFirstValue(ClaimTypes.Email);
            return View("ExternalLoginConfirmation", new ExternalLoginConfirmationDTO { Email = email });
        }
    }

    //
    // POST: /Account/ExternalLoginConfirmation
    [HttpPost]
    [AllowAnonymous]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ExternalLoginConfirmation(ExternalLoginConfirmationDTO model, string? returnUrl = null)
    {
        if (ModelState.IsValid)
        {
            // Get the information about the user from the external login provider
            var info = await _signInManager.GetExternalLoginInfoAsync();
            if (info == null)
            {
                return View("ExternalLoginFailure");
            }
            var user = new ApplicationUser { UserName = model.Email, Email = model.Email };
            var result = await _userManager.CreateAsync(user);
            if (result.Succeeded)
            {
                result = await _userManager.AddLoginAsync(user, info);
                if (result.Succeeded)
                {
                    await _signInManager.SignInAsync(user, isPersistent: false);
                    _logger.LogInformation(6, "User created an account using {Name} provider.", info.LoginProvider);

                    // Update any authentication tokens as well
                    await _signInManager.UpdateExternalAuthenticationTokensAsync(info);

                    return RedirectToLocal(returnUrl ?? nameof(HomeController.Index));
                }
            }
            AddErrors(result);
        }

        ViewData["ReturnUrl"] = returnUrl;
        return View(model);
    }
    #endregion

    #region ConfirmEmail
    // GET: /Account/ConfirmEmail
    [HttpGet, AllowAnonymous]
    public async Task<IActionResult> ConfirmEmail(string userId, string code)
    {
        if (userId == null || code == null)
        {
            return View("Error");
        }
        var user = await _userManager.FindByIdAsync(userId);
        if (user == null)
        {
            return View("Error");
        }
        var result = await _userManager.ConfirmEmailAsync(user, code);
        return View(result.Succeeded ? "ConfirmEmail" : "Error");
    }
    #endregion

    #region Forgot Password
    //
    // GET: /Account/ForgotPassword
    [HttpGet, AllowAnonymous]
    public IActionResult ForgotPassword()
    {
        return View();
    }

    //
    // POST: /Account/ForgotPassword
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> ForgotPassword(ForgotPasswordDTO model)
    {
        if (ModelState.IsValid)
        {
            if (!string.IsNullOrEmpty(model.Email))
            {
                var user = await _userManager.FindByEmailAsync(model.Email);
                if (user == null || !await _userManager.IsEmailConfirmedAsync(user))
                {
                    // Don't reveal that the user does not exist or is not confirmed for security purposes
                    return View("ForgotPasswordConfirmation");
                }
                else
                {
                    var code = await _userManager.GeneratePasswordResetTokenAsync(user);
                    code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));

                    var callbackUrl = Url.Action("ResetPassword", "Account", new { userId = user.Id, passwordResetToken = code }, protocol: Request.Scheme);
                    var emailContent = await _fileHelpers.EmailTemplatesReaderAsync("PasswordReset.html");
                    var emailSubject = emailContent.Replace("$$ResetPasswordLink$$", callbackUrl).Replace("$$CurrentYear$$", DateTime.Now.Year.ToString());
                    await _emailSender.SendEmailAsync(model.Email, "Reset Password",
                    emailSubject);
                    return View("ForgotPasswordConfirmation");
                }
            }

            var callbackUrl = Url.Action("ResetPassword", "Account", new { userId = user.Id, emailConfirmationToken = code }, protocol: Request.Scheme);
            await _emailSender.SendEmailAsync(model.Email, "Reset Password",
            $@"Please reset your password by clicking here: <a href='\{callbackUrl}\'>Reset Password</a>.");
            return View("ForgotPasswordConfirmation");
        }

        // If we got this far, something failed, redisplay form
        return View(model);
    }

    //
    // GET: /Account/ForgotPasswordConfirmation
    [HttpGet]
    [AllowAnonymous]
    public IActionResult ForgotPasswordConfirmation()
    {
        return View();
    }
    #endregion

    #region Reset Password
    //
    // GET: /Account/ResetPassword
    [HttpGet, AllowAnonymous]

    public IActionResult ResetPassword(string? passwordResetToken = null)
    {
        return passwordResetToken == null ? View("Error") : View();
        //return View();
    }

    //
    // POST: /Account/ResetPassword
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> ResetPassword(ResetPasswordDTO model)
    {
        if (!ModelState.IsValid)
        {
            return View(model);
        }
        if (!string.IsNullOrEmpty(model.Email))
        {
            var user = await _userManager.FindByEmailAsync(model.Email);
            if (user == null)
            {
                // Don't reveal that the user does not exist
                return RedirectToAction(nameof(AccountController.ResetPasswordConfirmation), "Account");
            }

            if (user != null && !string.IsNullOrEmpty(model.PasswordResetToken))
            {
                model.PasswordResetToken = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(model.PasswordResetToken));
                var isValid = await _userManager.VerifyUserTokenAsync(user, _userManager.Options.Tokens.PasswordResetTokenProvider, "ResetPassword", model.PasswordResetToken);
                if (isValid)
                {
                    var result = await _userManager.ResetPasswordAsync(user, model.PasswordResetToken, model.Password);
                    if (result.Succeeded)
                    {
                        return RedirectToAction(nameof(AccountController.ResetPasswordConfirmation), "Account");
                    }
                    AddErrors(result);
                }

                if (!string.IsNullOrEmpty(model.Code) && !string.IsNullOrEmpty(model.Password))
                {
                    var result = await _userManager.ResetPasswordAsync(user, model.Code, model.Password);
                    if (result.Succeeded)
                    {
                        return RedirectToAction(nameof(AccountController.ResetPasswordConfirmation), "Account");
                    }
                    AddErrors(result);

                }
            }
        }
        return View();
    }
    //
    // GET: /Account/ResetPasswordConfirmation
    [HttpGet, AllowAnonymous]
    public IActionResult ResetPasswordConfirmation()
    {
        return View();
    }
    #endregion

    #region Two Factor Auth
    //
    // GET: /Account/SendCode
    [HttpGet, AllowAnonymous]
    public async Task<ActionResult> SendCode(string? returnUrl = null, bool rememberMe = false)
    {
        var user = await _signInManager.GetTwoFactorAuthenticationUserAsync();
        if (user == null)
        {
            return View(nameof(HomeController.Error));
        }
        var userFactors = await _userManager.GetValidTwoFactorProvidersAsync(user);
        var factorOptions = userFactors.Select(purpose => new SelectListItem { Text = purpose, Value = purpose }).ToList();
        return View(new SendCodeDTO<SelectListItem> { Providers = factorOptions, ReturnUrl = returnUrl, RememberMe = rememberMe });
    }

    //
    // POST: /Account/SendCode
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> SendCode(SendCodeDTO<SelectListItem> model)
    {
        if (!ModelState.IsValid)
        {
            return View();
        }

        var user = await _signInManager.GetTwoFactorAuthenticationUserAsync();
        if (user == null)
        {
            return View("Error");
        }

        if (model.SelectedProvider == "Authenticator")
        {
            return RedirectToAction(nameof(VerifyAuthenticatorCode), new { model.ReturnUrl, model.RememberMe });
        }

        // Generate the token and send it
        if (!string.IsNullOrEmpty(model.SelectedProvider))
        {
            var code = await _userManager.GenerateTwoFactorTokenAsync(user, model.SelectedProvider);
            if (string.IsNullOrWhiteSpace(code))
            {
                return View("Error");
            }

            var message = "Your security code is: " + code;

            if (model.SelectedProvider == "Email")
            {
                var userEmail = await _userManager.GetEmailAsync(user);
                if (!string.IsNullOrEmpty(userEmail))
                {
                    await _emailSender.SendEmailAsync(userEmail, "Security Code", message);
                }
            }
            else if (model.SelectedProvider == "Phone")
            {
                var userPhone = await _userManager.GetPhoneNumberAsync(user);
                if (!string.IsNullOrEmpty(userPhone))
                {
                    await _smsSender.SendSmsAsync(userPhone, message);
                }
            }
        }
        return RedirectToAction(nameof(VerifyCode), new { Provider = model.SelectedProvider, model.ReturnUrl, model.RememberMe });
    }

    //
    // GET: /Account/VerifyCode
    [HttpGet, AllowAnonymous]
    public async Task<IActionResult> VerifyCode(string provider, bool rememberMe, string? returnUrl = null)
    {
        // Require that the user has already logged in via username/password or external login
        var user = await _signInManager.GetTwoFactorAuthenticationUserAsync();
        if (user == null)
        {
            return View("Error");
        }
        return View(new VerifyCodeDTO { Provider = provider, ReturnUrl = returnUrl, RememberMe = rememberMe });
    }

    //
    // POST: /Account/VerifyCode
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> VerifyCode(VerifyCodeDTO model)
    {
        if (!ModelState.IsValid)
        {
            return View(model);
        }

        // The following code protects for brute force attacks against the two factor codes.
        // If a user enters incorrect codes for a specified amount of time then the user account
        // will be locked out for a specified amount of time.

        Microsoft.AspNetCore.Identity.SignInResult? signInResult = default;
        if (!string.IsNullOrEmpty(model.Provider) && !string.IsNullOrEmpty(model.Code))
        {
            signInResult = await _signInManager.TwoFactorSignInAsync(model.Provider, model.Code, model.RememberMe, model.RememberBrowser);
        }

        if (signInResult?.Succeeded == true)
        {
            return RedirectToLocal(model.ReturnUrl ?? nameof(HomeController.Index));
        }
        if (signInResult?.IsLockedOut == true)
        {
            _logger.LogWarning(7, "User account locked out.");
            return View("Lockout");
        }
        else
        {
            ModelState.AddModelError(string.Empty, "Invalid code.");
            return View(model);
        }
    }

    //
    // GET: /Account/VerifyAuthenticatorCode
    [HttpGet, AllowAnonymous]
    public async Task<IActionResult> VerifyAuthenticatorCode(bool rememberMe, string? returnUrl = null)
    {
        // Require that the user has already logged in via username/password or external login
        var user = await _signInManager.GetTwoFactorAuthenticationUserAsync();
        if (user == null)
        {
            return View("Error");
        }
        return View(new VerifyAuthenticatorCodeDTO { ReturnUrl = returnUrl, RememberMe = rememberMe });
    }

    //
    // POST: /Account/VerifyAuthenticatorCode
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> VerifyAuthenticatorCode(VerifyAuthenticatorCodeDTO model)
    {
        if (!ModelState.IsValid)
        {
            return View(model);
        }

        // The following code protects for brute force attacks against the two factor codes.
        // If a user enters incorrect codes for a specified amount of time then the user account
        // will be locked out for a specified amount of time.
        Microsoft.AspNetCore.Identity.SignInResult? result = default;
        if (!string.IsNullOrEmpty(model.Code))
        {
            result = await _signInManager.TwoFactorAuthenticatorSignInAsync(model.Code, model.RememberMe, model.RememberBrowser);
        }

        if (result?.Succeeded == true)
        {
            return RedirectToLocal(model.ReturnUrl ?? nameof(HomeController.Index));
        }
        if (result?.IsLockedOut == true)
        {
            if (result != null && result.Succeeded)
            {
                return RedirectToLocal(model.ReturnUrl ?? nameof(HomeController.Index));
            }
        }

        if (result != null && result.IsLockedOut)

        {
            _logger.LogWarning(7, "User account locked out.");
            return View(nameof(HomeController.Lockout));
        }
        else
        {
            ModelState.AddModelError(string.Empty, "Invalid code.");
            return View(model);
        }
    }
    #endregion

    #region Recovery Code
    //
    // GET: /Account/UseRecoveryCode
    [HttpGet, AllowAnonymous]
    public async Task<IActionResult> UseRecoveryCode(string? returnUrl = null)
    {
        // Require that the user has already logged in via username/password or external login
        var user = await _signInManager.GetTwoFactorAuthenticationUserAsync();
        if (user == null)
        {
            return View(nameof(HomeController.Error));
        }
        return View(new UseRecoveryCodeDTO { ReturnUrl = returnUrl });
    }

    //
    // POST: /Account/UseRecoveryCode
    [HttpPost, AllowAnonymous, ValidateAntiForgeryToken]
    public async Task<IActionResult> UseRecoveryCode(UseRecoveryCodeDTO model)
    {
        if (ModelState.IsValid)
        {
            if (!string.IsNullOrEmpty(model.Code))
            {
                var result = await _signInManager.TwoFactorRecoveryCodeSignInAsync(model.Code);

                if (result?.Succeeded == true)
                {
                    if (result != null && result.Succeeded)

                    {
                        return RedirectToLocal(model.ReturnUrl ?? nameof(HomeController.Index));
                    }
                    else
                    {
                        ModelState.AddModelError(string.Empty, "Invalid code.");
                        return View(model);
                    }
                }
            }
        }

        // If we got this far, something failed, redisplay form
        return View(model);
    }
    #endregion

    #region Helpers

    private void AddErrors(IdentityResult? result)
    {
        if (result != null)
        {
            foreach (var error in result.Errors)
            {
                ModelState.AddModelError(string.Empty, error.Description);
            }
        }
    }

    private IActionResult RedirectToLocal(string returnUrl)
    {
        if (Url.IsLocalUrl(returnUrl))
        {
            return Redirect(returnUrl);
        }
        else
        {
            return RedirectToAction(nameof(HomeController.Index), nameof(HomeController));
        }
    }
    #endregion
}

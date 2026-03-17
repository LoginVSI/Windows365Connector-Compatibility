using System;
using System.Linq;
using System.Threading;
using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;

public class Windows365ConnectorScript : ConnectorScriptBase
{
    private const string LogPrefix = "[Win365]";

    private void LogInfo(string message) => Log($"{LogPrefix} {message}");

    private void LogDebug(string message) => Log($"{LogPrefix} [DEBUG] {message}");

    private void LogWarning(string message) => Log($"{LogPrefix} [WARNING] {message}");

    private void LogError(string message) => Log($"{LogPrefix} [ERROR] {message}");

    private IAutomationElement? _windowElement;

    private CancellationToken _cancellationToken;

    public override void RunScript(CancellationToken cancellationToken = new())
    {
        _cancellationToken = cancellationToken;

        try
        {
            LogConfiguration();
            StartWindowsAppAndWaitForMainWindow();
            HandleSignInFlow();
            SelectCloudPc();
            PostConnectionEstablished();

            LogInfo("Script execution completed");
            CleanupAndExit();
        }
        catch (OperationCanceledException)
        {
            LogWarning("Script execution cancelled");
            CleanupAndExit();
        }
        catch (Exception ex)
        {
            LogError($"Script execution failed: {ex.Message}");
            CleanupAndExit();
            throw;
        }
    }

    private void LogConfiguration()
    {
        LogInfo("Script execution started");
        LogInfo("Email: [REDACTED]");
        LogInfo($"Cloud PC: {CloudPcTitle}");
        LogInfo($"TOTP: {(string.IsNullOrEmpty(TotpSecret) ? "Disabled" : "Enabled")}");
    }

    private void StartWindowsAppAndWaitForMainWindow()
    {
        _cancellationToken.ThrowIfCancellationRequested();

        START(forceKillOnExit: true);
        LogInfo("Starting Windows 365 app");

        Wait(1);

        var mwDeadline = DateTime.UtcNow.AddSeconds(TimeoutSeconds);
        while (MainWindow == null && DateTime.UtcNow < mwDeadline)
        {
            _cancellationToken.ThrowIfCancellationRequested();

            LogDebug("Waiting for MainWindow...");
            Wait(1);
        }

        if (MainWindow == null)
            throw new Exception($"StartWindowsApp: MainWindow not found after {TimeoutSeconds}s timeout");

        LogInfo("MainWindow found and focused");
        MainWindow.Focus();
    }

    private void HandleSignInFlow()
    {
        const int globalWaitInSeconds = 5;
        const int globalCpmToType = 1000;

        if (IsItInSignInState(globalWaitInSeconds))
        {
            if (_windowElement is null)
                throw new Exception("HandleSignInFlow: Window element is null");

            _cancellationToken.ThrowIfCancellationRequested();
            var signInButton = _windowElement.FindAutomationElementByXPathOrInformation("//Group/Button[@Name=\"Sign in\"]",
                                                                                        "",
                                                                                        "",
                                                                                        "Sign in",
                                                                                        "Button",
                                                                                        globalWaitInSeconds);

            _cancellationToken.ThrowIfCancellationRequested();
            CleanUpAuthWindows(globalWaitInSeconds);

            signInButton.Click();

            Wait(1);
        }
        else if (!TryResetToSignInState(TimeoutSeconds, true, globalWaitInSeconds))
            throw new Exception($"HandleSignInFlow: Could not reset to sign-in state after {TimeoutSeconds}s");

        if (_windowElement is null)
            throw new Exception("HandleSignInFlow: Window element is null");

        _cancellationToken.ThrowIfCancellationRequested();
        var useAnotherAccount = _windowElement.FindAutomationElementByXPathOrInformation("//Group/List/Button[@ClassName=\"account-button\"][@Name=\"Use another account\"]",
                                                                                         "", "account-button", "Use another account", "Button",
                                                                                         globalWaitInSeconds,
                                                                                         continueOnError: true);

        useAnotherAccount?.Click();

        _cancellationToken.ThrowIfCancellationRequested();
        Wait(3);

        var stopAt = DateTime.UtcNow.AddSeconds(TimeoutSeconds);
        IAutomationElement? authWindow = null;

        while (DateTime.UtcNow < stopAt && authWindow == null)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            authWindow = FindAutomationElementByXPathOrInformation("/Window[@ClassName=\"ApplicationFrameWindow\"]",
                                                                   "", "ApplicationFrameWindow", "", "Window",
                                                                   globalWaitInSeconds, continueOnError: true);
            Wait(1);
        }

        if (authWindow == null)
            throw new Exception($"HandleSignInFlow: Authentication window not found after {TimeoutSeconds}s");

        _cancellationToken.ThrowIfCancellationRequested();
        authWindow.FindAutomationElementByXPathOrInformation("//Group/Text[@Name=\"Sign in\"]",
                                                             "", "", "Sign in", "Text",
                                                             globalWaitInSeconds);

        LogDebug("Entering credentials");

        _cancellationToken.ThrowIfCancellationRequested();
        var editEnterYourEmail = authWindow.FindAutomationElementByXPathOrInformation("//Group/Edit[@Name=\"Enter your email, phone, or Skype.\"][@AutomationId=\"i0116\"]", "i0116",
                                                                                      "", "Enter your email, phone, or Skype.", "Edit",
                                                                                      globalWaitInSeconds);
        editEnterYourEmail.Click(forceFocus: false);
        Wait(1);

        MainWindow.Type(RuntimeEmail, forceFocus: false, cpm: globalCpmToType);

        _cancellationToken.ThrowIfCancellationRequested();
        var buttonNext = authWindow.FindAutomationElementByXPathOrInformation("//Group/Button[@Name=\"Next\"][@AutomationId=\"idSIButton9\"]",
                                                                              "idSIButton9", "", "Next", "Button",
                                                                              globalWaitInSeconds);
        buttonNext.Click(forceFocus: false);
        Wait(1);

        _cancellationToken.ThrowIfCancellationRequested();
        var editEnterPassword = authWindow.FindAutomationElementByXPathOrInformation("//Group/Edit[@AutomationId=\"i0118\"]",
                                                                                     "i0118", "", "", "Edit",
                                                                                     globalWaitInSeconds);
        editEnterPassword.Click(forceFocus: false);
        Wait(1);

        _cancellationToken.ThrowIfCancellationRequested();
        MainWindow.Type(RuntimePassword, forceFocus: false, cpm: globalCpmToType);

        _cancellationToken.ThrowIfCancellationRequested();
        var buttonSignin0 = authWindow.FindAutomationElementByXPathOrInformation("//Group/Button[@Name=\"Sign in\"][@AutomationId=\"idSIButton9\"]",
                                                                                 "idSIButton9", "", "Sign in", "Button",
                                                                                 globalWaitInSeconds);
        buttonSignin0.Click(forceFocus: false);
        Wait(1);

        _cancellationToken.ThrowIfCancellationRequested();
        var changeAuthenticApp = authWindow.FindAutomationElementByXPathOrInformation("//Group/Hyperlink[@AutomationId=\"signInAnotherWay\"]",
                                                                                      "signInAnotherWay", "", "", "Hyperlink",
                                                                                      globalWaitInSeconds, continueOnError: true);
        changeAuthenticApp?.Click(forceFocus: false);
        Wait(1);

        LogDebug("Entering TOTP code");

        _cancellationToken.ThrowIfCancellationRequested();
        var buttonUseCode = authWindow.FindAutomationElementByXPathOrInformation("//Group/List/ListItem/Button[@Name=\"Use a verification code\"]",
                                                                                 "", "", "Use a verification code", "Button",
                                                                                 globalWaitInSeconds, continueOnError: true);
        buttonUseCode?.Click(forceFocus: false);
        Wait(1);

        _cancellationToken.ThrowIfCancellationRequested();
        var editEnterCode = authWindow.FindAutomationElementByXPathOrInformation("//Group/Edit[@Name=\"Enter code\"][@AutomationId=\"idTxtBx_SAOTCC_OTC\"]",
                                                                                 "idTxtBx_SAOTCC_OTC", "", "Enter code", "Edit",
                                                                                 globalWaitInSeconds, continueOnError: true);
        editEnterCode?.Click(forceFocus: false);
        Wait(1);

        if (editEnterCode == null)
        {
            LogInfo("MFA was not detected. Sign-in flow completed.");
            return;
        }

        _cancellationToken.ThrowIfCancellationRequested();
        var buttonVerify = authWindow.FindAutomationElementByXPathOrInformation("//Group/Button[@Name=\"Verify\"][@AutomationId=\"idSubmit_SAOTCC_Continue\"]",
                                                                                "idSubmit_SAOTCC_Continue", "", "Verify", "Button",
                                                                                globalWaitInSeconds);

        _cancellationToken.ThrowIfCancellationRequested();

        var code = GenerateTotpCode(TotpSecret, TotpTimeStep, TotpDigits, TotpAlgorithm);
        MainWindow.Type(code, forceFocus: false, cpm: globalCpmToType);

        buttonVerify.Click(forceFocus: false);

        LogInfo("Sign-in flow completed");
    }

    private void SelectCloudPc()
    {
        const int globalWaitInSeconds = 5;
        const int globalCpmToType = 1000;

        LogInfo("Selecting Cloud PC");
        LogDebug("Waiting for sign-in to complete and Devices button to appear...");

        var devicesButtonFound = false;
        var stopAt = DateTime.UtcNow.AddSeconds(TimeoutSeconds);
        IAutomationElement? buttonDevices = null;
        IAutomationElement? devicesWindow = null;

        while (DateTime.UtcNow < stopAt && !devicesButtonFound)
        {
            _cancellationToken.ThrowIfCancellationRequested();

            try
            {
                devicesWindow = FindAutomationElementByXPathOrInformation("/Window[@ClassName=\"MainWindow\"][@Name=\"Windows App\"]",
                                                                          "", "MainWindow", "Windows App", "Window",
                                                                          globalWaitInSeconds, continueOnError: true);
                if (devicesWindow != null)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var skip = devicesWindow.FindAutomationElementByXPathOrInformation("/Pane/Document/Group/Button[@Name=\\\"Skip\\\"]",
                                                                                       "", "", "Skip", "Button", 2,
                                                                                       continueOnError: true);
                    skip?.Click();

                    _cancellationToken.ThrowIfCancellationRequested();
                    buttonDevices = devicesWindow.FindAutomationElementByXPathOrInformation("//Group/Button[@AutomationId=\"nav-devices\"]",
                                                                                           "nav-devices", "", "Devices Button*", "Button",
                                                                                           globalWaitInSeconds, continueOnError: true);

                    if (buttonDevices != null)
                    {
                        devicesButtonFound = true;
                        LogInfo("Devices button found - sign-in completed");
                    }
                    else
                    {
                        LogDebug("Devices button not yet available, waiting...");
                        Wait(1);
                    }
                }
                else
                    Wait(1);

            }
            catch (Exception ex)
            {
                LogDebug($"Error while waiting for Devices button: {ex.Message}");
                Wait(1);
            }
        }

        if (!devicesButtonFound || buttonDevices == null || devicesWindow == null)
            throw new Exception($"SelectCloudPc: Devices button not found after {TimeoutSeconds}s - sign-in may have failed");

        buttonDevices.Click();
        Wait(1);

        _cancellationToken.ThrowIfCancellationRequested();
        var editSearch0 = devicesWindow.FindAutomationElementByXPathOrInformation("//Group/Edit[@Name=\"Search\"]", "", "", "Search", "Edit", globalWaitInSeconds);
        editSearch0.Click();
        Wait(1);

        MainWindow.Type(CloudPcTitle, forceFocus: false, cpm: globalCpmToType);
        Wait(3);

        _cancellationToken.ThrowIfCancellationRequested();
        var connectCard = devicesWindow.FindAutomationElementByXPathOrInformation($"//Document/Group/Group/Group/Group/Button[@Name=\"{CloudPcTitle}\"]",
                                                                                  "", "flex flex-col", $"{CloudPcTitle}", "Button", 5);
        CleanUpCloudPcAuth(globalWaitInSeconds);

        _cancellationToken.ThrowIfCancellationRequested();
        var connections = FindAllAutomationElementByXPathOrInformation("/Window[@ClassName=\"TscShellContainerClass\"][@Name=\"" + CloudPcTitle + "*\"]",
                                                                       "", "TscShellContainerClass", $"{CloudPcTitle}*", "Window",
                                                                       globalWaitInSeconds,
                                                                       continueOnError: true);
        connectCard.Click(forceFocus: true);

        LogInfo("Waiting for connection to establish...");

        Wait(10);

        MainWindow.Focus();
        HandleCloudPcAuth(connections?.Count() ?? 0, globalWaitInSeconds, globalCpmToType);

        LogInfo($"Initiated connection to Cloud PC: {CloudPcTitle}");
    }

    private void PostConnectionEstablished()
    {
        const int globalWaitInSeconds = 5;

        _cancellationToken.ThrowIfCancellationRequested();

        MainWindow.Focus();
        Wait(globalWaitInSeconds);

        if (!TryResetToSignInState(TimeoutSeconds, false, globalWaitInSeconds))
            LogWarning("Could not reset to sign-in state after connection");
    }

    private void HandleCloudPcAuth(int currentConnections, int waitInSeconds, int globalCpmToType)
    {
        LogInfo("Handle CloudPc Auth if needed");

        var stopAt = DateTime.UtcNow.AddSeconds(TimeoutSeconds);
        var resolved = false;

        while (DateTime.UtcNow < stopAt && !resolved)
        {
            LogInfo("Search for the connection window");
            _cancellationToken.ThrowIfCancellationRequested();
            var connections = FindAllAutomationElementByXPathOrInformation("/Window[@ClassName=\"TscShellContainerClass\"][@Name=\"" + CloudPcTitle + "*\"]",
                                                                           "", "TscShellContainerClass", $"{CloudPcTitle}*", "Window",
                                                                           waitInSeconds,
                                                                           continueOnError: true);
            if (connections?.Count() > currentConnections)
            {
                LogDebug("Connection window found");
                resolved = true;
                continue;
            }

            LogInfo("Search for the auth web view");
            _cancellationToken.ThrowIfCancellationRequested();
            var auth = FindAutomationElementByInformation("webView", null, null, "Pane", waitInSeconds, continueOnError: true);
            if (auth == null)
                continue;

            Wait(1);

            LogInfo("Try enter the password");
            _cancellationToken.ThrowIfCancellationRequested();
            var group = auth.FindAutomationElementByInformation("i0281", null, null, "Group", waitInSeconds, continueOnError: true);

            if (group != null)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var editEnterPassword = group.FindAutomationElementByXPathOrInformation("//Edit[@AutomationId=\"i0118\"]", "i0118", "", "", "Edit", waitInSeconds);
                editEnterPassword?.Click(forceFocus: false);

                if (editEnterPassword != null)
                    MainWindow.Type(RuntimePassword, forceFocus: false, cpm: globalCpmToType);

                _cancellationToken.ThrowIfCancellationRequested();
                var buttonSignin = group.FindAutomationElementByXPathOrInformation("//Button[@Name=\"Sign in\"][@AutomationId=\"idSIButton9\"]", "idSIButton9", "", "Sign in", "Button", waitInSeconds);
                buttonSignin?.Click(forceFocus: false);
                if (buttonSignin != null)
                    Wait(1);

                _cancellationToken.ThrowIfCancellationRequested();
                var mfa = FindAutomationElementByInformation("webView", "WebView", "Sign in to your account", "Pane", waitInSeconds, continueOnError: true);

                _cancellationToken.ThrowIfCancellationRequested();
                var changeAuthenticApp = mfa.FindAutomationElementByXPathOrInformation("//Hyperlink[@AutomationId=\"signInAnotherWay\"]", "signInAnotherWay", "", "", "Hyperlink", waitInSeconds, continueOnError: true);
                changeAuthenticApp?.Click(forceFocus: false);

                Wait(1);

                LogDebug("Entering TOTP code");

                _cancellationToken.ThrowIfCancellationRequested();
                var buttonUseCode = mfa.FindAutomationElementByXPathOrInformation("//Button[@Name=\"Use a verification code\"]", "", "", "Use a verification code", "Button", waitInSeconds, continueOnError: true);
                buttonUseCode?.Click(forceFocus: false);
                Wait(1);

                _cancellationToken.ThrowIfCancellationRequested();
                var editEnterCode = mfa.FindAutomationElementByXPathOrInformation("//Edit[@Name=\"Enter code\"][@AutomationId=\"idTxtBx_SAOTCC_OTC\"]", "idTxtBx_SAOTCC_OTC", "", "Enter code", "Edit", waitInSeconds, continueOnError: true);
                editEnterCode?.Click(forceFocus: false);
                Wait(1);

                _cancellationToken.ThrowIfCancellationRequested();
                var buttonVerify = mfa.FindAutomationElementByXPathOrInformation("//Button[@Name=\"Verify\"][@AutomationId=\"idSubmit_SAOTCC_Continue\"]", "idSubmit_SAOTCC_Continue", "", "Verify", "Button", waitInSeconds, continueOnError: true);

                _cancellationToken.ThrowIfCancellationRequested();
                var code = GenerateTotpCode(TotpSecret, TotpTimeStep, TotpDigits, TotpAlgorithm);
                _cancellationToken.ThrowIfCancellationRequested();

                MainWindow.Type(code, forceFocus: false, cpm: globalCpmToType);

                buttonVerify?.Click(forceFocus: false);

                LogInfo("CloudPc sign-in flow completed");
            }

            LogInfo("Try approve the resource");
            Wait(2);

            _cancellationToken.ThrowIfCancellationRequested();
            var webView = FindAutomationElementByInformation("webView", "WebView", "Sign in to your account", "Pane", waitInSeconds, continueOnError: true);

            _cancellationToken.ThrowIfCancellationRequested();
            var buttonYes = webView?.FindAutomationElementByInformation(null, null, "Yes", "Button", waitInSeconds, continueOnError: true);
            buttonYes?.Click();

            if (buttonYes != null)
                Wait(1);

            resolved = true;
        }
    }

    private void CleanUpCloudPcAuth(int waitInSeconds)
    {
        LogInfo("Clean up CloudPc Auth if needed");

        var groups = FindAllAutomationElementByInformation("webView", null, null, "Pane", waitInSeconds, continueOnError: true);

        foreach (var group in groups)
        {
            var window = group.GetRootWindow();
            window?.Close();
        }
    }

    private void CleanUpAuthWindows(int waitInSeconds)
    {
        LogInfo("Clean up Auth if needed");

        var authWindows = FindAllAutomationElementByXPathOrInformation("/Window[@ClassName=\"ApplicationFrameWindow\"]",
                                                                       "", "ApplicationFrameWindow", "", "Window",
                                                                       waitInSeconds, continueOnError: true);

        if (authWindows == null)
            return;

        foreach (var window in authWindows)
            window?.AsWindow().Close();
    }

    private void CleanupAndExit()
    {
        LogInfo("Stopping application.");
        STOP();
        CleanUpAuthWindows(5);
        CleanUpCloudPcAuth(5);
        LogDebug("Application stopped");
    }

    private bool IsItInSignInState(int waitInSeconds)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        _windowElement = FindAutomationElementByXPathOrInformation("/Window[@ClassName=\"MainWindow\"][@Name=\"Windows App\"]",
                                                                   "", "MainWindow", "Windows App", "Window",
                                                                   timeout: waitInSeconds, continueOnError: true);

        Wait(1);

        _cancellationToken.ThrowIfCancellationRequested();
        var signInProbe = _windowElement?.FindAutomationElementByXPathOrInformation("//Group/Button[@Name=\"Sign in\"]",
                                                                                    "", "", "Sign in", "Button",
                                                                                    waitInSeconds, continueOnError: true);

        if (signInProbe == null)
            return false;

        LogInfo("Sign-in button found - ready for authentication");
        return true;
    }

    private bool TryResetToSignInState(int timeoutSeconds, bool useAnotherAccount, int waitInSeconds)
    {
        LogDebug("Resetting to sign-in state...");

        var stopAt = DateTime.UtcNow.AddSeconds(timeoutSeconds);

        while (DateTime.UtcNow < stopAt)
        {
            _cancellationToken.ThrowIfCancellationRequested();

            try
            {
                MainWindow.Focus();

                _cancellationToken.ThrowIfCancellationRequested();
                var devicesWindow = FindAutomationElementByXPathOrInformation("/Window[@ClassName=\"MainWindow\"][@Name=\"Windows App\"]",
                                                                              "", "MainWindow", "Windows App", "Window",
                                                                              waitInSeconds, continueOnError: true);

                _cancellationToken.ThrowIfCancellationRequested();
                var accountProbe = devicesWindow.FindAutomationElementByXPathOrInformation("//Group/Button[@Name=\"Account\"][@AutomationId=\"header-account-button\"]",
                                                                                           "header-account-button", "", "Account", "Button",
                                                                                           waitInSeconds, continueOnError: true);

                if (accountProbe != null)
                {
                    LogInfo("Already signed in - signing out");

                    try
                    {
                        accountProbe.Click();
                        LogDebug("Clicked Account button");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"TryResetToSignInState: Failed to click Account button - {ex.Message}", ex);
                    }

                    Wait(1);

                    _cancellationToken.ThrowIfCancellationRequested();
                    var signOutBtn = devicesWindow.FindAutomationElementByXPathOrInformation("//Group/Window/Button[@Name=\"Sign out\"][@AutomationId=\"sign-out\"]",
                                                                                             "sign-out", "", "Sign out", "Button", waitInSeconds);

                    if (signOutBtn != null)
                    {
                        try
                        {
                            signOutBtn.Click();
                            LogDebug("Clicked Sign out button");
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"TryResetToSignInState: Failed to click Sign out button - {ex.Message}", ex);
                        }
                    }
                }

                Wait(1);

                _cancellationToken.ThrowIfCancellationRequested();
                _windowElement = FindAutomationElementByXPathOrInformation("/Window[@ClassName=\"MainWindow\"][@Name=\"Windows App\"]",
                                                                           "", "MainWindow", "Windows App", "Window",
                                                                           timeout: waitInSeconds, continueOnError: true);

                Wait(1);


                _cancellationToken.ThrowIfCancellationRequested();
                var useAnother = _windowElement.FindAutomationElementByXPathOrInformation("//Group/List/Button[@ClassName=\"account-button\"][@Name=\"Use another account\"]",
                                                                                          "", "account-button", "Use another account", "Button",
                                                                                          waitInSeconds, continueOnError: true);

                if (useAnother != null)
                {
                    if (useAnotherAccount)
                    {
                        LogInfo("Use another account button found - ready for authentication");
                        return true;
                    }

                    Wait(1);

                    LogDebug("Closing account picker dialog");

                    _cancellationToken.ThrowIfCancellationRequested();
                    var buttonClose = _windowElement.FindAutomationElementByXPathOrInformation("//Document/Group/Button[@Name=\"Close\"]",
                                                                                               "", "", "Close", "Button",
                                                                                               waitInSeconds);
                    buttonClose.Click(forceFocus: false);
                }

                Wait(1);

                _cancellationToken.ThrowIfCancellationRequested();
                var signInProbe = _windowElement.FindAutomationElementByXPathOrInformation("//Group/Button[@Name=\"Sign in\"]",
                                                                                           "", "", "Sign in", "Button", waitInSeconds, continueOnError: true);

                if (signInProbe != null)
                {
                    LogInfo("Sign-in button found - ready for authentication");
                    return true;
                }

                var nothingHereYet = devicesWindow.FindAutomationElementByXPathOrInformation("//Document/Group/Group/Group/Text[@Name=\"Nothing here yet\"]", 
                                                                                                 "", "", "Nothing here yet", "Text", 5, continueOnError: true);

                if (nothingHereYet != null)
                {
                    LogInfo("Sign-out completed");
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogDebug($"Reset attempt failed: {ex.Message}");
            }

            Wait(1);
        }

        return false;
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using Logger;
using PS19.ATM.ReturnStatus;
//using Interop.UIAutomationCore;
using NativeDesktopAutomationAPI;
using System.Linq;

namespace AutomationControlsFinder
{
    class Program
    {
        static Window window = null;

        static string logsPath = Path.Combine(Directory.GetCurrentDirectory() + @"\Logs.txt");
        static string csvlogsPath = Path.Combine(Directory.GetCurrentDirectory() + @"\Controls.csv");

        static FlaUI.Core.Application app;
        static DesktopAutomation DesktopAuto;

        static string automationId = "";
        static string name = "";
        static string controltype = "";
        static string parent = "";
        static string x = "";
        static string y = "";
        static string width = "";
        static string height = "";
        static string result;
        static string[] ArayOfComboBoxValues;

        static void Main(string[] args)
        {
            try
            {
                 KillApplication();
                 ProcessStartInfo start = new ProcessStartInfo(@"C:\Program Files (x86)\Smart Wires\SmartInterface\SmartInterface.exe");

                 app = FlaUI.Core.Application.AttachOrLaunch(start);
                 DesktopAuto = new DesktopAutomation();
                using (var automation = new UIA3Automation())
                {
                    result = $"Window Name,Automation Id,Name,Control Type,LocationX,LocationY,Height,Width";
                    resultChild.Add(result);

                    //Login Window
                    GetLoginWindowControls(automation);

                    //Login To SI
                    LoginToSmartInterface(automation);

                    //System Configuration Window
                    GetSystemConfigurationWindowControls(automation);

                    //Site Status Window
                    GetSiteStatusWindowControls(automation);

                    ////Alarms Configuration Window
                    //GetAlarmConfigurationWindowControls(automation);   // Not getable because AutomationID change every time

                    //Phase Naming Scheme Window
                    GetPhaseNamingSchemeWindowControls(automation);

                    #region Online
                    ////Circuit Status Window
                    //GetCircuitStatusWindowControls(automation);

                    ////Alarms Status Window
                    //GetAlarmsStatusWindowControls(automation);

                    ////Structure Summary Window
                    //GetStructureSummaryWindowControls(automation); 
                    #endregion

                    //Reports Window
                    GetReportsWindowControls(automation);

                    ////Switch To Online/Configuration Window
                    //GetSwitchToOnlineOrConfigurationWindowControls(automation);

                    //User Management Window
                    GetUserManagementWindowControls(automation);

                    //Firmware Update Window
                    GetFirmwareUpdateWindowControls(automation);

                    //Export Configuration Window
                    GetExportConfigurationWindowControls(automation);

                    //Import Configuration Window
                    GetImportConfigurationWindowControls(automation);

                    #region Online
                    ////Automatic Injection Test View Window
                    //GetAutomaticInjectionTestViewWindowControls(automation);

                    ////Tunable Settings Window
                    //GetTunableSettingsControls(automation);

                    ////Export Data Window
                    //GetExportDataControls(automation); 
                    #endregion

                    #region Online
                    ////EMS Point Testing Window
                    //GetEMSPointTestingControls(automation);

                    ////System Availablity Window
                    //GetSystemAvailabilityControls(automation); 
                    #endregion

                    //Preferences Window
                    GetPreferencesControls(automation);

                    ////Quiet Time Configuration Window
                    //GetQuietTimeConfigurationControls(automation);

                    //Import Key File Window
                    GetImportKeyFileControls(automation);

                    #region Extra
                    ////About Window
                    //GetAboutControls(automation);

                    ////Help Window
                    //GetHelpControls(automation); 
                    #endregion
                }
                if (File.Exists(csvlogsPath))
                {
                    File.Delete(csvlogsPath);
                }
                Loggers.LogCsv(csvlogsPath, resultChild);
                Thread.Sleep(100);
                //KillApplication();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }
        
        public static void GetLoginWindowControls(UIA3Automation automation)
        {
            try
            {
                window = app.GetMainWindow(automation);
                var loginWindow = window.FindAll(TreeScope.Descendants, TrueCondition.Default);
                if (loginWindow != null)
                {

                    foreach (var childLogin in loginWindow)
                    {
                       
                        automationId = "";
                        name = "";
                        controltype = "";
                        parent = "";
                        x = "";
                        y = "";
                        width = "";
                        height = "";
                        result = "";

                        if (childLogin.ControlType == ControlType.ComboBox)
                        {
                            GetAllComboboxValuesByAutomationId(automation, childLogin.AutomationId, window.Name);
                        }

                        //AutomationId
                        try
                        {
                            automationId = childLogin.AutomationId;
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            automationId = "";
                        }
                        //Name
                        try
                        {
                            name = childLogin.Name;
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            name = "";
                        }
                        //ControlType
                        try
                        {
                            controltype = childLogin.ControlType.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            controltype = "";
                        }
                        //Control Location
                        try
                        {
                            x = childLogin.BoundingRectangle.X.ToString();
                            y = childLogin.BoundingRectangle.Y.ToString();
                            width = childLogin.BoundingRectangle.Width.ToString();
                            height = childLogin.BoundingRectangle.Height.ToString();

                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";

                        }

                        result = $"{Environment.NewLine}{window.Name},{automationId},{name}," +
                                 $"{controltype}," +
                                 $"{x},{y}," +
                                 $"{width},{height}";
                        resultChild.Add(result);
                    }
                }
                else
                {
                    resultChild.Add($"{Environment.NewLine}{window.Name} Screen's Control Not Found");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        public static void LoginToSmartInterface(UIA3Automation automation)
        {
            try
            {
                window = app.GetMainWindow(automation);
                SelectComboboxValueByAutomationId(automation, "cmbMode", /*"Online Mode"*/"Configuration Mode");
                SetEditBoxValueByAutomationId(automation, "txtUserName", "ATM");
                SetEditBoxValueByAutomationId(automation, "txtPwd", "P@s@1@9#a#t#m");
                SelectComboboxValueByAutomationId(automation, "cmbAuthenticationMode", "Local Authentication(SmartInterface)");
                InvokeButtonById(automation, "btnLogin");
                Thread.Sleep(1000);
                window = app.GetMainWindow(automation);
                Thread.Sleep(1000);
                var warningMessage = window.FindFirstDescendant(z => z.ByName("OK"));
                if (warningMessage != null) { warningMessage.Click(); }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Could not login to SmartInterface " + ex.Message);
            }
        }

        public static void GetSystemConfigurationWindowControls(UIA3Automation automation)
        {
            try
            {
                //resultChild.Add($"{Environment.NewLine}{window.Name}");
                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Online Mode")
                {
                    var systemConfiguration = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ConfigButtonUC.ToDescriptionString()));
                    systemConfiguration.Click();
                    window = app.GetMainWindow(automation);
                    Thread.Sleep(100);
                    var systemConfigWIndow = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SystemConfiguration.ToDescriptionString()));
                    var Points = systemConfigWIndow.GetClickablePoint();
                    Points.X = Points.X - 30;
                    Points.Y = Points.Y - 7;
                    FlaUI.Core.Input.Mouse.Click(Points, FlaUI.Core.Input.MouseButton.Left);
                }

                window = app.GetMainWindow(automation);
                Thread.Sleep(500);

                var systemConfigWin = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SystemConfiguration.ToDescriptionString()));
                System.Drawing.Point Point = new System.Drawing.Point();
                Point.X = 1073;
                Point.Y = 292;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Right);
                Thread.Sleep(1000);
                Point.X = 1124;
                Point.Y = 304;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                Thread.Sleep(500);

                // Circuit Screen 
                GetCircuitScreenControls(automation);
                //Gateway Screen
                GetGatewayScreenControls(automation);
                // PLC Screen 
                GetPLCScreenControls(automation);
                //Device Screen
                GetDeviceScreenControls(automation);

                var childerns = window.FindAll(TreeScope.Subtree, TrueCondition.Default);

                if (childerns != null)
                {

                    foreach (var child in childerns)
                    {
                        automationId = "";
                        name = "";
                        controltype = "";
                        parent = "";
                        x = "";
                        y = "";
                        width = "";
                        height = "";
                        result = "";

                        //Toggle Checkboxex
                        if (child.ControlType == ControlType.CheckBox)
                        {

                            CheckCheckboxValueByAutomationId(automation, child.AutomationId, window.Name);
                        }

                        //AutomationId
                        try
                        {
                            automationId = child.AutomationId;
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            automationId = "";
                        }
                        //Name
                        try
                        {
                            name = child.Name;
                            if (child.Name is null)
                            {
                                name = "";
                            }
                            if (name.Contains("\n"))
                            {
                                name = name.Replace("\r\n", String.Empty);
                                Thread.Sleep(1000);
                            }
                            if (name.Contains("No Circuit Present"))
                            {
                                name = name.Replace("No Circuit Present, Add a Circuit to Continue", "No Circuit Present Add a Circuit to Continue");
                                Thread.Sleep(1000);
                            }
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            name = "";
                        }
                        //ControlType
                        try
                        {
                            controltype = child.ControlType.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            controltype = "";
                        }
                        //Control Location
                        try
                        {
                            x = child.BoundingRectangle.X.ToString();
                            y = child.BoundingRectangle.Y.ToString();
                            width = child.BoundingRectangle.Width.ToString();
                            height = child.BoundingRectangle.Height.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                        }

                        result = $"{Environment.NewLine}{window.Name},{automationId},{name}," +
                                 $"{controltype}," +
                                 $"{x},{y}," +
                                 $"{width},{height}";
                        if (name.Contains("\n"))
                        {
                            name = name.Replace("\r\n", String.Empty);
                            Thread.Sleep(1000);
                        }
                        resultChild.Add(result);
                        resultChild = removeDuplicates(resultChild);                      
                    }
                    resultChild = removeDuplicates(resultChild);
                }
                else
                {
                    resultChild.Add($"{Environment.NewLine}{window.Name} Screen's Control Not Found");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }        

        public static void GetSiteStatusWindowControls(UIA3Automation automation)
        {
            try
            {
                //resultChild.Add($"{Environment.NewLine}{window.Name}");
                window = app.GetMainWindow(automation);
                var dashBoardButton = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.DashboardButtonUC.ToDescriptionString()));
                dashBoardButton.Click();
                window = app.GetMainWindow(automation);
                var siteStatus = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SiteStatus.ToDescriptionString()));
                var Point = siteStatus.GetClickablePoint();
                Point.X = Point.X - 30;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);

               
                Thread.Sleep(500);

                var wins = window.FindAllChildren();
                var siteStatusPane = wins[2];
                var siteStatusWindow = siteStatusPane.FindFirstDescendant(cf => cf.ByName("Site Status")).AsWindow();

                var windowChild = siteStatusWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

                #region Add two window controls sample
                //List<AutomationElement> windowChild = new List<AutomationElement>();
                //windowChild = siteStatusWindow.FindAll(TreeScope.Descendants, TrueCondition.Default).ToList();

                //List<AutomationElement> windowChild2 = new List<AutomationElement>();
                //windowChild2 = siteStatusWindow.FindAll(TreeScope.Descendants, TrueCondition.Default).ToList();


                //windowChild.AddRange(windowChild2); 
                #endregion

                if (windowChild != null)
                {
                    foreach (var subChildSiteStatus in windowChild)
                    {
                        if (subChildSiteStatus.ControlType == ControlType.ComboBox)
                        {
                            GetAllComboboxValuesByAutomationId(automation, subChildSiteStatus.AutomationId, siteStatusWindow.Name);
                        }

                        automationId = "";
                        name = "";
                        controltype = "";
                        parent = "";
                        x = "";
                        y = "";
                        width = "";
                        height = "";
                        result = "";
                        //AutomationId
                        try { automationId = subChildSiteStatus.AutomationId; }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                        //Name
                        try 
                        { 
                            name = subChildSiteStatus.Name;
                            if (subChildSiteStatus.Name is null)
                            {
                                name = "";
                            }
                            if (name.Contains("\n"))
                            {
                                name = name.Replace("\r\n", String.Empty);
                                Thread.Sleep(1000);
                            }
                            if (name.Contains("No Circuit Present"))
                            {
                                name = name.Replace("No Circuit Present, Add a Circuit to Continue", "No Circuit Present Add a Circuit to Continue");
                                Thread.Sleep(1000);
                            }
                            if (name.Contains("The Network Configuration process failed for some device(s)"))
                            {
                                name = name.Replace("The Network Configuration process failed for some device(s), please use the Re-apply Configuration button to apply the same configuration again.", "The Network Configuration process failed for some device(s) please use the Re-apply Configuration button to apply the same configuration again.");
                                Thread.Sleep(1000);
                            }
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                        //ControlType
                        try { controltype = subChildSiteStatus.ControlType.ToString(); }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                        //Control Location
                        try
                        {
                            x = subChildSiteStatus.BoundingRectangle.X.ToString();
                            y = subChildSiteStatus.BoundingRectangle.Y.ToString();
                            width = subChildSiteStatus.BoundingRectangle.Width.ToString();
                            height = subChildSiteStatus.BoundingRectangle.Height.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                        }
                        result = $"{Environment.NewLine}{siteStatusWindow.Name},{automationId},{name}," +
                                 $"{controltype}," +
                                 $"{x},{y}," +
                                 $"{width},{height}";
                        resultChild.Add(result);
                        resultChild = removeDuplicates(resultChild);
                    }
                }
                else { resultChild.Add($"{Environment.NewLine}{siteStatusWindow.Name} Screen Controls Not Found"); }

                window = app.GetMainWindow(automation);
                var pane = window.FindFirstDescendant(cf => cf.ByName("label3")).AsWindow();
                var paneChild = pane.FindAll(TreeScope.Descendants, TrueCondition.Default);
                if (paneChild != null)
                {
                    foreach (var subChildSiteStatus in paneChild)
                    {
                        automationId = "";
                        name = "";
                        controltype = "";
                        parent = "";
                        x = "";
                        y = "";
                        width = "";
                        height = "";
                        result = "";
                        //AutomationId
                        try { automationId = subChildSiteStatus.AutomationId; }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                        //Name
                        try
                        {
                            name = subChildSiteStatus.Name;

                            if (name.Contains("\n"))
                            {
                                name = name.Replace("\r\n", String.Empty);
                                Thread.Sleep(1000);
                            }
                            if (name.Contains("No Circuit Present"))
                            {
                                name = name.Replace("No Circuit Present, Add a Circuit to Continue", "No Circuit Present Add a Circuit to Continue");
                                Thread.Sleep(1000);
                            }
                            if (name.Contains("The Network Configuration process failed for some device(s)"))
                            {
                                name = name.Replace("The Network Configuration process failed for some device(s), please use the Re-apply Configuration button to apply the same configuration again.", "The Network Configuration process failed for some device(s) please use the Re-apply Configuration button to apply the same configuration again.");
                                Thread.Sleep(1000);
                            }
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                        //ControlType
                        try { controltype = subChildSiteStatus.ControlType.ToString(); }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                        //Control Location
                        try
                        {
                            x = subChildSiteStatus.BoundingRectangle.X.ToString();
                            y = subChildSiteStatus.BoundingRectangle.Y.ToString();
                            width = subChildSiteStatus.BoundingRectangle.Width.ToString();
                            height = subChildSiteStatus.BoundingRectangle.Height.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                        }
                        result = $"{Environment.NewLine}{siteStatusWindow.Name},{automationId},{name}," +
                                 $"{controltype}," +
                                 $"{x},{y}," +
                                 $"{width},{height}";
                        resultChild.Add(result);
                        resultChild = removeDuplicates(resultChild);
                    }
                } 
                else
                {
                    resultChild.Add($"{Environment.NewLine}{siteStatusWindow.Name} Screen Controls Not Found");
                }
                dashBoardButton.Click(true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }

        public static void GetAlarmConfigurationWindowControls(UIA3Automation automation)
        {
            try
            {
                //resultChild.Add($"{Environment.NewLine}{window.Name}");
                var configButton = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ConfigButtonUC.ToDescriptionString()));
                Thread.Sleep(100);
                configButton.Click();
                window = app.GetMainWindow(automation);
                var alarmsConfiguration = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.AlarmsConfiguration.ToDescriptionString()));
                Thread.Sleep(100);
                var Point = alarmsConfiguration.GetClickablePoint();
                Point.Y = Point.Y - 10;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                Thread.Sleep(100);

                try
                {
                    var wins = window.FindAllChildren();
                    var alarmConfigPane = wins[2];
                    var paneChild = alarmConfigPane.FindAllChildren();
                    alarmConfigPane = paneChild[0];
                    alarmsConfiguration = alarmConfigPane.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                    var panel = alarmsConfiguration.FindFirstDescendant(z => z.ByAutomationId("panel1"));
                    var getAlarmsConfigWindow = panel.FindAll(TreeScope.Descendants, TrueCondition.Default);


                    var alarmConfigWindow = alarmConfigPane.FindFirstDescendant(cf => cf.ByName("Alarms Configuration")).AsWindow();

                    var windowChild = alarmConfigWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

                    window = app.GetMainWindow(automation);
                    //var alarmConfigPane = window.FindFirstDescendant(cf => cf.ByAutomationId("33361746"));
                    //var alarmConfigWindow = alarmConfigPane.FindFirstDescendant(cf => cf.ByName("Alarms Configuration")).AsWindow();
                    alarmsConfiguration = alarmConfigWindow.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                    
                    if (alarmsConfiguration != null)
                    {
                        //var getAlarmsConfigWindow = window.FindAll(TreeScope.Descendants, TrueCondition.Default);
                        var Pane = alarmsConfiguration.FindFirstDescendant(cf => cf.ByName("Alarms Configuration"));
                        var alarmDataGrid = Pane.FindFirstDescendant(cf => cf.ByName("grdAlarmView")).AsDataGridView();
                        getAlarmsConfigWindow = window.FindAll(TreeScope.Descendants, TrueCondition.Default);
                        //if (alarmDataGrid.ControlType == ControlType.CheckBox)
                        //{
                            

                        //}
                        
                        if (getAlarmsConfigWindow != null)
                        {
                            //var childsAlarmConfiguration = getAlarmsConfigWindow.FindAllDescendants();
                            foreach (var childAlarmsConfig in getAlarmsConfigWindow)
                            {
                                automationId = "";
                                name = "";
                                controltype = "";
                                parent = "";
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                                result = "";

                                //AutomationId
                                try { automationId = childAlarmsConfig.AutomationId; }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                                //Name
                                try { name = childAlarmsConfig.Name; }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                                //ControlType
                                try { controltype = childAlarmsConfig.ControlType.ToString(); }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                                //Control Location
                                try
                                {
                                    x = childAlarmsConfig.BoundingRectangle.X.ToString();
                                    y = childAlarmsConfig.BoundingRectangle.Y.ToString();
                                    width = childAlarmsConfig.BoundingRectangle.Width.ToString();
                                    height = childAlarmsConfig.BoundingRectangle.Height.ToString();
                                }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                {
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                }

                                result = $"{Environment.NewLine}{alarmConfigWindow.Name},{automationId},{name}," +
                                         $"{controltype}," +
                                         $"{x},{y}," +
                                         $"{width},{height}";
                                resultChild.Add(result);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    resultChild.Add($"{Environment.NewLine}Alarms Configuration Screen's Control Not Found");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }

        public static void GetPhaseNamingSchemeWindowControls(UIA3Automation automation)
        {
            try
            {
                //resultChild.Add($"{Environment.NewLine}Phase Naming Scheme");
                var configButton = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ConfigButtonUC.ToDescriptionString()));
                Thread.Sleep(100);
                configButton.Click();
                Thread.Sleep(500);
                System.Drawing.Point Point = new System.Drawing.Point();
                Point.X = 119;
                Point.Y = 280;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                Thread.Sleep(500);
                //window = app.GetMainWindow(automation);
                //var phaseNamingSchemeClick = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.PhaseNamingScheme.ToDescriptionString()));
                //var Point = phaseNamingSchemeClick.GetClickablePoint();
                //Point.Y = Point.Y - 10;
                //FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                //Thread.Sleep(100);

                window = app.GetMainWindow(automation);
                var getPhaseNamingSchemeWindow = window.FindFirstChild(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.PhaseNamingSchemeForm.ToDescriptionString()));

                //var dataGrid = getPhaseNamingSchemeWindow.FindFirstDescendant(cf => cf.ByName("grdPhaseNamingScheme")).AsDataGridView();               

                //var DataGrid = getPhaseNamingSchemeWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

                int noOfRows = 0;
                DesktopAuto.NGetMainWindowbyName(getPhaseNamingSchemeWindow.Name);
                DesktopAuto.NFindControlByAutomationId(6, "grdPhaseNamingScheme");
                DesktopAuto.NGetAllDataGridValues(14, out noOfRows);
                if (noOfRows > 0)
                {
                    var data = DesktopAuto.DataGridData;
                   
                    for (int i = 0; i < noOfRows; i++)
                    {
                        for (int j = 0; j < 14; j++)
                        {
                            result = $"{Environment.NewLine}{getPhaseNamingSchemeWindow.Name}," + $"{i}.{j}," + $"{data[i, j]}," +
                                                         $"DataGrid Cell Data," +
                                                         "0," + "0," +
                                                         "0," + "0";
                            resultChild.Add(result);
                        }
                    }
                }

                var CancelButton = getPhaseNamingSchemeWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnBack")).AsButton();
                var childsPhaseNamingScheme = getPhaseNamingSchemeWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);
                if (childsPhaseNamingScheme != null)
                {
                    foreach (var childPhNScheme in childsPhaseNamingScheme)
                    {
                        automationId = "";
                        name = "";
                        controltype = "";
                        parent = "";
                        x = "";
                        y = "";
                        width = "";
                        height = "";
                        result = "";

                        //AutomationId
                        try { automationId = childPhNScheme.AutomationId; }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                        //Name
                        try { name = childPhNScheme.Name; }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                        //ControlType
                        try { controltype = childPhNScheme.ControlType.ToString(); }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                        //Control Location
                        try
                        {
                            x = childPhNScheme.BoundingRectangle.X.ToString();
                            y = childPhNScheme.BoundingRectangle.Y.ToString();
                            width = childPhNScheme.BoundingRectangle.Width.ToString();
                            height = childPhNScheme.BoundingRectangle.Height.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                        }

                        result = $"{Environment.NewLine}{getPhaseNamingSchemeWindow.Name},{automationId},{name}," +
                                 $"{controltype}," +
                                 $"{x},{y}," +
                                 $"{width},{height}";
                        resultChild.Add(result);
                    }
                }
                else
                {
                    resultChild.Add($"{Environment.NewLine}{getPhaseNamingSchemeWindow.Name} Screen Controls Not Found");
                }
                CancelButton.Click(true);
                Thread.Sleep(100);
                configButton.Click(true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }
        //Online Screen
        public static void GetCircuitStatusWindowControls(UIA3Automation automation)
        {
            try
            {
                resultChild.Add($"{Environment.NewLine}Circuit Status");
                var statAndCmdBtn = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.StatusCommandsButtonUC.ToDescriptionString()));
                if (statAndCmdBtn != null)
                {
                    Thread.Sleep(100);
                    statAndCmdBtn.Click();
                    var circuitStatusbtn = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.CircuitStatus.ToDescriptionString()));
                    if (circuitStatusbtn != null)
                    {
                        Thread.Sleep(100);
                        var Point = circuitStatusbtn.GetClickablePoint();
                        Point.Y = Point.Y - 10;
                        Point.X = Point.X - 25;
                        FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                        Thread.Sleep(100);
                        window = app.GetMainWindow(automation);
                        var temp10 = window.FindFirstDescendant(z => z.ByAutomationId("pnlMain"));
                        if (temp10 != null)
                        {
                            var circuitStatusWindow = temp10.FindAll(TreeScope.Descendants, TrueCondition.Default);
                            if (circuitStatusWindow != null)
                            {
                                foreach (var childCircuitStatus in circuitStatusWindow)
                                {
                                    automationId = "";
                                    name = "";
                                    controltype = "";
                                    parent = "";
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                    result = "";

                                    //AutomationId
                                    try { automationId = childCircuitStatus.AutomationId; }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                                    //Name
                                    try { name = childCircuitStatus.Name; }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                                    //ControlType
                                    try { controltype = childCircuitStatus.ControlType.ToString(); }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                                    //Control Location
                                    try
                                    {
                                        x = childCircuitStatus.BoundingRectangle.X.ToString();
                                        y = childCircuitStatus.BoundingRectangle.Y.ToString();
                                        width = childCircuitStatus.BoundingRectangle.Width.ToString();
                                        height = childCircuitStatus.BoundingRectangle.Height.ToString();
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        x = "";
                                        y = "";
                                        width = "";
                                        height = "";
                                    }

                                    result = $"{Environment.NewLine}Circuit Status,{automationId},{name}," +
                                             $"{controltype}," +
                                             $"{x},{y}," +
                                             $"{width},{height}";
                                    resultChild.Add(result);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }
        //Online Screen
        public static void GetAlarmsStatusWindowControls(UIA3Automation automation)
        {
            try
            {
                resultChild.Add($"{Environment.NewLine}Alarm Status");
                window = app.GetMainWindow(automation);
                var alarmStatusbtn = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.AlarmStatus.ToDescriptionString()));
                if (alarmStatusbtn != null)
                {
                    Thread.Sleep(100);
                    var Point = alarmStatusbtn.GetClickablePoint();
                    Point.Y = Point.Y - 10;
                    Point.X = Point.X - 25;
                    FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                    Thread.Sleep(100);
                    window = app.GetMainWindow(automation);
                    var temp11 = window.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                    if (temp11 != null)
                    {
                        var alarmsStatusWindow = temp11.FindAll(TreeScope.Descendants, TrueCondition.Default);
                        if (alarmsStatusWindow != null)
                        {
                            //var childsAlarmConfiguration = getAlarmsConfigWindow.FindAllDescendants();
                            foreach (var childAlarmStatus in alarmsStatusWindow)
                            {
                                automationId = "";
                                name = "";
                                controltype = "";
                                parent = "";
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                                result = "";

                                //AutomationId
                                try { automationId = childAlarmStatus.AutomationId; }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                                //Name
                                try { name = childAlarmStatus.Name; }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                                //ControlType
                                try { controltype = childAlarmStatus.ControlType.ToString(); }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                                //Control Location
                                try
                                {
                                    x = childAlarmStatus.BoundingRectangle.X.ToString();
                                    y = childAlarmStatus.BoundingRectangle.Y.ToString();
                                    width = childAlarmStatus.BoundingRectangle.Width.ToString();
                                    height = childAlarmStatus.BoundingRectangle.Height.ToString();
                                }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                {
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                }

                                result = $"{Environment.NewLine}Alarm Status,{automationId},{name}," +
                                         $"{controltype}," +
                                         $"{x},{y}," +
                                         $"{width},{height}";
                                resultChild.Add(result);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }

        }
        //Online Screen
        public static void GetStructureSummaryWindowControls(UIA3Automation automation)
        {
            try
            {
                resultChild.Add($"{Environment.NewLine}Structure Summary");
                var structureSummary = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.StructureSummary.ToDescriptionString()));
                if (structureSummary != null)
                {
                    Thread.Sleep(100);
                    var Point = structureSummary.GetClickablePoint();
                    Point.Y = Point.Y - 10;
                    Point.X = Point.X - 25;
                    FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                    Thread.Sleep(100);
                    window = app.GetMainWindow(automation);
                    var temp12 = window.FindFirstDescendant(z => z.ByAutomationId("StructureSummary"));
                    if (temp12 != null)
                    {
                        var structureSummaryWindow = temp12.FindAll(TreeScope.Descendants, TrueCondition.Default);
                        if (structureSummaryWindow != null)
                        {
                            //var childsAlarmConfiguration = getAlarmsConfigWindow.FindAllDescendants();
                            foreach (var childStructureSummary in structureSummaryWindow)
                            {
                                automationId = "";
                                name = "";
                                controltype = "";
                                parent = "";
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                                result = "";

                                //AutomationId
                                try { automationId = childStructureSummary.AutomationId; }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                                //Name
                                try { name = childStructureSummary.Name; }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                                //ControlType
                                try { controltype = childStructureSummary.ControlType.ToString(); }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                                //Control Location
                                try
                                {
                                    x = childStructureSummary.BoundingRectangle.X.ToString();
                                    y = childStructureSummary.BoundingRectangle.Y.ToString();
                                    width = childStructureSummary.BoundingRectangle.Width.ToString();
                                    height = childStructureSummary.BoundingRectangle.Height.ToString();
                                }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                {
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                }

                                result = $"{Environment.NewLine}Structure Summary,{automationId},{name}," +
                                         $"{controltype}," +
                                         $"{x},{y}," +
                                         $"{width},{height}";
                                resultChild.Add(result);
                            }

                            structureSummary = window.FindFirstDescendant(z => z.ByAutomationId("btnCancel"));
                            structureSummary.Click();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }

        public static void GetReportsWindowControls(UIA3Automation automation)
        {
            try
            {
                
                var reportsButton = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ReportsButtonUC.ToDescriptionString()));
                reportsButton.Click();                               

                // Reports Telemetry Contrls
                GetReportsTelemetryControls(automation);

                #region Dummy
                //GetReportsAlarmControls(automation);
                // Reports Alarm controls
                //GetReportsAlarmControls(automation);

                //var childsReports = reportsTelemetry.FindAll(TreeScope.Descendants, TrueCondition.Default);
                //foreach (var subChildReport in childsReports)
                //{
                //    if (subChildReport.ControlType == ControlType.ComboBox)
                //    {
                //        GetAllComboboxValuesByAutomationId(automation, subChildReport.AutomationId, window.Name);
                //    }

                //    automationId = "";
                //    name = "";
                //    controltype = "";
                //    parent = "";
                //    x = "";
                //    y = "";
                //    width = "";
                //    height = "";
                //    result = "";
                //    ///AutomationId
                //    try { automationId = subChildReport.AutomationId; }
                //    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                //    //Name
                //    try { name = subChildReport.Name; }
                //    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                //    //ControlType
                //    try { controltype = subChildReport.ControlType.ToString(); }
                //    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                //    //Control Location
                //    try
                //    {
                //        x = subChildReport.BoundingRectangle.X.ToString();
                //        y = subChildReport.BoundingRectangle.Y.ToString();
                //        width = subChildReport.BoundingRectangle.Width.ToString();
                //        height = subChildReport.BoundingRectangle.Height.ToString();
                //    }
                //    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                //    {
                //        x = "";
                //        y = "";
                //        width = "";
                //        height = "";

                //    }

                //    result = $"{Environment.NewLine}Reports,{automationId},{name}," +
                //             $"{controltype}," +
                //             $"{x},{y}," +
                //             $"{width},{height}";
                //    resultChild.Add(result);
                //} 
                #endregion
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }

        public static void GetSwitchToOnlineOrConfigurationWindowControls(UIA3Automation automation)
        {
            try
            {
                resultChild.Add($"{Environment.NewLine}Switch To Online/Offline");

                var managementButton = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                managementButton.Click();
                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Online Mode")
                {
                    var switchToOnline = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SwitchtoConfiguration.ToDescriptionString()));
                    var Point = switchToOnline.GetClickablePoint();
                    Point.Y = Point.Y - 10;
                    FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                    Thread.Sleep(100);
                    window = app.GetMainWindow(automation);

                    //var temp = window.FindFirstChild().FindFirstDescendant(z => z.ByControlType(FlaUI.Core.Definitions.ControlType.TitleBar));
                    var temp = window.FindFirstDescendant(z => z.ByControlType(FlaUI.Core.Definitions.ControlType.TitleBar));
                    //var temp = window.AsTitleBar();
                    if (temp != null)
                    {
                        var childsSwToOnline = temp.FindAll(TreeScope.Descendants, TrueCondition.Default);

                        foreach (var subChildSwToOnline in childsSwToOnline)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try { automationId = subChildSwToOnline.AutomationId; }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                            //Name
                            try { name = subChildSwToOnline.Name; }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                            //ControlType
                            try { controltype = subChildSwToOnline.ControlType.ToString(); }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                            //Control Location
                            try
                            {
                                x = subChildSwToOnline.BoundingRectangle.X.ToString();
                                y = subChildSwToOnline.BoundingRectangle.Y.ToString();
                                width = subChildSwToOnline.BoundingRectangle.Width.ToString();
                                height = subChildSwToOnline.BoundingRectangle.Height.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                            }

                            result = $"{Environment.NewLine}Switch To Online/Offline,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                        }

                        switchToOnline = window.FindFirstDescendant(z => z.ByName("No"));
                        switchToOnline.Click();
                    }
                }
                else if (window.Name == "SmartInterface - Configuration Mode")
                {
                    var switchToConfiguration = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SwitchtoOnline.ToDescriptionString()));
                    var Point = switchToConfiguration.GetClickablePoint();
                    Point.Y = Point.Y - 10;
                    FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                    Thread.Sleep(100);
                    window = app.GetMainWindow(automation);
                    var temp = window.FindFirstDescendant(z => z.ByControlType(FlaUI.Core.Definitions.ControlType.TitleBar));
                    //var temp = window.AsTitleBar();
                    if (temp != null)
                    {
                        var childsSwToOnline = temp.FindAll(TreeScope.Descendants, TrueCondition.Default);

                        foreach (var subChildSwToOnline in childsSwToOnline)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try
                            { automationId = subChildSwToOnline.AutomationId; }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                            //Name
                            try { name = subChildSwToOnline.Name; }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                            //ControlType
                            try { controltype = subChildSwToOnline.ControlType.ToString(); }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                            //Control Location
                            try
                            {
                                x = subChildSwToOnline.BoundingRectangle.X.ToString();
                                y = subChildSwToOnline.BoundingRectangle.Y.ToString();
                                width = subChildSwToOnline.BoundingRectangle.Width.ToString();
                                height = subChildSwToOnline.BoundingRectangle.Height.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                            }

                            result = $"{Environment.NewLine}SmartInterface - Configuration Mode,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                        }

                        switchToConfiguration = window.FindFirstDescendant(z => z.ByName("No"));
                        switchToConfiguration.Click();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }

        public static void GetUserManagementWindowControls(UIA3Automation automation)
        {
            try
            {
                var window = app.GetMainWindow(automation);
                //resultChild.Add($"{Environment.NewLine}User Management");
                var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                management.Click();
                window = app.GetMainWindow(automation);
                var userManagement = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.UserManagement.ToDescriptionString()));
                var Point = userManagement.GetClickablePoint();
                Point.Y = Point.Y - 10;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                
                // Add/Edit User Window Contrls
                GetAddUserScreenControls(automation);
                // Active Directory Server Window Controls
                GetAddLDPAUserScreenControls(automation);
                // Reset Password Window Controls
                GetResetPasswordScreenControls(automation);
                // Account Policy Window Controls
                GetAccountPolicyScreenControls(automation);
                // User Roles Window Controls
                GetUserRolesScreenControls(automation);

                window = app.GetMainWindow(automation);

                var userManagementWind = window.FindFirstDescendant(z => z.ByAutomationId("UserManagement")).AsWindow();
                var temp2 = userManagementWind.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));

                 int noOfRows = 0;
                DesktopAuto.NGetMainWindowbyName(userManagementWind.Name);
                DesktopAuto.NFindControlByAutomationId(6, "lstUsers");
                DesktopAuto.NGetAllDataGridValues(14, out noOfRows);
                if (noOfRows > 0)
                {
                    var data = DesktopAuto.DataGridData;
                   
                    for (int i = 0; i < noOfRows; i++)
                    {
                        for (int j = 0; j < 14; j++)
                        {
                            result = $"{Environment.NewLine}{userManagementWind.Name}," + $"{i}.{j}," + $"{data[i, j]}," +
                                                         $"DataGrid Cell Data," +
                                                         "0," + "0," +
                                                         "0," + "0";
                            resultChild.Add(result);
                        }
                    }
                }

                if (temp2 != null)
                {
                    var childsUserManag = temp2.FindAll(TreeScope.Descendants, TrueCondition.Default);
                    if (childsUserManag != null)
                    {
                        foreach (var subChildUserManag in childsUserManag)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try
                            {
                                automationId = subChildUserManag.AutomationId;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                automationId = "";
                            }
                            //Name
                            try
                            {
                                name = subChildUserManag.Name;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                name = "";
                            }
                            //ControlType
                            try
                            {
                                controltype = subChildUserManag.ControlType.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                controltype = "";
                            }
                            //Control Location
                            try
                            {
                                x = subChildUserManag.BoundingRectangle.X.ToString();
                                y = subChildUserManag.BoundingRectangle.Y.ToString();
                                width = subChildUserManag.BoundingRectangle.Width.ToString();
                                height = subChildUserManag.BoundingRectangle.Height.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";

                            }

                            result = $"{Environment.NewLine}{userManagementWind.Name},{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                            resultChild = removeDuplicates(resultChild);
                        }
                        resultChild = removeDuplicates(resultChild);
                        management.Click(true);
                    }
                    else
                    {
                        resultChild.Add($"{Environment.NewLine}{userManagementWind.Name} Screen Controls Not Found");
                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }

        public static void GetFirmwareUpdateWindowControls(UIA3Automation automation)
        {
            #region Frimware Update
            try
            {
                //resultChild.Add($"{Environment.NewLine}Firmware Update");
                var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                management.Click();
                window = app.GetMainWindow(automation);
                var firmwareUpdate = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.FirmwareUpdate.ToDescriptionString()));
                var Point = firmwareUpdate.GetClickablePoint();
                Point.Y = Point.Y - 10;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);

                window = app.GetMainWindow(automation);
                var temp1 = window.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel"));
                if (temp1 != null)
                {
                    var childsfrmUpdt = temp1.FindAll(TreeScope.Descendants, TrueCondition.Default);
                    if (childsfrmUpdt != null)
                    {
                        foreach (var subChildfrmUpdt in childsfrmUpdt)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try
                            {
                                automationId = subChildfrmUpdt.AutomationId;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                automationId = "";
                            }
                            //Name
                            try
                            {
                                name = subChildfrmUpdt.Name;
                                if (name.Contains("PowerLine \r\nCoordinator(s)"))
                                {
                                    name = name.Replace("PowerLine \r\nCoordinator(s)", "PowerLine Coordinator(s)");
                                    Thread.Sleep(1000);
                                }
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                name = "";
                            }
                            //ControlType
                            try
                            {
                                controltype = subChildfrmUpdt.ControlType.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                controltype = "";
                            }
                            //Control Location
                            try
                            {
                                x = subChildfrmUpdt.BoundingRectangle.X.ToString();
                                y = subChildfrmUpdt.BoundingRectangle.Y.ToString();
                                width = subChildfrmUpdt.BoundingRectangle.Width.ToString();
                                height = subChildfrmUpdt.BoundingRectangle.Height.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";

                            }

                            result = $"{Environment.NewLine}Firmware Update,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                            resultChild = removeDuplicates(resultChild);
                        }
                        resultChild = removeDuplicates(resultChild);
                        management.Click(true);
                    }
                    else
                    {
                        resultChild.Add($"{Environment.NewLine}Firmware Update Screen Controls Not Found");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }

            #endregion
        }
        public static void GetExportConfigurationWindowControls(UIA3Automation automation)
        {
            #region Export Configuration
            try
            {
                //resultChild.Add($"{Environment.NewLine}Export Configuration");
                var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                management.Click();
                window = app.GetMainWindow(automation);
                var exptConfig = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.ExportConfiguration.ToDescriptionString()));
                if (exptConfig != null)
                {
                    try
                    {
                        var Point = exptConfig.GetClickablePoint();
                        Point.Y = Point.Y - 10;
                        FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                        window = app.GetMainWindow(automation);
                        Thread.Sleep(100);
                        var temp5 = window.FindFirstDescendant(z => z.ByName("Save As"));
                        if (temp5 != null)
                        {
                            var childsExptConfig = temp5.FindAll(TreeScope.Descendants, TrueCondition.Default);
                            var cancelButtin = temp5.FindFirstDescendant(z => z.ByName("Cancel")).AsButton();
                            if (childsExptConfig != null)
                            {
                                foreach (var subChildExptConfig in childsExptConfig)
                                {
                                    automationId = "";
                                    name = "";
                                    controltype = "";
                                    parent = "";
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                    result = "";
                                    //AutomationId
                                    try
                                    {
                                        automationId = subChildExptConfig.AutomationId;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        automationId = "";
                                    }
                                    //Name
                                    try
                                    {
                                        name = subChildExptConfig.Name;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        name = "";
                                    }
                                    //ControlType
                                    try
                                    {
                                        controltype = subChildExptConfig.ControlType.ToString();
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        controltype = "";
                                    }
                                    //Control Location
                                    try
                                    {
                                        x = subChildExptConfig.BoundingRectangle.X.ToString();
                                        y = subChildExptConfig.BoundingRectangle.Y.ToString();
                                        width = subChildExptConfig.BoundingRectangle.Width.ToString();
                                        height = subChildExptConfig.BoundingRectangle.Height.ToString();
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        x = "";
                                        y = "";
                                        width = "";
                                        height = "";

                                    }

                                    result = $"{Environment.NewLine}Export Configuration,{automationId},{name}," +
                                             $"{controltype}," +
                                             $"{x},{y}," +
                                             $"{width},{height}";
                                    resultChild.Add(result);
                                    resultChild = removeDuplicates(resultChild);
                                }
                                resultChild = removeDuplicates(resultChild);
                            }
                            else
                            {
                                resultChild.Add($"{Environment.NewLine}Export Configuration Screen Controls Not Found");
                            }
                            Thread.Sleep(100);
                            cancelButtin.Click();
                            Thread.Sleep(100);
                            management.Click(true);
                        }
                        else
                        {
                            resultChild.Add($"{Environment.NewLine}Export Configuration Screen's Control Not Found");
                        }
                    }
                    catch (Exception)
                    {
                        resultChild.Add($"{Environment.NewLine}Export Configuration Screen's Control Not Found");
                    }
                }
                
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }

        public static void GetImportConfigurationWindowControls(UIA3Automation automation)
        {
            #region Import Configuration
            try
            {
                resultChild.Add($"{Environment.NewLine}Import Configuration");

                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Configuration Mode") //SmartInterface - Online Mode
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var imptConfig = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.ImportConfiguration.ToDescriptionString()));
                    var Point = imptConfig.GetClickablePoint();
                    Point.Y = Point.Y + 5;
                    FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);

                    window = app.GetMainWindow(automation);
                    Thread.Sleep(1000);
                    var temp6 = window.FindFirstDescendant(z => z.ByName("Open"));
                    Thread.Sleep(1000);
                    var childsImptConfig = temp6.FindAll(TreeScope.Descendants, TrueCondition.Default);
                    var cancelButton = temp6.FindFirstDescendant(z => z.ByName("Cancel")).AsButton();
                    if (childsImptConfig != null)
                    {
                        foreach (var subChildExptConfig in childsImptConfig)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try
                            {
                                automationId = subChildExptConfig.AutomationId;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                automationId = "";
                            }
                            //Name
                            try
                            {
                                name = subChildExptConfig.Name;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                name = "";
                            }
                            //ControlType
                            try
                            {
                                controltype = subChildExptConfig.ControlType.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                controltype = "";
                            }
                            //Control Location
                            try
                            {
                                x = subChildExptConfig.BoundingRectangle.X.ToString();
                                y = subChildExptConfig.BoundingRectangle.Y.ToString();
                                width = subChildExptConfig.BoundingRectangle.Width.ToString();
                                height = subChildExptConfig.BoundingRectangle.Height.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                            }

                            result = $"{Environment.NewLine}Import Configuration,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                            resultChild = removeDuplicates(resultChild);
                        }
                        resultChild = removeDuplicates(resultChild);
                    }
                    else
                    {
                        resultChild.Add($"{Environment.NewLine}Import Configuration Screen Controls Not Found");
                    }
                    Thread.Sleep(100);
                    cancelButton.Click();
                    Thread.Sleep(100);
                    management.Click(true);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }
        //Online Screen
        public static void GetAutomaticInjectionTestViewWindowControls(UIA3Automation automation)
        {
            #region Automatic Injection Test View
            try
            {
                //resultChild.Add($"{Environment.NewLine}Automatic Injection Test View");

                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Online Mode")
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var autoInjTstView = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.AutomaticInjectionTestView.ToDescriptionString()));
                    if (autoInjTstView != null)
                    {
                        var Point = autoInjTstView.GetClickablePoint();
                        Point.Y = Point.Y - 10;
                        Point.X = Point.X - 25;
                        FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                        window = app.GetMainWindow(automation);
                        try
                        {
                            autoInjTstView = window.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1")).AsWindow();
                            var childsAutoInjTstView = window.FindAll(TreeScope.Descendants, TrueCondition.Default);
                            if (childsAutoInjTstView != null)
                            {
                                foreach (var subChildAutoInjTstView in childsAutoInjTstView)
                                {
                                    automationId = "";
                                    name = "";
                                    controltype = "";
                                    parent = "";
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                    result = "";
                                    //AutomationId
                                    try
                                    {
                                        automationId = subChildAutoInjTstView.AutomationId;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        automationId = "";
                                    }
                                    //Name
                                    try
                                    {
                                        name = subChildAutoInjTstView.Name;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        name = "";
                                    }
                                    //ControlType
                                    try
                                    {
                                        controltype = subChildAutoInjTstView.ControlType.ToString();
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        controltype = "";
                                    }
                                    //Control Location
                                    try
                                    {
                                        x = subChildAutoInjTstView.BoundingRectangle.X.ToString();
                                        y = subChildAutoInjTstView.BoundingRectangle.Y.ToString();
                                        width = subChildAutoInjTstView.BoundingRectangle.Width.ToString();
                                        height = subChildAutoInjTstView.BoundingRectangle.Height.ToString();

                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        x = "";
                                        y = "";
                                        width = "";
                                        height = "";
                                    }

                                    result = $"{Environment.NewLine}SmartInterface - Online Mode,{automationId},{name}, " +
                                             $"{controltype}," +
                                             $"{x},{y}," +
                                             $"{width},{height}";
                                    resultChild.Add(result);
                                }
                                management.Click(true);
                            }
                        }
                        catch (Exception)
                        {
                            resultChild.Add("Automatic Injection Test View Screen's Control Not Found");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }
        //Online/Offline Both
        public static void GetTunableSettingsControls(UIA3Automation automation)
        {
            #region Tunable Settings
            try
            {
                //resultChild.Add($"{Environment.NewLine}Tunable Settings");
                var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                management.Click();
                window = app.GetMainWindow(automation);
                var tunSettings = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.TunableSettings.ToDescriptionString()).And(z.ByControlType(FlaUI.Core.Definitions.ControlType.Hyperlink)));
                var Point = tunSettings.GetClickablePoint();
                Point.Y = Point.Y - 10;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                Thread.Sleep(100);
                window = app.GetMainWindow(automation);
                var temp3 = window.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                if (temp3 != null)
                {
                    var childsTunSettings = temp3.FindAll(TreeScope.Descendants, TrueCondition.Default);
                    if (childsTunSettings != null)
                    {
                        foreach (var subChildTunSettings in childsTunSettings)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try
                            {
                                automationId = subChildTunSettings.AutomationId;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                automationId = "";
                            }
                            //Name
                            try
                            {
                                name = subChildTunSettings.Name;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                name = "";
                            }
                            //ControlType
                            try
                            {
                                controltype = subChildTunSettings.ControlType.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                controltype = "";
                            }
                            //Control Location
                            try
                            {
                                x = subChildTunSettings.BoundingRectangle.X.ToString();
                                y = subChildTunSettings.BoundingRectangle.Y.ToString();
                                width = subChildTunSettings.BoundingRectangle.Width.ToString();
                                height = subChildTunSettings.BoundingRectangle.Height.ToString();

                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                            }

                            result = $"{Environment.NewLine}Tunable Settings,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                        }
                        management.Click(true);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }
        //Online Screen
        public static void GetExportDataControls(UIA3Automation automation)
        {
            #region Export Data
            try
            {
                resultChild.Add($"{Environment.NewLine}Export Data");
                //ExportData
                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Online Mode")
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var expData = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.ExportData.ToDescriptionString()));
                    if (expData != null)
                    {
                        var Point = expData.GetClickablePoint();
                        //Point.Y = Point.Y - 10;
                        Point.X = Point.X - 25;
                        FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                        window = app.GetMainWindow(automation);
                        expData = window.FindFirstDescendant(z => z.ByAutomationId("ExportData"));
                        if (expData != null)
                        {
                            var childsExpData = expData.FindAll(TreeScope.Descendants, TrueCondition.Default);
                            if (childsExpData != null)
                            {
                                foreach (var subChildExpDatas in childsExpData)
                                {
                                    automationId = "";
                                    name = "";
                                    controltype = "";
                                    parent = "";
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                    result = "";
                                    //AutomationId
                                    try
                                    {
                                        automationId = subChildExpDatas.AutomationId;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        automationId = "";
                                    }
                                    //Name
                                    try
                                    {
                                        name = subChildExpDatas.Name;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        name = "";
                                    }
                                    //ControlType
                                    try
                                    {
                                        controltype = subChildExpDatas.ControlType.ToString();
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        controltype = "";
                                    }
                                    //Control Location
                                    try
                                    {
                                        x = subChildExpDatas.BoundingRectangle.X.ToString();
                                        y = subChildExpDatas.BoundingRectangle.Y.ToString();
                                        width = subChildExpDatas.BoundingRectangle.Width.ToString();
                                        height = subChildExpDatas.BoundingRectangle.Height.ToString();

                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        x = "";
                                        y = "";
                                        width = "";
                                        height = "";
                                    }

                                    result = $"{Environment.NewLine}Export Data,{automationId},{name}, " +
                                             $"{controltype}," +
                                             $"{x},{y}," +
                                             $"{width},{height}";
                                    resultChild.Add(result);
                                }
                                Point.X = 1090;
                                Point.Y = 378;
                                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                                management.Click(true);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }
        
        public static void GetPreferencesControls(UIA3Automation automation)
        {
            #region Preferences
            try
            {
                #region Shehroz
                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Configuration Mode")
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var preferences = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.Preferences.ToDescriptionString()).And(z.ByControlType(FlaUI.Core.Definitions.ControlType.Hyperlink)));
                    if (preferences != null)
                    {
                        var Point = preferences.GetClickablePoint();
                        Point.Y = Point.Y - 10;
                        FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                        Thread.Sleep(500);
                        //Get OverLayForm
                        var wins = app.GetAllTopLevelWindows(automation);
                        // Get Prefrences Screen
                        var overlayForm = wins[0];
                        var prefrencesWindow = overlayForm.FindFirstDescendant(cf => cf.ByName("Preferences")).AsWindow();
                        var PrefrencesControls = prefrencesWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

                        var CancelButton = prefrencesWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnBack")).AsButton();

                        if (PrefrencesControls != null)
                        {
                            foreach (var Controls in PrefrencesControls)
                            {
                                if (Controls.ControlType == ControlType.ComboBox)
                                {

                                    var combo = prefrencesWindow.FindFirstDescendant(cf => cf.ByAutomationId(Controls.AutomationId)).AsComboBox();
                                    if (combo != null)
                                    {
                                        combo.Click(true);
                                        var comboBoxValues = combo.FindAll(TreeScope.Descendants, TrueCondition.Default);
                                        foreach (var listItem in comboBoxValues)
                                        {
                                            if (listItem.ControlType == ControlType.ListItem)
                                            {

                                                //AutomationId
                                                try
                                                {
                                                    automationId = listItem.AutomationId;
                                                }
                                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                                {
                                                    automationId = "";
                                                }
                                                //Name
                                                try
                                                {
                                                    name = listItem.Name;
                                                }
                                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                                {
                                                    name = "";
                                                }
                                                //ControlType
                                                try
                                                {
                                                    controltype = listItem.ControlType.ToString();
                                                }
                                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                                {
                                                    controltype = "";
                                                }
                                                //Control Location
                                                try
                                                {
                                                    x = listItem.BoundingRectangle.X.ToString();
                                                    y = listItem.BoundingRectangle.Y.ToString();
                                                    width = listItem.BoundingRectangle.Width.ToString();
                                                    height = listItem.BoundingRectangle.Height.ToString();
                                                }
                                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                                {
                                                    x = "";
                                                    y = "";
                                                    width = "";
                                                    height = "";

                                                }

                                                result = $"{Environment.NewLine}{prefrencesWindow.Name},{automationId},{name}," +
                                                         $"{controltype}," +
                                                         $"{x},{y}," +
                                                         $"{width},{height}";
                                                resultChild.Add(result);
                                            }
                                        }
                                        combo.Click(true);
                                        Thread.Sleep(500);
                                    }
                                }
                                //AutomationId
                                try
                                {
                                    automationId = Controls.AutomationId;
                                }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                {
                                    automationId = "";
                                }
                                //Name
                                try
                                {
                                    name = Controls.Name;
                                    if (name.Contains("\n"))
                                    {
                                        name = name.Replace("\r\n", String.Empty);
                                        Thread.Sleep(1000);
                                    }
                                }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                {
                                    name = "";
                                }
                                //ControlType
                                try
                                {
                                    controltype = Controls.ControlType.ToString();
                                }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                {
                                    controltype = "";
                                }
                                //Control Location
                                try
                                {
                                    x = Controls.BoundingRectangle.X.ToString();
                                    y = Controls.BoundingRectangle.Y.ToString();
                                    width = Controls.BoundingRectangle.Width.ToString();
                                    height = Controls.BoundingRectangle.Height.ToString();
                                }
                                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                {
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";

                                }

                                result = $"{Environment.NewLine}{prefrencesWindow.Name},{automationId},{name}," +
                                         $"{controltype}," +
                                         $"{x},{y}," +
                                         $"{width},{height}";
                                if (name.Contains("\n"))
                                {
                                    name = name.Replace("\r\n", String.Empty);
                                    Thread.Sleep(1000);
                                }
                                resultChild.Add(result);
                                resultChild = removeDuplicates(resultChild);

                            }
                        }
                        else
                        {
                            resultChild.Add($"{Environment.NewLine}{prefrencesWindow.Name} Screen Controls Not Found");
                        }
                        resultChild = removeDuplicates(resultChild);
                        CancelButton.Click(true);
                        Thread.Sleep(1000);
                        management.Click();
                        #endregion

                        #region Haroon
                        //var PrefrencesControls = prefrencesWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

                        //window = app.GetMainWindow(automation);
                        //Thread.Sleep(100);
                        //var temp4 = window.FindFirstDescendant(z => z.ByAutomationId("OverLayForm"));
                        //if (temp4 != null)
                        //{
                        //    try
                        //    {
                        //        Thread.Sleep(100);
                        //        temp4 = temp4.FindFirstDescendant(z => z.ByAutomationId("ApplicationSettingsForm")); //"ApplicationSettingsForm"
                        //        var childsPreferences = temp4.FindAll(TreeScope.Descendants, TrueCondition.Default);
                        //        if (childsPreferences != null)
                        //        {
                        //            foreach (var subChildPreferences in childsPreferences)
                        //            {
                        //                GetPrefrencesScreenControls(automation, subChildPreferences.AutomationId);

                        //                automationId = "";
                        //                name = "";
                        //                controltype = "";
                        //                parent = "";
                        //                x = "";
                        //                y = "";
                        //                width = "";
                        //                height = "";
                        //                result = "";
                        //                //AutomationId
                        //                try
                        //                {
                        //                    automationId = subChildPreferences.AutomationId;
                        //                }
                        //                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        //                {
                        //                    automationId = "";
                        //                }
                        //                //Name
                        //                try
                        //                {
                        //                    name = subChildPreferences.Name;
                        //                }
                        //                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        //                {
                        //                    name = "";
                        //                }
                        //                //ControlType
                        //                try
                        //                {
                        //                    controltype = subChildPreferences.ControlType.ToString();
                        //                }
                        //                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        //                {
                        //                    controltype = "";
                        //                }
                        //                //Control Location
                        //                try
                        //                {
                        //                    x = subChildPreferences.BoundingRectangle.X.ToString();
                        //                    y = subChildPreferences.BoundingRectangle.Y.ToString();
                        //                    width = subChildPreferences.BoundingRectangle.Width.ToString();
                        //                    height = subChildPreferences.BoundingRectangle.Height.ToString();

                        //                }
                        //                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        //                {
                        //                    x = "";
                        //                    y = "";
                        //                    width = "";
                        //                    height = "";
                        //                }

                        //                result = $"{automationId}{name}, " +
                        //                         $"{controltype}," +
                        //                         $"{x},{y}," +
                        //                         $"{width},{height}";
                        //                resultChild.Add(result);
                        //            }
                        //            preferences = window.FindFirstDescendant(z => z.ByName("Cancel"));
                        //            preferences.Click();
                        //        }
                        //    }
                        //    catch (Exception)
                        //    {
                        //        preferences = window.FindFirstDescendant(z => z.ByName("Cancel"));
                        //        preferences.Click();
                        //        resultChild.Add("Prefernces Screen Control Not Found.,");
                        //    }
                        //} 
                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }
        //Online Screen
        public static void GetEMSPointTestingControls(UIA3Automation automation)
        {
            #region EMS Points Testing
            try
            {
                resultChild.Add($"{Environment.NewLine}EMS Point Testing");

                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Online Mode")
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var emsPntTest = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.EMSPointsTesting.ToDescriptionString()));
                    if (emsPntTest != null)
                    {
                        var Point = emsPntTest.GetClickablePoint();
                        Point.X = Point.X - 25;
                        Point.Y = Point.Y + 4;
                        FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                        window = app.GetMainWindow(automation);
                        emsPntTest = window.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                        if (emsPntTest != null)
                        {
                            var childsEMSPntTest = emsPntTest.FindAll(TreeScope.Descendants, TrueCondition.Default);
                            if (childsEMSPntTest != null)
                            {
                                foreach (var subChildEMSPntTest in childsEMSPntTest)
                                {
                                    automationId = "";
                                    name = "";
                                    controltype = "";
                                    parent = "";
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                    result = "";
                                    //AutomationId
                                    try
                                    {
                                        automationId = subChildEMSPntTest.AutomationId;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        automationId = "";
                                    }
                                    //Name
                                    try
                                    {
                                        name = subChildEMSPntTest.Name;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        name = "";
                                    }
                                    //ControlType
                                    try
                                    {
                                        controltype = subChildEMSPntTest.ControlType.ToString();
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        controltype = "";
                                    }
                                    //Control Location
                                    try
                                    {
                                        x = subChildEMSPntTest.BoundingRectangle.X.ToString();
                                        y = subChildEMSPntTest.BoundingRectangle.Y.ToString();
                                        width = subChildEMSPntTest.BoundingRectangle.Width.ToString();
                                        height = subChildEMSPntTest.BoundingRectangle.Height.ToString();

                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        x = "";
                                        y = "";
                                        width = "";
                                        height = "";
                                    }

                                    result = $"{Environment.NewLine}{automationId},{name}, " +
                                             $"{controltype}," +
                                             $"{x},{y}," +
                                             $"{width},{height}";
                                    resultChild.Add(result);
                                }
                                management.Click(true);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }
        //Online
        public static void GetSystemAvailabilityControls(UIA3Automation automation)
        {
            #region System Availability
            try
            {
                resultChild.Add($"{Environment.NewLine}System Availablity");
                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Online Mode")
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var sysAvailability = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SystemAvailability.ToDescriptionString()));
                    if (sysAvailability != null)
                    {
                        var Point = sysAvailability.GetClickablePoint();
                        Point.Y = Point.Y - 10;
                        FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                        window = app.GetMainWindow(automation);
                        sysAvailability = window.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                        if (sysAvailability != null)
                        {
                            var childsSysAvailability = sysAvailability.FindAll(TreeScope.Descendants, TrueCondition.Default);
                            if (childsSysAvailability != null)
                            {
                                foreach (var subChildSysAvailability in childsSysAvailability)
                                {
                                    automationId = "";
                                    name = "";
                                    controltype = "";
                                    parent = "";
                                    x = "";
                                    y = "";
                                    width = "";
                                    height = "";
                                    result = "";
                                    //AutomationId
                                    try
                                    {
                                        automationId = subChildSysAvailability.AutomationId;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        automationId = "";
                                    }
                                    //Name
                                    try
                                    {
                                        name = subChildSysAvailability.Name;
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        name = "";
                                    }
                                    //ControlType
                                    try
                                    {
                                        controltype = subChildSysAvailability.ControlType.ToString();
                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        controltype = "";
                                    }
                                    //Control Location
                                    try
                                    {
                                        x = subChildSysAvailability.BoundingRectangle.X.ToString();
                                        y = subChildSysAvailability.BoundingRectangle.Y.ToString();
                                        width = subChildSysAvailability.BoundingRectangle.Width.ToString();
                                        height = subChildSysAvailability.BoundingRectangle.Height.ToString();

                                    }
                                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                                    {
                                        x = "";
                                        y = "";
                                        width = "";
                                        height = "";
                                    }

                                    result = $"{Environment.NewLine}System Availablity,{automationId},{name}, " +
                                             $"{controltype}," +
                                             $"{x},{y}," +
                                             $"{width},{height}";
                                    resultChild.Add(result);
                                }
                                management.Click(true);
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }

        public static void GetQuietTimeConfigurationControls(UIA3Automation automation)
        {
            #region Quiet Time Configuration
            try
            {
                //resultChild.Add($"{Environment.NewLine}Quiet Time Configuration");
                var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                management.Click();
                window = app.GetMainWindow(automation);
                var quietTimeConfiguration = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.QuietTimeConfiguration.ToDescriptionString()).And(z.ByControlType(FlaUI.Core.Definitions.ControlType.Hyperlink)));
                var Point = quietTimeConfiguration.GetClickablePoint();
                Point.Y = Point.Y - 10;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);

                window = app.GetMainWindow(automation);
                var quietTimeWindow = window.FindFirstDescendant(cf => cf.ByName("Quiet Time Configuration & Scheduling")).AsWindow();
                quietTimeConfiguration = quietTimeWindow.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                var childsQuietTimeConfiguration = quietTimeConfiguration.FindAll(TreeScope.Descendants, TrueCondition.Default);
                if (childsQuietTimeConfiguration != null)
                {
                    foreach (var subChildQuietTimeConfiguration in childsQuietTimeConfiguration)
                    {
                        automationId = "";
                        name = "";
                        controltype = "";
                        parent = "";
                        x = "";
                        y = "";
                        width = "";
                        height = "";
                        result = "";
                        //AutomationId
                        try
                        {
                            automationId = subChildQuietTimeConfiguration.AutomationId;
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            automationId = "";
                        }
                        //Name
                        try
                        {
                            name = subChildQuietTimeConfiguration.Name;
                            if (name.Contains("Not connected to the PowerLine Gateway"))
                            {
                                name = name.Replace("Not connected to the PowerLine Gateway, trying to reconnect.", "Not connected to the PowerLine Gateway trying to reconnect.");
                                Thread.Sleep(1000);
                            }
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            name = "";
                        }
                        //ControlType
                        try
                        {
                            controltype = subChildQuietTimeConfiguration.ControlType.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            controltype = "";
                        }
                        //Control Location
                        try
                        {
                            x = subChildQuietTimeConfiguration.BoundingRectangle.X.ToString();
                            y = subChildQuietTimeConfiguration.BoundingRectangle.Y.ToString();
                            width = subChildQuietTimeConfiguration.BoundingRectangle.Width.ToString();
                            height = subChildQuietTimeConfiguration.BoundingRectangle.Height.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";

                        }
                        result = $"{Environment.NewLine}{quietTimeWindow.Name},{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";                        
                        resultChild.Add(result);
                        resultChild = removeDuplicates(resultChild);                       
                    }
                    management.Click();
                }
                else
                {
                    resultChild.Add($"{Environment.NewLine}{quietTimeWindow.Name} Screen Controls Not Found");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }

        public static void GetImportKeyFileControls(UIA3Automation automation)
        {
            #region Key Handling => Child->Import Key File
            
            try
            {
                window = app.GetMainWindow(automation);
                if (window.Name == "SmartInterface - Configuration Mode")
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var imprtKeyFile = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.ImportKeyFile.ToDescriptionString()));
                    var Point = imprtKeyFile.GetClickablePoint();
                    Point.Y = Point.Y - 10;
                    FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                    window = app.GetMainWindow(automation);
                    Thread.Sleep(100);

                    var temp7 = window.FindFirstDescendant(z => z.ByName("Open"));
                    var childsImprtKeyFile = temp7.FindAll(TreeScope.Descendants, TrueCondition.Default);
                    var cancelButton = temp7.FindFirstDescendant(z => z.ByName("Cancel")).AsButton();
                    if (childsImprtKeyFile != null)
                    {
                        foreach (var subChildImprtKeyFile in childsImprtKeyFile)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try
                            {
                                automationId = subChildImprtKeyFile.AutomationId;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                automationId = "";
                            }
                            //Name
                            try
                            {
                                name = subChildImprtKeyFile.Name;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                name = "";
                            }
                            //ControlType
                            try
                            {
                                controltype = subChildImprtKeyFile.ControlType.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                controltype = "";
                            }
                            //Control Location
                            try
                            {
                                x = subChildImprtKeyFile.BoundingRectangle.X.ToString();
                                y = subChildImprtKeyFile.BoundingRectangle.Y.ToString();
                                width = subChildImprtKeyFile.BoundingRectangle.Width.ToString();
                                height = subChildImprtKeyFile.BoundingRectangle.Height.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";

                            }

                            result = $"{Environment.NewLine}Import Key File,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                        }
                    }
                    else
                    {
                        resultChild.Add($"{Environment.NewLine}Import Key File Screen Controls Not Found");
                    }
                    Thread.Sleep(100);
                    cancelButton.Click();
                    Thread.Sleep(100);
                    management.Click(true);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }

        public static void GetAboutControls(UIA3Automation automation)
        {
            //Couldn't Access Control of the Popup Screen
            #region About
            try
            {
                //resultChild.Add($"{Environment.NewLine}About");
                var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                management.Click();
                window = app.GetMainWindow(automation);
                var about = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.About.ToDescriptionString()));
                var Point = about.GetClickablePoint();
                Point.X = Point.X - 66;
                Point.Y = Point.Y - 10;
                FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);

                window = app.GetMainWindow(automation);
                var temp8 = window.FindFirstDescendant(z => z.ByAutomationId("LoginPopup"));
                var childsAbout = temp8.FindAllDescendants();
                if (childsAbout != null)
                {
                    foreach (var subChildAbout in childsAbout)
                    {
                        automationId = "";
                        name = "";
                        controltype = "";
                        parent = "";
                        x = "";
                        y = "";
                        width = "";
                        height = "";
                        result = "";
                        //AutomationId
                        try
                        {
                            automationId = subChildAbout.AutomationId;
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            automationId = "";
                        }
                        //Name
                        try
                        {
                            name = subChildAbout.Name;
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            name = "";
                        }
                        //ControlType
                        try
                        {
                            controltype = subChildAbout.ControlType.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            controltype = "";
                        }
                        //Parent
                        //try
                        //{
                        //    parent = subChildSiteStatus.Parent.ToString();
                        //}
                        //catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        //{
                        //    parent = "";
                        //}
                        //Control Location
                        try
                        {
                            x = subChildAbout.BoundingRectangle.X.ToString();
                            y = subChildAbout.BoundingRectangle.Y.ToString();
                            width = subChildAbout.BoundingRectangle.Width.ToString();
                            height = subChildAbout.BoundingRectangle.Height.ToString();

                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                        }

                        result = $"{Environment.NewLine}About,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                        resultChild.Add(result);
                    }
                    management.Click(true);
                    Thread.Sleep(100);
                    management.Click(true);
                }
            }
            catch (Exception ex)
            {
                resultChild.Add("Couldn't find the control of 'About Screen'");
            }

            #endregion
        }

        public static void GetHelpControls(UIA3Automation automation)
        {
            #region Help
            try
            {
                //resultChild.Add($"{Environment.NewLine}Help");
                window = app.GetMainWindow(automation);
                if (window.Name.Contains("Configuration"))
                {
                    var management = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ManagementButtonUC.ToDescriptionString()));
                    management.Click();
                    window = app.GetMainWindow(automation);
                    var help = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.Help.ToDescriptionString()).And(z.ByControlType(FlaUI.Core.Definitions.ControlType.Hyperlink)));
                    var Point = help.GetClickablePoint();
                    FlaUI.Core.Input.Mouse.Click(Point, FlaUI.Core.Input.MouseButton.Left);
                    window = app.GetMainWindow(automation);
                    Thread.Sleep(100);
                    var temp9 = window.FindFirstDescendant(z => z.ByName("SmartInterface™ User Manual -Adobe Acrobat Reader DC(64 - bit)")).AsWindow();
                    //var temp9 = window.FindFirstDescendant(z => z.ByControlType(FlaUI.Core.Definitions.ControlType.TitleBar));
                    var childsHelp = temp9.FindAll(TreeScope.Descendants, TrueCondition.Default);

                    var closeButton = temp9.FindFirstDescendant(cf => cf.ByName("Close")).AsButton();
                    if (childsHelp != null)
                    {
                        foreach (var subChildHelp in childsHelp)
                        {
                            automationId = "";
                            name = "";
                            controltype = "";
                            parent = "";
                            x = "";
                            y = "";
                            width = "";
                            height = "";
                            result = "";
                            //AutomationId
                            try
                            {
                                automationId = subChildHelp.AutomationId;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                automationId = "";
                            }
                            //Name
                            try
                            {
                                name = subChildHelp.Name;
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                name = "";
                            }
                            //ControlType
                            try
                            {
                                controltype = subChildHelp.ControlType.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                controltype = "";
                            }
                            //Control Location
                            try
                            {
                                x = subChildHelp.BoundingRectangle.X.ToString();
                                y = subChildHelp.BoundingRectangle.Y.ToString();
                                width = subChildHelp.BoundingRectangle.Width.ToString();
                                height = subChildHelp.BoundingRectangle.Height.ToString();
                            }
                            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            {
                                x = "";
                                y = "";
                                width = "";
                                height = "";
                            }

                            result = $"{Environment.NewLine}Help,{automationId},{name}," +
                                     $"{controltype}," +
                                     $"{x},{y}," +
                                     $"{width},{height}";
                            resultChild.Add(result);
                        }
                    }
                    Thread.Sleep(100);
                    closeButton.Click(true);
                    //help = window.FindFirstDescendant(z => z.ByName("OK"));
                    //help.Click();
                    Thread.Sleep(100);
                    management.Click(true);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
            #endregion
        }

        public static void LoadChildren(/*string logsPath, */ AutomationElement mainchild)
        {
            var subChild = mainchild.FindAllDescendants(FlaUI.Core.Conditions.TrueCondition.Default);
            if (subChild.Length > 0)
            {
                foreach (var child in subChild)
                {
                    string automationId = "";
                    string name = "";
                    string controltype = "";
                    string parent = "";
                    string x = "";
                    string y = "";
                    string width = "";
                    string height = "";
                    string str = "";

                    try
                    {
                        automationId = child.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = child.Name;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = child.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Parent
                    try
                    {
                        parent = child.Parent.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        parent = "";
                    }
                    //Control Location
                    try
                    {
                        if ((automationId == "" || name == "" || controltype == "" || parent == "") && controltype != "Text" && controltype != "MenuItem")
                        {

                            x = child.BoundingRectangle.X.ToString();
                            y = child.BoundingRectangle.Y.ToString();
                            width = child.BoundingRectangle.Width.ToString();
                            height = child.BoundingRectangle.Height.ToString();

                            FlaUI.Core.Input.Mouse.Click(child.GetClickablePoint(), FlaUI.Core.Input.MouseButton.Left);
                        }
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                    }

                    str = $"AutomationId = {automationId}, Name = {name}, " +
                             $"ControlType = {controltype}, Parent = {parent}, " +
                             $"LocationX = {x}, LocationY = {y}," +
                             $"Width = {width}, Height = {height}{ Environment.NewLine}";

                    var innerChild = child.FindAllDescendants(FlaUI.Core.Conditions.TrueCondition.Default);

                    if (innerChild.Length > 0 && child.ControlType == FlaUI.Core.Definitions.ControlType.Window)
                    {
                        LoadChildren(child);
                    }
                    resultChild.Add(str);
                }
            }

            #region MyRegion
            //try
            //{
            //    #region Commented By MHA for debugging
            //    //if (children.ControlType == FlaUI.Core.Definitions.ControlType.Button ||
            //    //                children.ControlType == FlaUI.Core.Definitions.ControlType.Hyperlink)
            //    //{
            //    //    if (children.IsAvailable && children.IsEnabled /*&& children.AutomationId != "UpButton"*/)
            //    //    {

            //    //        children.Click(false);
            //    //        var subChild = children.FindAllNested(FlaUI.Core.Conditions.TrueCondition.Default);
            //    //        if (subChild.Length > 0)
            //    //        {
            //    //            foreach (var subChildren in subChild)
            //    //            {
            //    //                File.AppendAllText(logsPath,
            //    //                               $"Automation ID = {subChildren.AutomationId} {Environment.NewLine}" +
            //    //                               $"Name = {subChildren.Name} {Environment.NewLine}" +
            //    //                               $"Control Type = {subChildren.ControlType}{Environment.NewLine}");
            //    //                LoadChildren(logsPath, subChildren);

            //    //            }
            //    //        }
            //    //    }
            //    //} 
            //    #endregion

            //    #region Commented By MHA
            //    //else if (children.ControlType == FlaUI.Core.Definitions.ControlType.Hyperlink && children.IsAvailable)
            //    //{
            //    //    children.Click(false);
            //    //    var subChild = children.FindAllNested(FlaUI.Core.Conditions.TrueCondition.Default);
            //    //    if (subChild.Length > 0)
            //    //    {
            //    //        foreach (var subChildren in subChild)
            //    //        {
            //    //            File.AppendAllText(@"C:\Users\muhammad.haroon\Desktop\Async Example\Logs.txt",
            //    //                           $"Automation ID = {subChildren.AutomationId} {Environment.NewLine}" +
            //    //                           $"Name = {subChildren.Name} {Environment.NewLine}" +
            //    //                           $"Control Type = {subChildren.ControlType}{Environment.NewLine}");
            //    //        }

            //    //    }
            //    //}
            //    //else if (children.ControlType == FlaUI.Core.Definitions.ControlType.Pane && children.IsAvailable)
            //    //{
            //    //    children.Click(false);
            //    //    var subChild = children.FindAllNested(FlaUI.Core.Conditions.TrueCondition.Default);
            //    //    if (subChild.Length > 0)
            //    //    {
            //    //        foreach (var subChildren in subChild)
            //    //        {
            //    //            File.AppendAllText(@"C:\Users\muhammad.haroon\Desktop\Async Example\Logs.txt",
            //    //                           $"Automation ID = {subChildren.AutomationId} {Environment.NewLine}" +
            //    //                           $"Name = {subChildren.Name} {Environment.NewLine}" +
            //    //                           $"Control Type = {subChildren.ControlType}{Environment.NewLine}");
            //    //        }

            //    //    }
            //    //}
            //    //else
            //    //{
            //    //    children.Click(false);
            //    //    var subChild = children.FindAllNested(FlaUI.Core.Conditions.TrueCondition.Default);
            //    //} 
            //    #endregion
            //}
            //catch (FlaUI.Core.Exceptions.PropertyNotSupportedException)
            //{
            //    #region Commented By MHA for debugging
            //    //if (children.ControlType == FlaUI.Core.Definitions.ControlType.Button ||
            //    //    children.ControlType == FlaUI.Core.Definitions.ControlType.Hyperlink)
            //    //{
            //    //    if (children.IsAvailable && children.IsEnabled /*&& children.AutomationId != "UpButton"*/)
            //    //    {

            //    //        children.Click(false);
            //    //        var subChild = children.FindAllNested(FlaUI.Core.Conditions.TrueCondition.Default);
            //    //        if (subChild.Length > 0)
            //    //        {
            //    //            foreach (var subChildren in subChild)
            //    //            {
            //    //                try
            //    //                {
            //    //                    File.AppendAllText(logsPath,
            //    //                                  $"{Environment.NewLine}Name = {subChildren.Name} {Environment.NewLine}" +
            //    //                                  $"Control Type = {subChildren.ControlType}{Environment.NewLine}");

            //    //                    LoadChildren(logsPath, subChildren);
            //    //                }
            //    //                catch (FlaUI.Core.Exceptions.PropertyNotSupportedException)
            //    //                {
            //    //                    File.AppendAllText(logsPath, $"{Environment.NewLine}Name = {subChildren.Name} {Environment.NewLine}");

            //    //                    LoadChildren(logsPath, subChildren);
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //} 
            //    #endregion
            //} 
            #endregion
        }

        public static void SelectComboboxValueByAutomationId(UIA3Automation automation, string comboBoxId, string valueToSelect)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            var combo = window.FindFirstDescendant(cf => cf.ByAutomationId(comboBoxId)).AsComboBox();
            Thread.Sleep(500);
            if (combo.IsReadOnly)
            {
                combo.Select(valueToSelect);
            }
            //var selectedItem = combo.SelectedItem;
            //Assert.That(selectedItem, Is.Not.Null);
            //Assert.That(selectedItem.Text, Is.EqualTo(valueToSelect));
        }

        public static void CheckCheckboxValueByAutomationId(UIA3Automation automation, string checkboxId, string windowName)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            var checkBox = window.FindFirstDescendant(cf => cf.ByAutomationId(checkboxId)).AsCheckBox();
            Thread.Sleep(500);
            if ((bool)!checkBox.IsChecked && checkBox.IsEnabled == true)
            {
                checkBox.IsChecked = true;
            }
            else { return; }
            var CheckBoxControls = window.FindAll(TreeScope.Descendants, TrueCondition.Default);

            foreach (var checkBoxSave in CheckBoxControls)
            {

                try
                {
                    if (checkBoxSave.ControlType == ControlType.ComboBox)
                    {
                        GetAllComboboxValuesByAutomationId(automation, checkBoxSave.AutomationId, window.Name);
                    }
                    //AutomationID
                    try
                    {
                        automationId = checkBoxSave.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = checkBoxSave.Name;
                        try
                        {
                            if (checkBoxSave.Name is null)
                            {
                                name = "";
                            }
                            if (name.Contains("\n"))
                            {
                                name = name.Replace("\r\n", String.Empty);
                                Thread.Sleep(1000);
                            }
                            if (name.Contains("No Circuit Present"))
                            {
                                name = name.Replace("No Circuit Present, Add a Circuit to Continue", "No Circuit Present Add a Circuit to Continue");
                                Thread.Sleep(1000);
                            }
                            if (name.Contains("The Network Configuration process failed for some device(s)"))
                            {
                                name = name.Replace("The Network Configuration process failed for some device(s), please use the Re-apply Configuration button to apply the same configuration again.", "The Network Configuration process failed for some device(s) please use the Re-apply Configuration button to apply the same configuration again.");
                                Thread.Sleep(1000);
                            }
                        }
                        catch (Exception)
                        {
                            name = "";
                        }
                        
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = checkBoxSave.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = checkBoxSave.BoundingRectangle.X.ToString();
                        y = checkBoxSave.BoundingRectangle.Y.ToString();
                        width = checkBoxSave.BoundingRectangle.Width.ToString();
                        height = checkBoxSave.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{windowName},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    if (name.Contains("\n"))
                    {
                        name = name.Replace("\r\n", String.Empty);
                        Thread.Sleep(1000);
                    }
                    resultChild.Add(result);

                    resultChild = removeDuplicates(resultChild);

                }

                catch (Exception ex)
                {
                    x = "";
                    y = "";
                    width = "";
                    height = "";
                }
            }
            resultChild = removeDuplicates(resultChild);
            if ((bool)checkBox.IsChecked && checkBox.IsEnabled == true)
            {
                checkBox.IsChecked = false;
            }

            //var selectedItem = combo.SelectedItem;
            //Assert.That(selectedItem, Is.Not.Null);
            //Assert.That(selectedItem.Text, Is.EqualTo(valueToSelect));
        }

        public static void GetAllComboboxValuesByAutomationId(UIA3Automation automation, string comboBoxId, string windowName)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            var combo = window.FindFirstDescendant(cf => cf.ByAutomationId(comboBoxId)).AsComboBox();
            if (combo != null)
            {
                combo.Click(true);
                var comboBoxValues = combo.FindAll(TreeScope.Descendants, TrueCondition.Default);
                //ArayOfComboBoxValues = new string[comboBoxValues.Length];
                //for (int i = 0; i < ArayOfComboBoxValues.Length; i++)
                //{
                //    if (comboBoxValues[i].Name == "" || comboBoxValues[i].Name == "Close" || comboBoxValues[i].Name == "Open"){}
                //    else { ArayOfComboBoxValues[i] = comboBoxValues[i].Name; }
                //}
                foreach (var listItem in comboBoxValues)
                {
                    if (listItem.ControlType == ControlType.ListItem)
                    {

                        //AutomationId
                        try
                        {
                            automationId = listItem.AutomationId;
                            //if (listItem.AutomationId is null)
                            //{
                            //    automationId = "";
                            //}
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            automationId = "";
                        }
                        //Name
                        try
                        {
                            name = listItem.Name;
                            //if (listItem.Name is null)
                            //{
                            //    name = "";
                            //}
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            name = "";
                        }
                        //ControlType
                        try
                        {
                            controltype = listItem.ControlType.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            controltype = "";
                        }
                        //Control Location
                        try
                        {
                            x = listItem.BoundingRectangle.X.ToString();
                            y = listItem.BoundingRectangle.Y.ToString();
                            width = listItem.BoundingRectangle.Width.ToString();
                            height = listItem.BoundingRectangle.Height.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";

                        }

                        result = $"{Environment.NewLine}{windowName},{automationId},{name}," +
                                 $"{controltype}," +
                                 $"{x},{y}," +
                                 $"{width},{height}";
                        resultChild.Add(result);
                    }
                }
                combo.Click(true);
                Thread.Sleep(500);
            }

        }
        
        public static void GetCircuitScreenControls(UIA3Automation automation)
        {
            try
            {
                var window = app.GetMainWindow(automation);
                Thread.Sleep(500);

                var wins = app.GetAllTopLevelWindows(automation);
                var overlayForm = wins[0];
                var gridWindow = overlayForm.FindFirstDescendant(cf => cf.ByName("Circuit")).AsWindow();
                //var dataGrid = overlayForm.FindFirstDescendant(cf => cf.ByName("grdCircuit")).AsDataGridView();

                var DataGrid = gridWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

                int noOfRows = 0;
                DesktopAuto.NGetMainWindowbyName(window.Name);
                DesktopAuto.NGetMainWindowbyName(wins[0].Name);
                DesktopAuto.NFindControlByAutomationId(6, "grdCircuit");
                DesktopAuto.NGetAllDataGridValues(2, out noOfRows);
                if (noOfRows > 0)
                {
                    var data = DesktopAuto.DataGridData;
                    for (int i = 0; i < noOfRows; i++)
                    {
                        for (int j = 0; j < 2; j++)
                        {
                            result = $"{Environment.NewLine}{gridWindow.Name}," + $"{i}.{j}," + $"{data[i, j]}," +
                                                         $"DataGrid Cell Data," +
                                                         "0," + "0," +
                                                         "0," + "0";
                            resultChild.Add(result);
                        }
                    }
                }

                var CancelButton = gridWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnBack")).AsButton();
                if (DataGrid != null)
                {
                    foreach (var Controls in DataGrid)
                    {
                        //AutomationId
                        try
                        {
                            automationId = Controls.AutomationId;
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            automationId = "";
                        }
                        //Name
                        try
                        {
                            name = Controls.Name;
                            if (name.Contains("\n"))
                            {
                                name = name.Replace("\r\n", String.Empty);
                                Thread.Sleep(1000);
                            }
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            name = "";
                        }
                        //ControlType
                        try
                        {
                            controltype = Controls.ControlType.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            controltype = "";
                        }
                        //Control Location
                        try
                        {
                            x = Controls.BoundingRectangle.X.ToString();
                            y = Controls.BoundingRectangle.Y.ToString();
                            width = Controls.BoundingRectangle.Width.ToString();
                            height = Controls.BoundingRectangle.Height.ToString();
                        }
                        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                        {
                            x = "";
                            y = "";
                            width = "";
                            height = "";

                        }

                        result = $"{Environment.NewLine}{gridWindow.Name},{automationId},{name}," +
                                 $"{controltype}," +
                                 $"{x},{y}," +
                                 $"{width},{height}";
                        if (name.Contains("\n"))
                        {
                            name = name.Replace("\r\n", String.Empty);
                            Thread.Sleep(1000);
                        }
                        resultChild.Add(result);
                        resultChild = removeDuplicates(resultChild);

                    }
                }
                else
                {
                    resultChild.Add($"{Environment.NewLine}{gridWindow.Name} Screen Controls Not Found");
                }
                resultChild = removeDuplicates(resultChild);
                CancelButton.Click(true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }

        }
        
        public static void GetPLCScreenControls(UIA3Automation automation)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            var systemConfigWin = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SystemConfiguration.ToDescriptionString()));
            System.Drawing.Point Point = new System.Drawing.Point();
            Point.X = 328;
            Point.Y = 603;
            FlaUI.Core.Input.Mouse.DoubleClick(Point, FlaUI.Core.Input.MouseButton.Left);
            Thread.Sleep(1000);                      
                        
            var gridWindow = window.FindFirstDescendant(cf => cf.ByName("PowerLine Coordinator")).AsWindow();
            //var dataGrid = window.FindFirstDescendant(cf => cf.ByName("grdNetworkCoordinator")).AsDataGridView();

            var DataGrid = gridWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

            int noOfRows = 0;
            DesktopAuto.NGetMainWindowbyName(gridWindow.Name);
            DesktopAuto.NFindControlByAutomationId(6, "grdNetworkCoordinator");
            DesktopAuto.NGetAllDataGridValues(2, out noOfRows);
            if (noOfRows > 0)
            {
                var data = DesktopAuto.DataGridData;
                for (int i = 0; i < noOfRows; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        result = $"{Environment.NewLine}{gridWindow.Name}," + $"{i}.{j}," + $"{data[i, j]}," +
                                                     $"DataGrid Cell Data," +
                                                     "0," + "0," +
                                                     "0," + "0";
                        resultChild.Add(result);
                    }
                }
            }
            
            var CancelButton = gridWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnBack")).AsButton();
            if (DataGrid != null)
            {
                foreach (var Controls in DataGrid)
                {
                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        if (name.Contains("\n"))
                        {
                            name = name.Replace("\r\n", String.Empty);
                            Thread.Sleep(1000);
                        }
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{gridWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    if (name.Contains("\n"))
                    {
                        name = name.Replace("\r\n", String.Empty);
                        Thread.Sleep(1000);
                    }
                    resultChild.Add(result);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{gridWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);
            CancelButton.Click(true);

        }

        public static void GetGatewayScreenControls(UIA3Automation automation)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            var systemConfigWin = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SystemConfiguration.ToDescriptionString()));
            System.Drawing.Point Point = new System.Drawing.Point();
            Point.X = 330;
            Point.Y = 456;
            FlaUI.Core.Input.Mouse.DoubleClick(Point, FlaUI.Core.Input.MouseButton.Left);
            Thread.Sleep(1000);

            var gridWindow = window.FindFirstDescendant(cf => cf.ByName("PowerLine Gateway")).AsWindow();
            //var dataGrid = window.FindFirstDescendant(cf => cf.ByName("grdGateway")).AsDataGridView();

            var DataGrid = gridWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

            int noOfRows = 0;
            DesktopAuto.NGetMainWindowbyName(gridWindow.Name);
            DesktopAuto.NFindControlByAutomationId(6, "grdGateway");
            DesktopAuto.NGetAllDataGridValues(2, out noOfRows);
            if (noOfRows > 0)
            {
                var data = DesktopAuto.DataGridData;
                for (int i = 0; i < noOfRows; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        result = $"{Environment.NewLine}{gridWindow.Name}," + $"{i}.{j}," + $"{data[i, j]}," +
                                                     $"DataGrid Cell Data," +
                                                     "0," + "0," +
                                                     "0," + "0";
                        if (result.Contains("Elapsed Time, Line Current Stable"))
                        {
                            result = result.Replace("Elapsed Time, Line Current Stable", "Elapsed Time Line Current Stable");
                            Thread.Sleep(1000);
                        }
                        resultChild.Add(result);
                        
                    }
                }
            }           

            var CancelButton = gridWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnBack")).AsButton();
            if (DataGrid != null)
            {
                // for Configuration
                foreach (var Controls in DataGrid)
                {
                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        if (name.Contains("\n"))
                        {
                            name = name.Replace("\r\n", String.Empty);
                            Thread.Sleep(1000);
                        }
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{gridWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    if (name.Contains("\n"))
                    {
                        name = name.Replace("\r\n", String.Empty);
                        Thread.Sleep(1000);
                    }
                    resultChild.Add(result);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{gridWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);

            // for SFTP Setings
            var SFTP_Settings = gridWindow.FindFirstDescendant(cf => cf.ByName("SFTP Settings")).AsTabItem();
            SFTP_Settings.Click(true);
            var SFTP_Controls = gridWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

            if (SFTP_Controls != null)
            {
                foreach (var Controls in SFTP_Controls)
                {
                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        if (name.Contains("\n"))
                        {
                            name = name.Replace("\r\n", String.Empty);
                            Thread.Sleep(1000);
                        }
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{gridWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    if (name.Contains("\n"))
                    {
                        name = name.Replace("\r\n", String.Empty);
                        Thread.Sleep(1000);
                    }
                    resultChild.Add(result);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{gridWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);
            CancelButton.Click(true);

        }

        public static void GetDeviceScreenControls(UIA3Automation automation)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            var systemConfigWin = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SystemConfiguration.ToDescriptionString()));
            System.Drawing.Point Point = new System.Drawing.Point();
            Point.X = 571;
            Point.Y = 406;
            FlaUI.Core.Input.Mouse.DoubleClick(Point, FlaUI.Core.Input.MouseButton.Left);
            Thread.Sleep(1000);

            var gridWindow = window.FindFirstDescendant(cf => cf.ByName("Device")).AsWindow();
            //var dataGrid = window.FindFirstDescendant(cf => cf.ByName("grdUnit")).AsDataGridView();

            var DataGrid = gridWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

            int noOfRows = 0;
            DesktopAuto.NGetMainWindowbyName(gridWindow.Name);
            DesktopAuto.NFindControlByAutomationId(6, "grdUnit");
            DesktopAuto.NGetAllDataGridValues(2, out noOfRows);
            if (noOfRows > 0)
            {
                var data = DesktopAuto.DataGridData;
                for (int i = 0; i < noOfRows; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        result = $"{Environment.NewLine}{gridWindow.Name}," + $"{i}.{j}," + $"{data[i, j]}," +
                                                     $"DataGrid Cell Data," +
                                                     "0," + "0," +
                                                     "0," + "0";                        
                        resultChild.Add(result);

                    }
                }
            }

            var CancelButton = gridWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnBack")).AsButton();

            if (DataGrid != null)
            {
                foreach (var Controls in DataGrid)
                {
                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        if (name.Contains("\n"))
                        {
                            name = name.Replace("\r\n", String.Empty);
                            Thread.Sleep(1000);
                        }
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{gridWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    if (name.Contains("\n"))
                    {
                        name = name.Replace("\r\n", String.Empty);
                        Thread.Sleep(1000);
                    }
                    resultChild.Add(result);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{gridWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);
            CancelButton.Click(true);

        }

        public static void GetAddUserScreenControls(UIA3Automation automation)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            var systemConfigWin = window.FindFirstDescendant(z => z.ByName(EnumsSIMainScreen.AutmationIdOrName.SystemConfiguration.ToDescriptionString()));
            System.Drawing.Point Point = new System.Drawing.Point();
            Point.X = 1855;
            Point.Y = 208;
            FlaUI.Core.Input.Mouse.DoubleClick(Point, FlaUI.Core.Input.MouseButton.Left);
            Thread.Sleep(1000);

            var editUserWindow = window.FindFirstDescendant(cf => cf.ByName("Edit User")).AsWindow();

            var editUsercontrols = editUserWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);
                        
            var CancelButton = editUserWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnBack")).AsButton();

            if (editUsercontrols != null)
            {
                foreach (var Controls in editUsercontrols)
                {
                    if (Controls.ControlType == ControlType.ComboBox)
                    {
                        GetAllComboboxValuesByAutomationId(automation, Controls.AutomationId, editUserWindow.Name);
                    }

                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        if (name.Contains("\n"))
                        {
                            name = name.Replace("\r\n", String.Empty);
                            Thread.Sleep(1000);
                        }
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{editUserWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    if (name.Contains("\n"))
                    {
                        name = name.Replace("\r\n", String.Empty);
                        Thread.Sleep(1000);
                    }
                    resultChild.Add(result);
                    resultChild = removeDuplicates(resultChild);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{editUserWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);
            CancelButton.Click(true);

        }

        public static void GetAddLDPAUserScreenControls(UIA3Automation automation)
        {          

            var window = app.GetMainWindow(automation);

            var userManagementWind = window.FindFirstDescendant(z => z.ByAutomationId("UserManagement")).AsWindow();
            var addLDAPUserButton = userManagementWind.FindFirstDescendant(cf => cf.ByAutomationId("btnAddLDAPUser")).AsButton();
            addLDAPUserButton.Click(true);

            var activeDirectoryServerWindow = window.FindFirstDescendant(cf => cf.ByName("Active Directory Server")).AsWindow();

            var activeDirectoryControls = activeDirectoryServerWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

            if (activeDirectoryControls != null)
            {
                foreach (var Controls in activeDirectoryControls)
                {

                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        //if (name.Contains("\n"))
                        //{
                        //    name = name.Replace("\r\n", String.Empty);
                        //    Thread.Sleep(1000);
                        //}
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{activeDirectoryServerWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    if (name.Contains("\n"))
                    {
                        name = name.Replace("\r\n", String.Empty);
                        Thread.Sleep(1000);
                    }
                    resultChild.Add(result);
                    resultChild = removeDuplicates(resultChild);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{activeDirectoryServerWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);

            System.Drawing.Point Close = new System.Drawing.Point();
            Close.X = 1148;
            Close.Y = 365;
            FlaUI.Core.Input.Mouse.DoubleClick(Close, FlaUI.Core.Input.MouseButton.Left);
            Thread.Sleep(1000);
        }

        public static void GetResetPasswordScreenControls(UIA3Automation automation)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(500);
            
            System.Drawing.Point Point = new System.Drawing.Point();
            Point.X = 402;
            Point.Y = 246;
            FlaUI.Core.Input.Mouse.DoubleClick(Point, FlaUI.Core.Input.MouseButton.Left);
            Thread.Sleep(1000);


            window = app.GetMainWindow(automation);

            var userManagementWind = window.FindFirstDescendant(z => z.ByAutomationId("UserManagement")).AsWindow();
            var resetPasswordButton = userManagementWind.FindFirstDescendant(cf => cf.ByAutomationId("btnResetPassword")).AsButton();
            resetPasswordButton.Click(true);

            var resetPasswordWindow = window.FindFirstDescendant(cf => cf.ByName("Reset Password")).AsWindow();

            var resetPasswordControls = resetPasswordWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);

            var cancelButton = resetPasswordWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnCancel")).AsButton();

            if (resetPasswordControls != null)
            {
                foreach (var Controls in resetPasswordControls)
                {

                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        //if (name.Contains("\n"))
                        //{
                        //    name = name.Replace("\r\n", String.Empty);
                        //    Thread.Sleep(1000);
                        //}
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{resetPasswordWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    //if (name.Contains("\n"))
                    //{
                    //    name = name.Replace("\r\n", String.Empty);
                    //    Thread.Sleep(1000);
                    //}
                    resultChild.Add(result);
                    resultChild = removeDuplicates(resultChild);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{resetPasswordWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);
            Thread.Sleep(500);
            cancelButton.Click(true);

        }

        public static void GetAccountPolicyScreenControls(UIA3Automation automation)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(100);
            var userManagementWind = window.FindFirstDescendant(z => z.ByAutomationId("UserManagement")).AsWindow();
            var accountPolicyButton = userManagementWind.FindFirstDescendant(cf => cf.ByAutomationId("btnAccountPolicy")).AsButton();
            accountPolicyButton.Click(true);

            var accountPolicyWindow = window.FindFirstDescendant(cf => cf.ByName("Account Policy")).AsWindow();
            var accountPolicyControls = accountPolicyWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);
            var cancelButton = accountPolicyWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnCancel")).AsButton();

            if (accountPolicyControls != null)
            {
                foreach (var Controls in accountPolicyControls)
                {

                    //if (Controls.ControlType == ControlType.CheckBox)
                    //{

                    //    CheckCheckboxValueByAutomationId(automation, Controls.AutomationId, accountPolicyWindow.Name);
                    //}

                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        //if (name.Contains("\n"))
                        //{
                        //    name = name.Replace("\r\n", String.Empty);
                        //    Thread.Sleep(1000);
                        //}
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{accountPolicyWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    //if (name.Contains("\n"))
                    //{
                    //    name = name.Replace("\r\n", String.Empty);
                    //    Thread.Sleep(1000);
                    //}
                    resultChild.Add(result);
                    resultChild = removeDuplicates(resultChild);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{accountPolicyWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);
            Thread.Sleep(100);
            cancelButton.Click(true);

        }

        public static void GetUserRolesScreenControls(UIA3Automation automation)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(100);
            var userManagementWind = window.FindFirstDescendant(z => z.ByAutomationId("UserManagement")).AsWindow();
            var userRolesButton = userManagementWind.FindFirstDescendant(cf => cf.ByAutomationId("btnUserRoles")).AsButton();
            userRolesButton.Click(true);

            var userRolesWindow = window.FindFirstDescendant(cf => cf.ByName("User Roles")).AsWindow();
            var userRolesControls = userRolesWindow.FindAll(TreeScope.Descendants, TrueCondition.Default);
            var cancelButton = userRolesWindow.FindFirstDescendant(cf => cf.ByAutomationId("btnCancel")).AsButton();

            if (userRolesControls != null)
            {
                foreach (var Controls in userRolesControls)
                {

                    if (Controls.ControlType == ControlType.ComboBox)
                    {

                        GetAllComboboxValuesByAutomationId(automation, Controls.AutomationId, userRolesWindow.Name);
                    }

                    //AutomationId
                    try
                    {
                        automationId = Controls.AutomationId;
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        automationId = "";
                    }
                    //Name
                    try
                    {
                        name = Controls.Name;
                        //if (name.Contains("\n"))
                        //{
                        //    name = name.Replace("\r\n", String.Empty);
                        //    Thread.Sleep(1000);
                        //}
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        name = "";
                    }
                    //ControlType
                    try
                    {
                        controltype = Controls.ControlType.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        controltype = "";
                    }
                    //Control Location
                    try
                    {
                        x = Controls.BoundingRectangle.X.ToString();
                        y = Controls.BoundingRectangle.Y.ToString();
                        width = Controls.BoundingRectangle.Width.ToString();
                        height = Controls.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{userRolesWindow.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    //if (name.Contains("\n"))
                    //{
                    //    name = name.Replace("\r\n", String.Empty);
                    //    Thread.Sleep(1000);
                    //}
                    resultChild.Add(result);
                    resultChild = removeDuplicates(resultChild);
                }
            }
            else
            {
                resultChild.Add($"{Environment.NewLine}{userRolesWindow.Name} Screen Controls Not Found");
            }
            resultChild = removeDuplicates(resultChild);
            Thread.Sleep(100);
            cancelButton.Click(true);

        }

        public static void GetReportsTelemetryControls(UIA3Automation automation)
        {
            try
            {
                //resultChild.Add($"{Environment.NewLine}Reports");
                //var reportsButton = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ReportsButtonUC.ToDescriptionString()));
                //reportsButton.Click();

                var reportsWin = window.FindFirstDescendant(z => z.ByName("Reports & Analysis")).AsWindow();
                var reportsTelemetry = reportsWin.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                var childsReports = reportsTelemetry.FindAll(TreeScope.Descendants, TrueCondition.Default);
                foreach (var subChildReport in childsReports)
                {
                    try
                    {
                        if (subChildReport.ControlType == ControlType.ComboBox)
                        {
                            GetAllComboboxValuesByAutomationId(automation, subChildReport.AutomationId, reportsWin.Name);
                            #region Extra
                            //subChildReport.Click(true);
                            //for (int i = 1; i < ArayOfComboBoxValues.Length; i++)
                            //{
                            //    var comboBoxValues = subChildReport.FindFirstDescendant(z => z.ByAutomationId(ArayOfComboBoxValues[i]));
                            //    comboBoxValues.Click(true);
                            //    reportsTelemetry = reportsWin.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
                            //    childsReports = reportsTelemetry.FindAll(TreeScope.Descendants, TrueCondition.Default);
                            //    foreach (var item in childsReports)
                            //    {
                            //        try
                            //        {
                            //            if (item.ControlType == ControlType.ComboBox)
                            //            {
                            //                GetAllComboboxValuesByAutomationId(automation, item.AutomationId, reportsWin.Name);
                            //            }
                            //        }
                            //        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { }

                            //        automationId = "";
                            //        name = "";
                            //        controltype = "";
                            //        parent = "";
                            //        x = "";
                            //        y = "";
                            //        width = "";
                            //        height = "";
                            //        result = "";

                            //        ///AutomationId
                            //        try { automationId = item.AutomationId; }
                            //        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                            //        //Name
                            //        try { name = item.Name; }
                            //        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                            //        //ControlType
                            //        try { controltype = item.ControlType.ToString(); }
                            //        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                            //        //Control Location
                            //        try
                            //        {
                            //            x = item.BoundingRectangle.X.ToString();
                            //            y = item.BoundingRectangle.Y.ToString();
                            //            width = item.BoundingRectangle.Width.ToString();
                            //            height = item.BoundingRectangle.Height.ToString();
                            //        }
                            //        catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                            //        {
                            //            x = "";
                            //            y = "";
                            //            width = "";
                            //            height = "";

                            //        }

                            //        result = $"{Environment.NewLine}{reportsWin.Name},{automationId},{name}," +
                            //                 $"{controltype}," +
                            //                 $"{x},{y}," +
                            //                 $"{width},{height}";
                            //        resultChild.Add(result);

                            //    }


                            //} 
                            #endregion
                        }
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { }


                    #region Commented By SM
                    //try
                    //{
                    //    // Inner Pane of reports section
                    //    if (subChildReport.AutomationId == "pnlReport")
                    //    {
                    //        var innerChildsReports = reportsTelemetry.FindAll(TreeScope.Descendants, TrueCondition.Default);
                    //        foreach (var loop in innerChildsReports)
                    //        {
                    //            if (loop.ControlType == ControlType.ComboBox)
                    //            {
                    //                GetAllComboboxValuesByAutomationId(automation, loop.AutomationId, reportsWin.Name);
                    //            }

                    //            automationId = "";
                    //            name = "";
                    //            controltype = "";
                    //            parent = "";
                    //            x = "";
                    //            y = "";
                    //            width = "";
                    //            height = "";
                    //            result = "";
                    //            ///AutomationId
                    //            try
                    //            {

                    //                automationId = loop.AutomationId;
                    //                if (loop.AutomationId is null)
                    //                {
                    //                    automationId = "";
                    //                }
                    //            }
                    //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                    //            //Name
                    //            try
                    //            {
                    //                name = loop.Name;
                    //                if (loop.Name is null)
                    //                {
                    //                    name = "";
                    //                }
                    //            }
                    //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                    //            //ControlType
                    //            try { controltype = loop.ControlType.ToString(); }
                    //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                    //            //Control Location
                    //            try
                    //            {
                    //                x = loop.BoundingRectangle.X.ToString();
                    //                y = loop.BoundingRectangle.Y.ToString();
                    //                width = loop.BoundingRectangle.Width.ToString();
                    //                height = loop.BoundingRectangle.Height.ToString();
                    //            }
                    //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    //            {
                    //                x = "";
                    //                y = "";
                    //                width = "";
                    //                height = "";

                    //            }

                    //            result = $"{Environment.NewLine}Reports,{automationId},{name}," +
                    //                     $"{controltype}," +
                    //                     $"{x},{y}," +
                    //                     $"{width},{height}";
                    //            resultChild.Add(result);
                    //        }

                    //    }
                    //}
                    //catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { } 
                    #endregion

                    automationId = "";
                    name = "";
                    controltype = "";
                    parent = "";
                    x = "";
                    y = "";
                    width = "";
                    height = "";
                    result = "";

                    ///AutomationId
                    try { automationId = subChildReport.AutomationId; }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
                    //Name
                    try { name = subChildReport.Name; }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
                    //ControlType
                    try { controltype = subChildReport.ControlType.ToString(); }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
                    //Control Location
                    try
                    {
                        x = subChildReport.BoundingRectangle.X.ToString();
                        y = subChildReport.BoundingRectangle.Y.ToString();
                        width = subChildReport.BoundingRectangle.Width.ToString();
                        height = subChildReport.BoundingRectangle.Height.ToString();
                    }
                    catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
                    {
                        x = "";
                        y = "";
                        width = "";
                        height = "";

                    }

                    result = $"{Environment.NewLine}{reportsWin.Name},{automationId},{name}," +
                             $"{controltype}," +
                             $"{x},{y}," +
                             $"{width},{height}";
                    resultChild.Add(result);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
            }
        }

        //public static void GetReportsAlarmControls(UIA3Automation automation)
        //{
        //    try
        //    {
        //        //resultChild.Add($"{Environment.NewLine}Reports");
        //        //var reportsButton = window.FindFirstDescendant(z => z.ByAutomationId(EnumsSIMainScreen.AutmationIdOrName.ReportsButtonUC.ToDescriptionString()));
        //        //reportsButton.Click();
        //        window = app.GetMainWindow(automation);
        //        var reportsWin = window.FindFirstDescendant(z => z.ByName("Reports & Analysis")).AsWindow();
        //        var reportsTelemetry = reportsWin.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
        //        var comboClick = reportsTelemetry.FindFirstDescendant(z => z.ByAutomationId("cmbReport"));
        //        comboClick.Click(true);
        //        var alarmClick = reportsTelemetry.FindFirstDescendant(z => z.ByName("Alarms"));
        //        alarmClick.Click(true);
        //        reportsTelemetry = reportsWin.FindFirstDescendant(z => z.ByAutomationId("tableLayoutPanel1"));
        //        var childsReports = reportsTelemetry.FindAll(TreeScope.Descendants, TrueCondition.Default);
        //        foreach (var subChildReport in childsReports)
        //        {
        //            try
        //            {
        //                if (subChildReport.ControlType == ControlType.ComboBox)
        //                {
        //                    GetAllComboboxValuesByAutomationId(automation, subChildReport.AutomationId, reportsWin.Name);
                           
        //                }
        //            }
        //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { }


                    

        //            automationId = "";
        //            name = "";
        //            controltype = "";
        //            parent = "";
        //            x = "";
        //            y = "";
        //            width = "";
        //            height = "";
        //            result = "";

        //            ///AutomationId
        //            try { automationId = subChildReport.AutomationId; }
        //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { automationId = ""; }
        //            //Name
        //            try { name = subChildReport.Name; }
        //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { name = ""; }
        //            //ControlType
        //            try { controltype = subChildReport.ControlType.ToString(); }
        //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex) { controltype = ""; }
        //            //Control Location
        //            try
        //            {
        //                x = subChildReport.BoundingRectangle.X.ToString();
        //                y = subChildReport.BoundingRectangle.Y.ToString();
        //                width = subChildReport.BoundingRectangle.Width.ToString();
        //                height = subChildReport.BoundingRectangle.Height.ToString();
        //            }
        //            catch (FlaUI.Core.Exceptions.PropertyNotSupportedException ex)
        //            {
        //                x = "";
        //                y = "";
        //                width = "";
        //                height = "";

        //            }

        //            result = $"{Environment.NewLine}Reports,{automationId},{name}," +
        //                     $"{controltype}," +
        //                     $"{x},{y}," +
        //                     $"{width},{height}";
        //            resultChild.Add(result);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex.Message);
        //        Loggers.Log(Api.DesktopAutomationControlsFinder, ex);
        //    }
        //}
        public static void SetEditBoxValueByAutomationId(UIA3Automation automation, string editBoxId, string valueToEnter)
        {
            var window = app.GetMainWindow(automation);
            Thread.Sleep(300);
            var editBox = window.FindFirstDescendant(cf => cf.ByAutomationId(editBoxId)).AsTextBox();
            editBox.Enter(valueToEnter);
            //var selectedItem = combo.SelectedItem;
            //Assert.That(selectedItem, Is.Not.Null);
            //Assert.That(selectedItem.Text, Is.EqualTo("Configuration Mode"));
        }
        public static void InvokeButtonById(UIA3Automation automation, string editBoxId)
        {
            var window = app.GetMainWindow(automation);
            var btn = window.FindFirstDescendant(cf => cf.ByAutomationId(editBoxId)).AsButton();
            btn.Click();
            //var selectedItem = combo.SelectedItem;
            //Assert.That(selectedItem, Is.Not.Null);
            //Assert.That(selectedItem.Text, Is.EqualTo("Configuration Mode"));
        }
        public static void KillApplication()
        {
            Process[] processes = Process.GetProcesses().Where(x => x.MainWindowTitle.Length > 0).ToArray();
            foreach (Process process in processes)
            {
                if (process.ProcessName == "SmartInterface")
                {
                    process.Kill();
                    process.WaitForExit();
                    process.Dispose();
                }
            }
        }
        static List<string> removeDuplicates(List<string> inputList)
        {
            Dictionary<string, int> uniqueStore = new Dictionary<string, int>();
            List<string> finalList = new List<string>();
            foreach (string currValue in inputList)
            {
                if (!uniqueStore.ContainsKey(currValue))
                {
                    uniqueStore.Add(currValue, 0);
                    finalList.Add(currValue);
                }
            }
            return finalList;
        }

        //[Test]
        //public static void MenuInPopupTest()
        //{
        //    var app = FlaUI.Core.Application.Attach("SmartInterface");
        //    var automation = new UIA3Automation();
        //    var window = app.GetMainWindow(automation);

        //    //var btn = window.FindFirstDescendant(cf => cf.ByAutomationId("DashboardButtonUC"));
        //    var btn = window.FindFirstDescendant(x => x.ByControlType(FlaUI.Core.Definitions.ControlType.Hyperlink));
        //    btn.Click();
        //    Wait.UntilInputIsProcessed();
        //    //var popup = window.Popup;
        //    //Assert.That(popup, Is.Not.Null);
        //    var popupChildren = btn.FindAllDescendants();
        //    Assert.That(popupChildren, Has.Length.EqualTo(1));
        //    var menu = popupChildren[0].AsMenu();
        //    Assert.That(menu.Items, Has.Length.EqualTo(1));
        //    var menuItem = menu.Items[0];
        //    Assert.That(menuItem.Text, Is.EqualTo("Some MenuItem"));
        //}

        public static AutomationElement AutomationElement { get; }

        public static List<string> resultChild  = new List<string>();
    }
}

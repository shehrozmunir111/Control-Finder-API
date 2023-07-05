using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationControlsFinder
{
    public static class EnumsSIMainScreen
    {
        public static string ToDescriptionString(this AutmationIdOrName val)
        {
            DescriptionAttribute[] attributes = (DescriptionAttribute[])val
               .GetType()
               .GetField(val.ToString())
               .GetCustomAttributes(typeof(DescriptionAttribute), false);
            return attributes.Length > 0 ? attributes[0].Description : string.Empty;
        }
        public enum AutmationIdOrName
        {
            //AutomationId
            [Description("DashboardButtonUC")]
            DashboardButtonUC,
            //Child - Name/Hyperlink
            [Description("Site Status")]
            SiteStatus,

            //AutomationId
            [Description("ConfigButtonUC")]
            ConfigButtonUC,
            //Childs - Name/Hyperlink
            [Description("System Configuration")]
            SystemConfiguration,
            [Description("Alarms Configuration")]
            AlarmsConfiguration,
            [Description("Phase Naming Scheme")]
            PhaseNamingScheme,
            //PhaseNamingScreen AutomationId
            [Description("PhaseNamingSchemeForm")]
            PhaseNamingSchemeForm,

            //AutomationId
            [Description("StatusCommandsButtonUC")]
            StatusCommandsButtonUC,
            //Childs - Name/Hyperlink
            [Description("Circuit Status")]
            CircuitStatus,
            [Description("Alarm Status")]
            AlarmStatus,
            [Description("Structure Summary")]
            StructureSummary,

            //AutomationId
            [Description("ReportsButtonUC")]
            ReportsButtonUC,

            //AutomationId
            [Description("ManagementButtonUC")]
            ManagementButtonUC,
            //Childs - Name/Hyperlink
            [Description("Switch to Online")]
            SwitchtoOnline,
            [Description("Switch to Configuration")]
            SwitchtoConfiguration,
            [Description("User Management")]
            UserManagement,
            [Description("Firmware Update")]
            FirmwareUpdate,
            [Description("Export Configuration")]
            ExportConfiguration,
            [Description("Import Configuration")]
            ImportConfiguration,
            [Description("Automatic Injection Test View")]
            AutomaticInjectionTestView,
            [Description("Tunable Settings")]
            TunableSettings,
            [Description("Export Data")]
            ExportData,
            [Description("Preferences")]
            Preferences,
            [Description("EMS Points Testing")]
            EMSPointsTesting,
            [Description("System Availability")]
            SystemAvailability,
            [Description("Quiet Time Configuration")]
            QuietTimeConfiguration,

            //Child - Name/Hyperlink
            [Description("Key Handling")]
            KeyHandling,
            //SubChild - Name/Hyperlink
            [Description("Import Key File")]
            ImportKeyFile,

            [Description("About")]
            About,
            [Description("Help")]
            Help

        }
    }
}

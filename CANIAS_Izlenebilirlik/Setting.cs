using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace CANIAS_Izlenebilirlik
{
    [CompilerGenerated]
    [GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "14.0.0.0")]
    internal sealed class Setting : ApplicationSettingsBase
    {
        private static Setting defaultInstance = (Setting)SettingsBase.Synchronized((SettingsBase)new Setting());
        public static Setting Default
        {
            get
            {
                return Setting.defaultInstance;
            }
        }
        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue(@"\\192.168.0.220\eps\EPS\Ortak\ALPPLAS ÇİZİMLER\")]
        public string publicFolder
        {
            get
            {
                return (string)this[nameof(publicFolder)];
            }
            set
            {
                this[nameof(publicFolder)] = (object)value;
            }
        }

        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("COM11")]
        public string portNameSet
        {
            get
            {
                return (string)this[nameof(portNameSet)];
            }
            set
            {
                this[nameof(portNameSet)] = (object)value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("false")]
        public bool portNameSituation
        {
            get
            {
                return (bool)this[nameof(portNameSituation)];
            }
            set
            {
                this[nameof(portNameSituation)] = (object)value;
            }
        }


        [UserScopedSetting]
        [DebuggerNonUserCode]
        [DefaultSettingValue("false")]
        public bool publicFolderUsingSit
        {
            get
            {
                return (bool)this[nameof(publicFolderUsingSit)];
            }
            set
            {
                this[nameof(publicFolderUsingSit)] = (object)value;
            }
        }
    }
}

﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace SaveAsPDF.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "17.10.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("J:\\")]
        public string rootDrive {
            get {
                return ((string)(this["rootDrive"]));
            }
            set {
                this["rootDrive"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(".SaveAsPDF\\")]
        public string xmlSaveAsPdfFolder {
            get {
                return ((string)(this["xmlSaveAsPdfFolder"]));
            }
            set {
                this["xmlSaveAsPdfFolder"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(".SaveAsPDF_Project.xml")]
        public string xmlProjectFile {
            get {
                return ((string)(this["xmlProjectFile"]));
            }
            set {
                this["xmlProjectFile"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(".SaveAsPDF_Emploeeys.xml")]
        public string xmlEmploeeysFile {
            get {
                return ((string)(this["xmlEmploeeysFile"]));
            }
            set {
                this["xmlEmploeeysFile"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("DefaultTree.fld")]
        public string defaultTreeFile {
            get {
                return ((string)(this["defaultTreeFile"]));
            }
            set {
                this["defaultTreeFile"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_מספר_פרויקט_")]
        public string defaultFolder {
            get {
                return ((string)(this["defaultFolder"]));
            }
            set {
                this["defaultFolder"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("8192")]
        public int minAttachmentSize {
            get {
                return ((int)(this["minAttachmentSize"]));
            }
            set {
                this["minAttachmentSize"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_תאריך_")]
        public string dateTag {
            get {
                return ((string)(this["dateTag"]));
            }
            set {
                this["dateTag"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_מספר_פרויקט_")]
        public string projectRootTag {
            get {
                return ((string)(this["projectRootTag"]));
            }
            set {
                this["projectRootTag"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("1")]
        public int defaultFolderID {
            get {
                return ((int)(this["defaultFolderID"]));
            }
            set {
                this["defaultFolderID"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("False")]
        public bool OpenPDF {
            get {
                return ((bool)(this["OpenPDF"]));
            }
            set {
                this["OpenPDF"] = value;
            }
        }
    }
}

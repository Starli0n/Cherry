using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Diagnostics;

namespace Cherry
{
    class ResourceInfo
    {
        private string m_sLocation;
        private Dictionary<string, string> m_Resource = new Dictionary<string, string>();

        public ResourceInfo()
        {

        }

        public void Check()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(m_sLocation);
            m_Resource.Clear();
            m_Resource.Add("Comments", fvi.Comments);
            m_Resource.Add("CompanyName", fvi.CompanyName);
            m_Resource.Add("FileBuildPart", fvi.FileBuildPart.ToString());
            m_Resource.Add("FileDescription", fvi.FileDescription);
            m_Resource.Add("FileMajorPart", fvi.FileMajorPart.ToString());
            m_Resource.Add("FileMinorPart", fvi.FileMinorPart.ToString());
            m_Resource.Add("FileName", fvi.FileName);
            m_Resource.Add("FilePrivatePart", fvi.FilePrivatePart.ToString());
            m_Resource.Add("FileVersion", fvi.FileVersion);
            m_Resource.Add("InternalName", fvi.InternalName);
            m_Resource.Add("IsDebug", fvi.IsDebug.ToString());
            m_Resource.Add("IsPatched", fvi.IsPatched.ToString());
            m_Resource.Add("IsPreRelease", fvi.IsPreRelease.ToString());
            m_Resource.Add("IsPrivateBuild", fvi.IsPrivateBuild.ToString());
            m_Resource.Add("IsSpecialBuild", fvi.IsSpecialBuild.ToString());
            m_Resource.Add("Language", fvi.Language);
            m_Resource.Add("LegalCopyright", fvi.LegalCopyright);
            m_Resource.Add("LegalTrademarks", fvi.LegalTrademarks);
            m_Resource.Add("OriginalFilename", fvi.OriginalFilename);
            m_Resource.Add("PrivateBuild", fvi.PrivateBuild);
            m_Resource.Add("ProductBuildPart", fvi.ProductBuildPart.ToString());
            m_Resource.Add("ProductMajorPart", fvi.ProductMajorPart.ToString());
            m_Resource.Add("ProductMinorPart", fvi.ProductMinorPart.ToString());
            m_Resource.Add("ProductName", fvi.ProductName);
            m_Resource.Add("ProductPrivatePart", fvi.ProductPrivatePart.ToString());
            m_Resource.Add("ProductVersion", fvi.ProductVersion);
            m_Resource.Add("SpecialBuild", fvi.SpecialBuild);
        }

        public string Location
        {
            get
            {
                return m_sLocation;
            }
            set
            {
                m_sLocation = value;
                Check();
            }
        }

        public Dictionary<string, string> Resource
        {
            get
            {
                return m_Resource;
            }
        }
    }
}

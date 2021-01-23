/********************************************************************
      Author: Patrick Li  Date: 2021.1.23

  Each component info got by SolidWorks Document Management API (DM API) 

 * Copyright© the copyright belongs to Patrick Li
 *******************************************************************/
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Threading;

namespace Traverse
{
    /// <summary>
    /// Each component info in a SolidWorks document
    /// Got by SolidWorks Document Management API (DM API) 
    /// </summary>
    public class SldDMComponentItem
    {
        /// <summary>
        /// Id, starting from 1
        /// </summary>
        public int Id
        {
            get; set;
        }

        /// <summary>
        /// Level, the top level is 0, and then 1, 2, 3, N
        /// </summary>
        public int Level
        {
            get; set;
        }

        /// <summary>
        /// The parent Id
        /// </summary>
        public int ParentId
        {
            get; set;
        }

        /// <summary>
        /// Children count, which doesn't include myself
        /// </summary>
        public int ChildrenCount
        {
            get; set;
        }

        /// <summary>
        /// The full path of current component
        /// </summary>
        public string FullPath
        {
            get; set;
        }

        /// <summary>
        /// The directory
        /// </summary>
        public string DirectoryName
        {
            get
            {
                try
                {
                    return Path.GetDirectoryName(this.FullPath);
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// THe file name without extention
        /// </summary>
        public string FileNameWithoutExt
        {
            get
            {
                try
                {
                    return Path.GetFileNameWithoutExtension(this.FullPath);
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// The file info with extention
        /// </summary>
        public string FileName
        {
            get
            {
                try
                {
                    return Path.GetFileName(this.FullPath);
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>
        /// Tag
        /// </summary>
        public object Tag
        {
            get; set;
        }

        /// <summary>
        /// Name of the component
        /// </summary>
        public string FeatureName
        {
            get; set;
        }

        /// <summary>
        /// Id of the component, same components might have the same Id
        /// </summary>
        public int FeatureId
        {
            get; set;
        }

        /// <summary>
        /// Is it suppressed
        /// </summary>
        public bool IsSuppressed
        {
            get; set;
        }

        /// <summary>
        /// Configuration count
        /// </summary>
        public int ConfigurationCount
        {
            get; set;
        }

        /// <summary>
        /// Engineering file name (*.SLDDRW)
        /// </summary>
        public string FileName2D
        {
            get; set;
        }

        /// <summary>
        /// Is it visible?
        /// </summary>
        public bool Visible
        {
            get; set;
        }

        /// <summary>
        /// All custom properties
        /// </summary>
        public NameValueCollection AllCustomProperties
        {
            get; set;
        }

        /// <summary>
        /// The code of the material
        /// </summary>
        public string FCode
        {
            get
            {
                if (this.AllCustomProperties != null)
                {
                    string fCodePropertyName = this.AllCustomProperties.AllKeys.FirstOrDefault(x => string.Compare(x, GetDocumentPropertiesViaDM.PropertyNamesOftenUsed[0], true) == 0);
                    if (!string.IsNullOrEmpty(fCodePropertyName))
                        return this.AllCustomProperties[fCodePropertyName];
                }
                return string.Empty;
            }
        }

        /// <summary>
        /// Private construction
        /// </summary>
        private SldDMComponentItem()
        {
            this.Id = -1;
            this.Visible = true;
            //this.IsSuppressed = false;
        }

        /// <summary>
        /// Clone an item
        /// </summary>
        /// <param name="item">The item to be cloned from</param>
        /// <param name="newLevel">new level for the cloned item</param>
        /// <param name="newParentId">new parent if of the cloned item</param>
        /// <returns>The cloned item</returns>
        public static SldDMComponentItem Clone(SldDMComponentItem item, int newLevel, int newParentId)
        {
            if (item == null)
                return null;
            SldDMComponentItem cloned = GenerateTheNext();
            cloned.Level = newLevel;
            cloned.ParentId = newParentId;
            cloned.ChildrenCount = item.ChildrenCount;
            cloned.FullPath = item.FullPath;
            cloned.Tag = item.Tag;
            cloned.FeatureName = item.FeatureName;
            cloned.FeatureId = item.FeatureId;
            cloned.IsSuppressed = item.IsSuppressed;
            cloned.ConfigurationCount = item.ConfigurationCount;
            cloned.FileName2D = item.FileName2D;
            cloned.Visible = item.Visible;
            cloned.AllCustomProperties = item.AllCustomProperties;
            return cloned;
        }

        /// <summary>
        /// Current Id
        /// </summary>
        private static int currentId = 0;

        /// <summary>
        /// Generate the top item
        /// </summary>
        /// <returns>The root item with ID = 1</returns>
        public static SldDMComponentItem GenerateTheRoot()
        {
            currentId = 1;
            return new SldDMComponentItem()
            {
                Id = currentId,
                Level = 0,
                ParentId = 0
            };
        }

        /// <summary>
        /// Generate an item which is not the root (top) one
        /// </summary>
        /// <returns>Current item</returns>
        public static SldDMComponentItem GenerateTheNext()
        {
            int nextId = Interlocked.Increment(ref currentId);
            return new SldDMComponentItem()
            {
                Id = nextId
            };
        }
    }
}

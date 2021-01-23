/********************************************************************
      Author: Patrick Li   Date: 2021.1.23

    When traversing all components in a SolidWorks document, we won't
    get same named component by consuming COM operations, so create 
    this class for the component at first time. Subsequent same named 
    components will clone the previous item by taking advantage of
    this class

 * Copyright© the copyright belongs to Patrick Li
 *******************************************************************/
using System.Threading.Tasks;

namespace Traverse
{
    /// <summary>
    /// A SolidWorks component the the task getting its information
    /// </summary>
    internal abstract class ComponentAndTask
    {
        public string FullPath
        {
            get; set;
        }

        public int Id
        {
            get; set;
        }

        public int ParentId
        {
            get; set;
        }

        public Task TheTask
        {
            get; set;
        }
    }
    internal class ComponentAndTask_Assembly : ComponentAndTask
    {
    }

    internal class ComponentAndTask_Part : ComponentAndTask
    {
    }
}

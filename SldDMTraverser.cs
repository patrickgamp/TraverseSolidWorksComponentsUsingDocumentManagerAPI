/********************************************************************
      Author: Patrick Li   Date: 2021.1.23

    Traverse all components in a SolidWorks document using multiple
    tasks.
    Get each component's information by using Document Manager API

 * Copyright© the copyright belongs to Patrick Li
 *******************************************************************/
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using SolidWorks.Interop.swdocumentmgr;

namespace Traverse
{
    public class SldDMTraverser
    {
        public string TopAssemblyPath
        {
            get;
            protected set;
        }

        public List<SldDMComponentItem> Results
        {
            get;
            protected set;
        }

        public delegate bool GetComponentExtraInfo(ref SldDMComponentItem currentComponentItem);

        public GetComponentExtraInfo CallbackGettingExtraInfo
        {
            get; set;
        }

        protected Task WalkerTask
        {
            get;
            set;
        }

        protected CancellationTokenSource CancellationSwitch
        {
            get;
            set;
        }

        protected ConcurrentQueue<SldDMComponentItem> ResultsAtPopulating
        {
            get;
            set;
        }

        internal ConcurrentQueue<ComponentAndTask> AllWalkingTasks
        {
            get; set;
        }

        protected ConcurrentQueue<object> AllComObjects
        {
            get; set;
        }

        protected SwDMSearchOption GlobalSearchOption
        {
            get;
            set;
        }


        public bool IsRunning
        {
            get
            {
                return (this.CancellationSwitch != null && this.WalkerTask != null && !this.CancellationSwitch.IsCancellationRequested && this.WalkerTask.Status != TaskStatus.RanToCompletion);
            }
        }

        public void Abort()
        {
            if (this.CancellationSwitch != null)
            {
                this.CancellationSwitch.Cancel();
                this.CancellationSwitch = null;
            }
        }

        public void StopPossibleTask()
        {
            if (this.CancellationSwitch != null)
            {
                try
                {
                    this.CancellationSwitch.Cancel();
                    Thread.Sleep(50);
                    this.CancellationSwitch.Dispose();
                }
                catch
                { }
                finally
                {
                    this.CancellationSwitch = null;
                }
            }
            if (this.WalkerTask != null)
            {
                try
                {
                    this.WalkerTask.Dispose();
                }
                catch
                { }
                finally
                {
                    this.WalkerTask = null;
                }
            }
            this.Results = null;
            this.ResultsAtPopulating = null;
        }

        public SldDMTraverser()
        {
        }

        public Task Start(string topAssemblyPath)
        {
            if (string.IsNullOrEmpty(topAssemblyPath))
                return null;
            Debug.WriteLine(string.Format("SldDMTraverser.Start: traverse using Document Manager API to {0}", topAssemblyPath));
            this.StopPossibleTask();

            this.AllWalkingTasks = new ConcurrentQueue<ComponentAndTask>();
            this.ResultsAtPopulating = new ConcurrentQueue<SldDMComponentItem>();
            this.AllComObjects = new ConcurrentQueue<object>();

            this.TopAssemblyPath = Path.GetDirectoryName(topAssemblyPath);

            this.CancellationSwitch = new CancellationTokenSource();
            this.WalkerTask = Task.Factory.StartNew(() =>
            {
                try
                {
                    if (this.TraverseCore(topAssemblyPath, null))
                    {
                        Task[] allPendingTasks = this.AllWalkingTasks.Select(x => x.TheTask).ToArray();
                        Debug.WriteLine(string.Format("SldDMTraverser.Start:finished，wait for the {0} tasks signaled", allPendingTasks.Length));
                        Task.WaitAll(allPendingTasks);
                        Debug.WriteLine(string.Format("SldDMTraverser.Start:finished，there are  {0} items", this.ResultsAtPopulating.Count));
                        this.Results = this.ResultsAtPopulating.ToList();
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(string.Format("SldDMTraverser.Start:exception:{0}", ex.Message));
                }
                finally
                {
                    while (this.AllComObjects.Count > 0)
                    {
                        object comObject = null;
                        if (this.AllComObjects.TryDequeue(out comObject))
                        {
                            try
                            {
                                Marshal.ReleaseComObject(comObject);
                            }
                            catch
                            { }
                        }
                    }
                }
            });
            return this.WalkerTask;
        }

        protected bool TraverseCore(string assemblyPath, string configurationName)
        {
            if (string.IsNullOrEmpty(assemblyPath))
                return false;
            if (!File.Exists(assemblyPath))
                return false;
            string originalExt;
            SwDmDocumentType docType = SldFileExtentionChecker.CheckDM(assemblyPath, out originalExt);
            if (docType != SwDmDocumentType.swDmDocumentAssembly && docType != SwDmDocumentType.swDmDocumentPart)
            {
                return false;
            }
            SwDMClassFactory swDMClassFactory = new SwDMClassFactory();
            //this.AllComObjects.Enqueue(swDMClassFactory);
            SwDMApplication swDMApp = swDMClassFactory.GetApplication(GetDocumentPropertiesViaDM.LinktronLicenseKey);
            this.GlobalSearchOption = swDMApp.GetSearchOptionObject();
            //this.AllComObjects.Enqueue(swDMApp);
            SwDmDocumentOpenError returnValue = 0;
            SwDMDocument17 swDoc = (SwDMDocument17)swDMApp.GetDocument(assemblyPath, docType, true, out returnValue);
            if (swDoc == null || returnValue != SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
            {
                return false;
            }
            //this.AllComObjects.Enqueue(swDoc);
            SwDMConfigurationMgr dmConfigMgr = swDoc.ConfigurationManager;
            //this.AllComObjects.Enqueue(dmConfigMgr);
            string[] configurationNames = (string[])dmConfigMgr.GetConfigurationNames();
            if (configurationNames == null || configurationNames.Length <= 0)
            { 
                return false;
            }
            string configNameToOpen = null;
            if (string.IsNullOrEmpty(configurationName))
            { 
                configNameToOpen = dmConfigMgr.GetActiveConfigurationName();
            }
            else
            {
                configNameToOpen = configurationName;
            }
            SwDMConfiguration14 activeCfg = (SwDMConfiguration14)dmConfigMgr.GetConfigurationByName(configNameToOpen);
            if (activeCfg == null)
            {
                return false;
            }
            //this.AllComObjects.Enqueue(activeCfg);
            int topId = this.GetTopDocumentInfo(assemblyPath, swDoc, activeCfg);
            if (topId <= 0) 
            {
                return false;
            }

            if (docType == SwDmDocumentType.swDmDocumentAssembly)
            {
                try
                {
                    object[] allComponents = activeCfg.GetComponents();
                    if (allComponents != null)
                    {
                        foreach (object o in allComponents)
                        {
                            SwDMComponent9 subComponent = o as SwDMComponent9;
                            this.TraverseRecursively(subComponent, 1, topId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(string.Format("TraverseCore:exception:{0}", ex.Message));
                }
            }
            return true;
        }

        private void CopyFromPreviousSubAssembly(ComponentAndTask_Assembly assemblyComponentAndTask, int newLevel, int newParentId)
        {
            if (assemblyComponentAndTask == null || assemblyComponentAndTask.Id <= 0)
            {
                Debug.WriteLine(string.Format("CopyFromPreviousSubAssembly: invalid parameter"));
                return;
            }
            assemblyComponentAndTask.TheTask.Wait();
            SldDMComponentItem theAssemblyInfo = this.ResultsAtPopulating.FirstOrDefault(x => x.Id == assemblyComponentAndTask.Id);
            if (theAssemblyInfo == null)
            {
                Debug.WriteLine(string.Format("CopyFromPreviousSubAssembly: couldn't find sub assembly with Id={0}", assemblyComponentAndTask.Id));
                return;
            }
            SldDMComponentItem me = SldDMComponentItem.Clone(theAssemblyInfo, newLevel, newParentId);
            this.ResultsAtPopulating.Enqueue(me);

            IEnumerable<ComponentAndTask> subComponentsAndTasks = this.AllWalkingTasks.Where(x => x.ParentId == assemblyComponentAndTask.Id); 
            if (subComponentsAndTasks == null || subComponentsAndTasks.Count() <= 0)
                return; 
            Task[] allSubTasks = subComponentsAndTasks.Select(x => x.TheTask).ToArray();
            Task.WaitAll(allSubTasks); 
            foreach (var v in subComponentsAndTasks)
            {
                if (v is ComponentAndTask_Assembly) 
                    this.CopyFromPreviousSubAssembly(v as ComponentAndTask_Assembly, me.Level + 1, me.Id);
                else
                    this.CopyFromPreviousPart(v as ComponentAndTask_Part, me.Level + 1, me.Id);
            }
        }

        private void CopyFromPreviousPart(ComponentAndTask_Part previousComponentAndTask, int newLevel, int newParentId)
        {
            if (previousComponentAndTask == null || previousComponentAndTask.Id <= 0)
            {
                Debug.WriteLine(string.Format("CopyFromPreviousPart: invalid parameter"));
                return;
            }
            previousComponentAndTask.TheTask.Wait();
            SldDMComponentItem thePartInfo = this.ResultsAtPopulating.FirstOrDefault(x => x.Id == previousComponentAndTask.Id);
            if (thePartInfo == null)
            {
                Debug.WriteLine(string.Format("CopyFromPreviousPart: couldn't find component with Id={0}", previousComponentAndTask.Id));
                return;
            }
            SldDMComponentItem me = SldDMComponentItem.Clone(thePartInfo, newLevel, newParentId);
            this.ResultsAtPopulating.Enqueue(me);
        }

        protected bool TraverseRecursively(SwDMComponent9 currentComponent, int level, int parentId)
        {
            if (currentComponent == null || level <= 0)
                return false;
            try
            {
                string fullPath = currentComponent.PathName;
                ComponentAndTask componentAndTask = this.AllWalkingTasks.FirstOrDefault(x => (x is ComponentAndTask_Assembly) && string.Compare(fullPath, ((ComponentAndTask_Assembly)x).FullPath, true) == 0);
                if (componentAndTask != null)
                {
                    this.CopyFromPreviousSubAssembly(componentAndTask as ComponentAndTask_Assembly, level, parentId);
                    return true;
                }
                int myId = this.GetCurrentComponentInfo(currentComponent, fullPath, level, parentId);
                Debug.WriteLine(string.Format("TraverseRecursively:will handle ID={0}:{1}", myId, fullPath));
                if (currentComponent.DocumentType == SwDmDocumentType.swDmDocumentAssembly)
                {
                    List<SwDMComponent9> componentsToTraverseInTasks = new List<SwDMComponent9>();
                    SwDmDocumentOpenError openResult;
                    SwDMDocument17 doc = currentComponent.GetDocument2(true, this.GlobalSearchOption, out openResult) as SwDMDocument17;
                    if (doc == null || openResult != SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
                    {
                        Debug.WriteLine(string.Format("TraverseRecursively:failed to open{0}, return value is:{1}", fullPath, openResult.ToString()));
                        return false;
                    }
                    SwDMConfigurationMgr swDMConfigurationMgr = doc.ConfigurationManager;
                    string currentConfigName = currentComponent.ConfigurationName;
                    SwDMConfiguration2 config = swDMConfigurationMgr.GetConfigurationByName(currentConfigName) as SwDMConfiguration2;
                    object[] allComponents = null;
                    try
                    {
                        object componentsRaw = config.GetComponents();
                        if (componentsRaw != null)
                            allComponents = (object[])componentsRaw;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(string.Format("TraverseRecursively:exception occurred when dealing {0}, exception message:{1}", fullPath, ex.Message));
                        return false;
                    }
                    if (allComponents != null)
                    {
                        foreach (object o in allComponents)
                        {
                            SwDMComponent9 subComponent = o as SwDMComponent9;
                            //this.AllComObjects.Enqueue(subComponent);
                            if (subComponent.DocumentType == SwDmDocumentType.swDmDocumentAssembly)
                            { 
                                this.TraverseRecursively(subComponent, level + 1, myId);
                            }
                            else 
                            {
                                componentsToTraverseInTasks.Add(subComponent);
                            }
                        }
                    }

                    foreach (var v in componentsToTraverseInTasks)
                    { 
                        string subComponentPath = v.PathName;
                        ComponentAndTask subComponentAndTask = this.AllWalkingTasks.FirstOrDefault(x => string.Compare(subComponentPath, x.FullPath, true) == 0);
                        if (componentAndTask != null)
                        { 
                            this.CopyFromPreviousPart(componentAndTask as ComponentAndTask_Part, level, parentId);
                            return true;
                        }
                        else
                        {
                            this.GetCurrentComponentInfo(v, subComponentPath, level + 1, myId);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(string.Format("TraverseRecursively:when getting an item under the parent ID={0} {1}, exception occurred:{2}", parentId, currentComponent.PathName, ex.Message));
            }
            return true;
        }

        private void IterateBatch(IEnumerable<SwDMComponent9> items, int level, int parentId)
        {
            if (items == null || items.Count() <= 0 || level < 0)
                return;
            try
            {
                foreach (SwDMComponent9 activeComponent in items)
                {
                    SldDMComponentItem singlePartInfo = this.GetComponentInfo(activeComponent, level, parentId);
                    this.ResultsAtPopulating.Enqueue(singlePartInfo);
                    Console.WriteLine(string.Format("IterateBatch:遍历成功ID={0}的零件:{1}", singlePartInfo.Id, singlePartInfo.FullPath));
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(string.Format("IterateBatch:遍历ID={0}装配体下的零件时发生异常:{1}", parentId, ex.Message));
            }
        }

        private int GetTopDocumentInfo(
            string fullPath,
            SwDMDocument17 currentDocument,
            SwDMConfiguration14 currentConfig)
        {
            if (string.IsNullOrEmpty(fullPath) || currentDocument == null || currentConfig == null)
                return -1;
            SldDMComponentItem root = SldDMComponentItem.GenerateTheRoot();
            Task task = Task.Factory.StartNew(() =>
            {
                root.FullPath = fullPath;
                root.FeatureName = currentDocument.Title;
                root.ConfigurationCount = currentDocument.GetComponentCount();
                object[] subComponents = currentConfig.GetComponents();
                root.ChildrenCount = subComponents == null ? 0 : subComponents.Length;
                if (File.Exists(fullPath))
                {
                    root.AllCustomProperties = GetDocumentPropertiesViaDM.RetrievePropertiesUnderConfig(
                        fullPath,
                        currentConfig,
                        GetDocumentPropertiesViaDM.PropertyNamesOftenUsed,
                        true);
                }
                this.ResultsAtPopulating.Enqueue(root);
                this.CallbackGettingExtraInfo?.Invoke(ref root);
                Debug.WriteLine(string.Format("GetTopDocumentInfo:got [{0}]", fullPath));
            });
            this.AllWalkingTasks.Enqueue(new ComponentAndTask_Assembly()
            {
                ParentId = -1,
                TheTask = task,
                FullPath = fullPath,
                Id = root.Id
            });
            return root.Id;
        }

        private SldDMComponentItem GetComponentInfo(
            SwDMComponent9 currentComponent,
            int level,
            int parentId)
        {
            if (currentComponent == null || level < 0 || parentId <= 0)
                return null;
            string fullPath = currentComponent.PathName; ;
            SldDMComponentItem item = SldDMComponentItem.GenerateTheNext();
            item.ParentId = parentId;
            item.Level = level;
            try
            {
                item.FullPath = fullPath;
                item.FeatureName = currentComponent.Name2;
                item.FeatureId = currentComponent.GetID();
                item.IsSuppressed = currentComponent.IsSuppressed();
                item.FileName2D = EngineeringDrawingFile.Get2DFileNameIfExisting(fullPath);
                item.Visible = !currentComponent.IsHidden();

                SwDmDocumentOpenError openResult;
                SwDMDocument17 doc = currentComponent.GetDocument2(true, this.GlobalSearchOption, out openResult) as SwDMDocument17; 
                if (doc == null || openResult != SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
                {
                    Debug.WriteLine(string.Format("GetComponentInfo:couldn't open {0},the return value is:{1}", fullPath, openResult.ToString()));
                    return null;
                }
                SwDMConfiguration14 config = doc.ConfigurationManager.GetConfigurationByName(currentComponent.ConfigurationName) as SwDMConfiguration14;

                if (File.Exists(fullPath))
                    item.AllCustomProperties = GetDocumentPropertiesViaDM.RetrievePropertiesUnderConfig(
                        fullPath,
                        config,
                        GetDocumentPropertiesViaDM.PropertyNamesOftenUsed,
                        true);

                item.ConfigurationCount = doc.ConfigurationManager.GetConfigurationCount();
                item.ChildrenCount = doc.GetComponentCount();
                doc.CloseDoc();
                Marshal.ReleaseComObject(doc);
                this.CallbackGettingExtraInfo?.Invoke(ref item);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(string.Format("GetComponentInfo:when dealing {0}, exception occurred:{1}", fullPath, ex.Message));
            }
            return item;
        }

        private int GetCurrentComponentInfo(
            SwDMComponent9 currentComponent,
            string fullPath,
            int level,
            int parentId)
        {
            if (currentComponent == null || string.IsNullOrEmpty(fullPath) || level < 0 || parentId <= 0)
                return -1;
            SldDMComponentItem item = SldDMComponentItem.GenerateTheNext();
            Task task = Task.Factory.StartNew(() =>
            {
                item.ParentId = parentId;
                item.Level = level;
                item.FullPath = fullPath;
                item.FeatureName = currentComponent.Name2;
                item.FeatureId = currentComponent.GetID();
                item.IsSuppressed = currentComponent.IsSuppressed();
                item.FileName2D = EngineeringDrawingFile.Get2DFileNameIfExisting(fullPath);
                item.Visible = !currentComponent.IsHidden();

                try
                {
                    SwDmDocumentOpenError openResult;
                    SwDMDocument17 doc = currentComponent.GetDocument2(true, this.GlobalSearchOption, out openResult) as SwDMDocument17; 
                    if (doc == null || openResult != SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
                    {
                        Debug.WriteLine(string.Format("GetCurrentComponentInfo:couldn't open {0},the return value is :{1}", fullPath, openResult.ToString()));
                        return; 
                    }
                    SwDMConfiguration14 config = doc.ConfigurationManager.GetConfigurationByName(currentComponent.ConfigurationName) as SwDMConfiguration14;

                    if (File.Exists(fullPath))
                    {
                        item.AllCustomProperties = GetDocumentPropertiesViaDM.RetrievePropertiesUnderConfig(
                           fullPath,
                           config,
                           GetDocumentPropertiesViaDM.PropertyNamesOftenUsed,
                           true);
                    }
                    item.ConfigurationCount = doc.ConfigurationManager.GetConfigurationCount();
                    item.ChildrenCount = doc.GetComponentCount();
                    this.ResultsAtPopulating.Enqueue(item);
                    doc.CloseDoc();
                    Marshal.ReleaseComObject(doc);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(string.Format("GetCurrentComponentInfo:when dealing {0}, exception occurred:{1}", fullPath, ex.Message));
                    return;
                }
                this.CallbackGettingExtraInfo?.Invoke(ref item);
                Debug.WriteLine(string.Format("GetCurrentComponentInfo:got [{0}] (ID={1})", fullPath, item.Id));
            });
            string originalExt;
            SwDmDocumentType docType = SldFileExtentionChecker.CheckDM(fullPath, out originalExt);
            if (docType == SwDmDocumentType.swDmDocumentAssembly)
            {
                this.AllWalkingTasks.Enqueue(new ComponentAndTask_Assembly()
                {
                    ParentId = parentId,
                    TheTask = task,
                    FullPath = fullPath,
                    Id = item.Id
                });
            }
            else
            {
                this.AllWalkingTasks.Enqueue(new ComponentAndTask_Part()
                {
                    ParentId = parentId,
                    TheTask = task,
                    FullPath = fullPath,
                    Id = item.Id
                });
            }
            return item.Id;
        }
    }
}

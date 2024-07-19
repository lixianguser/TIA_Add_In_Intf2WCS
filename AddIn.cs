using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.HW;
using Siemens.Engineering.SW.Blocks;
using System;
using System.Collections.Generic;
#if DEBUG
            using System.Diagnostics;
#endif
using System.Globalization;
using System.Linq;
using System.IO;
using Siemens.Engineering.Online;
using System.Windows.Forms;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;

namespace TIA_Add_In_Intf2WCS
{
    public class AddIn : ContextMenuAddIn
    {
        #region Definition

        /// <summary>
        /// 博图实例
        /// </summary>
        private readonly TiaPortal _tiaPortal;

        /// <summary>
        /// Base class for projects
        /// can be used in multi-user environment
        /// </summary>
        private ProjectBase _projectBase;

        // /// <summary>
        // /// Path of the project file
        // /// </summary>
        // private string _projectPath;

        /// <summary>
        /// Path of the project directory
        /// </summary>
        private string _projectDir;

        /// <summary>
        /// 导出文件夹路径
        /// </summary>
        private string _exportFileDir;

        // /// <summary>
        // /// 导出文件夹信息
        // /// </summary>
        // private DirectoryInfo _exportDirInfo;

        // /// <summary>
        // /// 获取PLC目标
        // /// </summary>
        // private PlcSoftware _plcSoftware;

        /// <summary>
        /// 获取实例的DB名称
        /// </summary>
        private string _globalDbName;

        /// <summary>
        /// 获取DB的编号
        /// </summary>
        private string _number;

        /// <summary>
        /// 获取程序语言类型
        /// </summary>
        private string _programmingLanguage;

        /// <summary>
        /// 获取InstanceDB Xml文件路径
        /// </summary>
        private string _iDbXmlFilePath;

        /// <summary>
        /// 写入数据流
        /// </summary>
        private StreamWriter _streamWriter;

        /// <summary>
        /// csv文件保存路径
        /// </summary>
        private string _csvFilePath;

        #endregion

        public AddIn(TiaPortal tiaPortal) : base("WCS工具")
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<GlobalDB>("生成数据", Generate_OnClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("如需使用，请选中全局数据块", 
                menuSelectionProvider => { }, GenerateStatus);
        }

        private void Generate_OnClick(MenuSelectionProvider<GlobalDB> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            //获取项目数据
            GetProjectData();
            //确定PLC的在线状态
            if (!IsOffline())
            {
                throw new Exception(string.Format(CultureInfo.InvariantCulture,
                            "PLC 在线状态！"));
            }

            try
            {
                // 打开窗口获取.csv文件保存位置
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "逗号分隔值|*.csv",
                    Title = "请选择保存文件的路径"
                };

                if (saveFileDialog.ShowDialog(new Form()
                { TopMost = true, WindowState = FormWindowState.Maximized }) == DialogResult.OK)
                {
                    // 独占窗口
                    using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("导出中……"))
                    {
                        //定义csv数据流
                        _csvFilePath = saveFileDialog.FileName;
                        _streamWriter = new StreamWriter(_csvFilePath);

                        //写入标题行
                        _streamWriter.WriteLine("\"输送机编号\",\"变量名\",\"类型\",\"DB\",\"开始地址\",\"DB名称\"");

                        foreach (GlobalDB globalDb in menuSelectionProvider.GetSelection())
                        {
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                break;
                            }

                            // 获取 GlobalDB 名称
                            _globalDbName = globalDb.Name;
                            // 获取Number
                            _number = globalDb.Number.ToString();
                            // 获取ProgrammingLanguage
                            _programmingLanguage = globalDb.ProgrammingLanguage.ToString();

                            //创建导出文件夹
                            _exportFileDir = Path.Combine(_projectDir, "WCS");
                            if (!Directory.Exists(_exportFileDir))
                            {
                                Directory.CreateDirectory(_exportFileDir);
                            }
                            if (Directory.Exists(_exportFileDir))
                            {
                                // 导出DB块
                                _iDbXmlFilePath = Path.Combine(_exportFileDir, StringHandle(_globalDbName) + ".xml");
                                exclusiveAccess.Text = "导出中-> " + Export(globalDb, _iDbXmlFilePath);
                            }

                            XmlReader xmlReader = new XmlReader
                            {
                                DbXmlFilePath = _iDbXmlFilePath,
                                ProgrammingLanguage = _programmingLanguage,
                                Number = _number,
                                InstanceName = _globalDbName,
                                StreamWriter = _streamWriter
                            };
                            xmlReader.Run();
                        }
                        _streamWriter.Close();

                        //导出完成
                        MessageBox.Show($"目标文件:{saveFileDialog.FileName}", "导出完成", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //删除csv文件
                if (File.Exists(_csvFilePath))
                {
                    File.Delete(_csvFilePath);
                }
                throw;
            }
            finally
            {
                // 删除导出的文件夹
                DeleteDirectoryAndContents(_exportFileDir);
            }
        }
        
        /// <summary>
        /// 导出PLC数据
        /// </summary>
        /// <param name="exportItem"></param>
        /// <param name="exportPath"></param>
        /// <returns></returns>
        private string Export(IEngineeringObject exportItem, string exportPath)
        {
            const ExportOptions exportOption = ExportOptions.WithDefaults | ExportOptions.WithReadOnly;

            switch (exportItem)
            {
                case PlcBlock item:
                    {
                        if (item.ProgrammingLanguage == ProgrammingLanguage.ProDiag ||
                            item.ProgrammingLanguage == ProgrammingLanguage.ProDiag_OB)
                            return null;
                        if (item.IsConsistent)
                        {
                            // filePath = Path.Combine(filePath, AdjustNames.AdjustFileName(GetObjectName(item)) + ".xml");
                            if (File.Exists(exportPath))
                            {
                                File.Delete(exportPath);
                            }

                            item.Export(new FileInfo(exportPath), exportOption);

                            return exportPath;
                        }

                        throw new EngineeringException(string.Format(CultureInfo.InvariantCulture,
                            "程序块: {0} 是不一致的! 请编译程序块! 导出将终止!", item.Name));
                    }
            }

            return null;
        }

        /// <summary>
        /// 字符串处理将"/"替换为"_"
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private string StringHandle(string input)
        {
            //查询名称是否包含“/”，如果包含替换更“_”
            string ret = input;
            while (ret.Contains("/"))
            {
                ret = ret.Replace("/", "_");
            }
            return ret;
        }

        /// <summary>
        /// 删除文件夹及其内容
        /// </summary>
        /// <param name="targetDir"></param>
        private static void DeleteDirectoryAndContents(string targetDir)
        {
            if (!Directory.Exists(targetDir))
            {
                throw new DirectoryNotFoundException($"目录 {targetDir} 不存在。");
            }

            string[] files = Directory.GetFiles(targetDir);
            string[] dirs = Directory.GetDirectories(targetDir);

            // 先删除所有文件
            foreach (string file in files)
            {
                File.Delete(file);
            }

            // 然后递归删除所有子目录
            foreach (string dir in dirs)
            {
                DeleteDirectoryAndContents(dir);
            }

            // 最后，删除目录本身
            Directory.Delete(targetDir, false);
        }

        /// <summary>
        /// 获取ProjectBase：支持多用户项目
        /// </summary>
        private void GetProjectData()
        {
            try
            {
                // Multi-user support
                // If TIA Portal is in multiuser environment (connected to project server)
                if (_tiaPortal.LocalSessions.Any())
                {
                    _projectBase = _tiaPortal.LocalSessions
                        .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
                }
                else
                {
                    // Get local project
                    _projectBase = _tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
                }

                // _projectPath = _projectBase?.Path.FullName;
                _projectDir = _projectBase?.Path.Directory?.FullName;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        /// <summary>
        /// PLC是否为离线模式
        /// </summary>
        /// <returns></returns>
        private bool IsOffline()
        {
            bool ret = false;

            foreach (Device device in _projectBase.Devices)
            {
                DeviceItem deviceItem = device.DeviceItems[1];
                if (deviceItem.GetAttribute("Classification") is DeviceItemClassifications.CPU)
                {
                    OnlineProvider onlineProvider = deviceItem.GetService<OnlineProvider>();
                    ret = (onlineProvider.State == OnlineState.Offline);
                }
            }
            return ret;
        }

        /// <summary>
        /// 获取选中项类型关闭显示项目树按钮
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        private static MenuStatus GenerateStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            var show = false;

            foreach (IEngineeringObject engineeringObject in menuSelectionProvider.GetSelection())
            {
                if (!(engineeringObject is GlobalDB))
                {
                    show = true;
                    break;
                }
            }
            return show ? MenuStatus.Disabled : MenuStatus.Hidden;
        }

        /// <summary>
        /// AddIn Tester 插件调试器
        /// </summary>
        /// <param name="label"></param>
        /// <returns></returns>
        public IEnumerable<IEngineeringObject> GetSelection(string label)
        {
            PlcSoftware plcSoftware = null;
            ProjectBase projectBase;
            
            if (_tiaPortal.LocalSessions.Any())
            {
                projectBase = _tiaPortal.LocalSessions
                    .FirstOrDefault(s => s.Project != null && s.Project.IsPrimary)?.Project;
            }
            else
            {
                // Get local project
                projectBase = _tiaPortal.Projects.FirstOrDefault(p => p.IsPrimary);
            }

            if (projectBase != null)
            {
                foreach (Device device in projectBase.Devices)
                {
                    if (device.DeviceItems[1].GetAttribute("Classification") is DeviceItemClassifications.CPU)
                    {
                        DeviceItemComposition deviceItemComposition = device.DeviceItems;
                        foreach (DeviceItem deviceItem in deviceItemComposition)
                        {
                            SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                            if (softwareContainer != null)
                            {
                                Software softwareBase = softwareContainer.Software;
                                plcSoftware = softwareBase as PlcSoftware;
                            }
                        }
                    }
                }
            }

            var selection = new List<IEngineeringObject>();

            if (plcSoftware != null)
            {
                selection.Add(plcSoftware.BlockGroup.Blocks.Find("数据块_1"));
            }

            return selection;
        }
    }
}

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace TIA_Add_In_Intf2WCS
{
    public class XmlReader
    {
        /// <summary>
        /// 获取Xml的命名空间
        /// </summary>
        private XmlNamespaceManager _xmlns;

        /// <summary>
        /// 获取实例的DB名称
        /// </summary>
        public string InstanceName;

        // /// <summary>
        // /// 获取实例的FB名称
        // /// </summary>
        // public string InstanceOfName;

        /// <summary>
        /// 获取DB的编号
        /// </summary>
        public string Number;

        /// <summary>
        /// 获取程序语言类型
        /// </summary>
        public string ProgrammingLanguage;

        /// <summary>
        /// 获取InstanceDB Xml文件路径
        /// </summary>
        public string DbXmlFilePath;

        // /// <summary>
        // /// 获取FB Xml文件路径
        // /// </summary>
        // public string FBXmlFilePath;

        /// <summary>
        /// 写入数据流
        /// </summary>
        public StreamWriter StreamWriter;

        public void Run()
        {
            //读取Xml文件
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(DbXmlFilePath);
            //获取"Sections"节点
            XmlNode sections = xmlDocument.GetElementsByTagName("Sections")[0];
            //从Sections和BlockInstSupervisionGroups节点获取NameSpace
            _xmlns = GetXmlns(xmlDocument, sections.NamespaceURI);

            //获取变量和偏移量
            foreach (XmlNode members in sections)
            {
                foreach (XmlNode member in members)
                {
                    var iStruct = GetAttribute(member, "Name");
                    string iInterface;
                    string type;
                    string offset;
                    string length;
                    switch (GetAttribute(member, "Datatype"))
                    {
                        case "Struct":
                            foreach (XmlNode item in GetMember(member))
                            {
                                iInterface = GetAttribute(item, "Name");
                                type = GetAttribute(item, "Datatype");
                                offset = CalOffset(GetOffset(member) + GetOffset(item));
                                length = GetLength(type);
                                SetSacdaTag(iStruct, iInterface, type, offset, length);
                            }
                            break;
                        case "\"Conv_To_WCS\"":
                        case "\"WCS_To_Conv\"":
                            foreach (XmlNode item in GetSection(member))
                            {
                                iInterface = GetAttribute(item, "Name");
                                type = GetAttribute(item, "Datatype");
                                offset = CalOffset(GetOffset(member) + GetOffset(item));
                                length = GetLength(type);
                                SetSacdaTag(iStruct, iInterface, type, offset, length);
                            }
                            break;
                        default:
                                iInterface = GetAttribute(member, "Name");
                                type = GetAttribute(member, "Datatype");
                                offset = CalOffset(GetOffset(member));
                                length = GetLength(type);
                                SetSacdaTag(iStruct, iInterface, type, offset, length);
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 获取命名空间
        /// </summary>
        /// <param name="xmlDocument"></param>
        /// <param name="namespaceUri">Sections命名空间</param>
        /// <returns>xmlns="http://www.siemens.com/automation/Openness/SW/Interface/v5"</returns>
        private XmlNamespaceManager GetXmlns(XmlDocument xmlDocument, string namespaceUri)
        {
            XmlNamespaceManager xmlns = new XmlNamespaceManager(xmlDocument.NameTable);
            xmlns.AddNamespace("x", namespaceUri);
            return xmlns;
        }

        /// <summary>
        /// 获取节点的元素的值
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <param name="nameItem"></param>
        /// <returns>Name="MsgType" 或 Datatype="Int"</returns>
        private static string GetAttribute(XmlNode xmlNode, string nameItem)
        {
            var name = xmlNode.Attributes?.GetNamedItem(nameItem).Value;
            return name;
        }

        /// <summary>
        /// 获取Section
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNode GetSection(XmlNode xmlNode)
        {
            XmlNode section = xmlNode.SelectSingleNode("./x:Sections", _xmlns)?.SelectSingleNode("./x:Section", _xmlns);
            return section;
        }

        /// <summary>
        /// 获取Member
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns></returns>
        private XmlNodeList GetMember(XmlNode xmlNode)
        {
            XmlNodeList section = xmlNode.SelectNodes("./x:Member", _xmlns);
            return section;
        }

        /// <summary>
        /// 计算偏移量
        /// </summary>
        /// <param name="getOffset"></param>
        /// <returns>40.1</returns>
        private static string CalOffset(int getOffset)
        {
            //计算偏移量
            var consult = (getOffset / 8).ToString(); //商，"40"
            var rem = (getOffset % 8).ToString(); //余数，"1"
            var offset = $"{consult}.{rem}"; //偏移量结果：40.1
            // Console.WriteLine(offset);

            return offset;
        }

        /// <summary>
        /// 获取偏移量节点
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <returns>321</returns>
        private int GetOffset(XmlNode xmlNode)
        {
            XmlNode integerAttribute = xmlNode.SelectSingleNode("./x:AttributeList", _xmlns)?.FirstChild.FirstChild;
            if (integerAttribute == null) return 0;
            var offset = Convert.ToInt32(integerAttribute.Value);
            return offset;
        }

        /// <summary>
        /// 设置数据类型长度
        /// </summary>
        /// <param name="dataType"></param>
        /// <returns></returns>
        private static string GetLength(string dataType)
        {
            string ret = null;

            if (dataType.Contains("Array["))
            {
                //起始数值的正则表达式
                string patternStart = @"\[(.*?)\.\.";
                //结束数值的正则表达式
                string patternEnd = @"\.\.(.*?)\]";

                // 使用正则表达式进行匹配
                Match matchStart = Regex.Match(dataType, patternStart);
                Match matchEnd = Regex.Match(dataType, patternEnd);
                
                // 提取匹配的组中的值
                string resultStart = matchStart.Groups[1].Value;
                // 尝试将字符串转换为整数
                int.TryParse(resultStart, out int startNumber);
                
                // 提取匹配的组中的值
                string resultEnd = matchEnd.Groups[1].Value;
                // 尝试将字符串转换为整数
                int.TryParse(resultEnd, out int endNumber);

                //计算差值
                int sub = endNumber - startNumber + 1;
                
                if (dataType.Contains("Bool"))
                {
                    ret = 1 * sub < 8 ? $"0.{1 * sub}" : CalOffset(1 * sub);
                }
                if (dataType.Contains("Byte") || dataType.Contains("SInt") || dataType.Contains("UInt") || dataType.Contains("Char"))
                {
                    ret = (1 * sub).ToString();
                }
                if (dataType.Contains("Int") || dataType.Contains("UInt") || dataType.Contains("Word"))
                {
                    ret = (2 * sub).ToString();
                }
                if (dataType.Contains("DInt") || dataType.Contains("DWord"))
                {
                    ret = (4 * sub).ToString();
                }
                if (dataType.Contains("LWord"))
                {
                    ret = (8 * sub).ToString();
                }
                if (dataType.Contains("String"))
                {
                    ret = (254 * sub).ToString();
                }
            }
            else
            {
                switch (dataType)
                {
                    case "Bool":
                        ret = "0.1";
                        break;
                    case "Byte":
                    case "SInt":
                    case "UInt":
                        ret = "1";
                        break;
                    case "Int":
                    case "Uint":
                    case "Word":
                        ret = "2";
                        break;
                    case "DInt":
                    case "DWord":
                        ret = "4";
                        break;
                    case "LWord":
                        ret = "8";
                        break;
                    case "String":
                        ret = "254";
                        break;
                    default:
                        ret = "0";
                        break;
                }
            }
            
            return ret;
        }

        /// <summary>
        /// 设置WCS接口值
        /// </summary>
        /// <param name="iStruct">输送机编号</param>
        /// <param name="iInterface">接口名称</param>
        /// <param name="datatype">数据类型</param>
        /// <param name="offset">偏移量</param>
        /// <param name="length">长度</param>
        /// <returns>"P1001","MsgType","Int","DB83","0.0","PLC_To_WCS","2"</returns>
        private void SetSacdaTag(string iStruct, string iInterface, string datatype, string offset, string length)
        {
            var str =
                $"\"{iStruct}\"," +
                $"\"{iInterface}\"," +
                $"\"{datatype}\"," +
                $"\"{ProgrammingLanguage}{Number}\"," +
                $"\"{offset}\"," +
                $"\"{InstanceName}\",\"{length}\"";
            StreamWriter.WriteLine(str);
            Console.WriteLine(str);
        }
    }
}

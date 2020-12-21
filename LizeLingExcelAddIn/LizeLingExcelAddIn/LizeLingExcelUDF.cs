using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.IntelliSense;
using System.IO;
using ExcelDna.Integration.CustomUI;
using System.Text.RegularExpressions;

namespace LizeLingExcelAddIn
{
    public class LizeLingExcelUDF
    {
        readonly static String ERRORMessage = "#ERROR";

        [ExcelFunction(Name = "TEXTSPLIT", Description = "지정된 문자열을 지정된 문자로 잘라서 반환합니다.", IsVolatile = true)]
        public static object TextSplit(
             [ExcelArgument(Name = "SplitText", Description = "자를 문자열 원본입니다.")]
            String SplitMessage,
             [ExcelArgument(Name = "SplitChar", Description = "원본 문자열을 자를 문자입니다.")]
            String SplitChar,
             [ExcelArgument(Name = "[Number]", Description = "반환할 배열의 번호 입니다.")]
            double Number = 0)
        {
            try
            {
                if (Number == 0)
                {
                    object[] a = SplitMessage.Split(SplitChar.ToCharArray()[0]);
                    object[,] retArray = new object[1, a.Length];
                    for (int i = 0; i < a.Length; i++)
                    {
                        retArray[0, i] = a[i];
                    }
                    return retArray;
                }
                else
                {
                    return SplitMessage.Split(SplitChar.ToCharArray()[0])[Convert.ToInt32(Number - 1)];
                }
            }
            catch
            {
                return ERRORMessage;
            }
        }

        [ExcelFunction(Name = "GETFILES", Description = "지정된 경로에 있는 파일리스트를 반환합니다.(Office 2019이상)", IsVolatile = true)]
        public static object GetFiles(
            [ExcelArgument(Name = "DirectoryPath", Description = "파일을 불러올 디렉토리 경로입니다.")]
            String DirectoryPath,
            [ExcelArgument(Name = "[SearchPattern]", Description = "검색할 파일 패턴입니다.")]
            String SearchPattern = "",
            [ExcelArgument(Name = "[AllDirectories]", Description = "하위 디렉토리의 파일까지 가져올지 여부입니다.")]
            bool AllDirectories = false,
            [ExcelArgument(Name = "[NameType]", Description = "0(기본) : 확장자를 제외한 이름만 출력합니다.\n1 : 확장자를 포함하여 출력합니다.\n2 : 파일 경로를 포함하여 출력합니다.")]
            double NameType = 0)
        {
            try
            {
                SearchOption option = SearchOption.TopDirectoryOnly;

                if (AllDirectories)
                {
                    option = SearchOption.AllDirectories;
                }

                if (String.IsNullOrEmpty(SearchPattern))
                {
                    SearchPattern = "*.*";
                }

                DirectoryInfo directoryInfo = new DirectoryInfo(DirectoryPath);
                FileInfo[] Files = directoryInfo.GetFiles(SearchPattern, option);
                object[,] retArray = new object[Files.Length, 1];
                for (int i = 0; i < Files.Length; i++)
                {
                    switch (NameType)
                    {
                        case 0: retArray[i, 0] = Path.GetFileNameWithoutExtension(Files[i].FullName); break;
                        case 1: retArray[i, 0] = Files[i].Name; break;
                        case 2: retArray[i, 0] = Files[i].FullName; break;
                        default: return ERRORMessage;
                    }
                }
                return retArray;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        [ExcelFunction(Name = "REGEXP", Description = "정규식을 적용하여 일치하는 문자열을 가져옵니다.", IsVolatile = true)]
        public static object REGEXP(
            [ExcelArgument(Name = "RegExText", Description = "정규식의 추출 대상이 되는 문자열입니다.")]
            String RegExText,
            [ExcelArgument(Name = "RegPattern", Description = "적용할 정규식의 패턴입니다.")]
            String RegPattern,
            [ExcelArgument(Name = "Number", Description = "정규식을 적용하여 추출할 문자열의 번호입니다. 입력하지 않으면 모든 문자를 반환합니다.")]
            double Number = 0)
        {
            try
            {
                Regex reg = new Regex(RegPattern);
                if (Number == 0)
                {
                    object[,] Matches = new object[1, reg.Matches(RegExText).Count];
                    int index = 0;
                    foreach (Match match in reg.Matches(RegExText))
                    {
                        Matches[0, index] = match.Value;
                        index++;
                    }
                    return Matches;
                }
                else
                {
                    return reg.Matches(RegExText)[Convert.ToInt32(Number - 1)].Value;
                }
            }
            catch
            {
                return ERRORMessage;
            }
        }

        [ExcelFunction(Name = "REGEXPREPLACE", Description = "정규식을 적용하여 일치하는 문자열을 지정된 문자열로 대처합니다.", IsVolatile = true)]
        public static object REGEXPREPLACE(
            [ExcelArgument(Name = "RegExText", Description = "정규식의 추출 대상이 되는 문자열입니다.")]
            String RegExText,
            [ExcelArgument(Name = "RegPattern", Description = "적용할 정규식의 패턴입니다.")]
            String RegPattern,
            [ExcelArgument(Name = "ReplaceText", Description = "대처할 문자열입니다.")]
            String ReplaceText
            )
        {
            try
            {
                Regex reg = new Regex(RegPattern);
                return reg.Replace(RegExText, ReplaceText);
            }
            catch
            {
                return ERRORMessage;
            }
        }
    }

    /// <summary>
    /// 에드인 작동에 필요한 클래스
    /// </summary>
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }


    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                <ribbon>
                    <tabs>
                        <tab id='tab1' label='LizeLingAddin'>
                            <group id='group1' label='LizeLingAddin1'>
                                <button id='button1' label='TextButton1' onAction='TextButton1Action'/>
                            </group >
                            <group id='group2' label='LizeLingAddin2'>
                                <button id='button2' label='TextButton2' onAction='TextButton2Action'/>
                            </group >
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>";
        }
        public void TextButton1Action(IRibbonControl control)
        {
            Form1 form1 = new Form1();
            form1.Show();
        }
        public void TextButton2Action(IRibbonControl control)
        {

        }
    }
}

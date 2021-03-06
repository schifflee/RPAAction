using Microsoft.Vbe.Interop;
using RPAAction.Base;
using System;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 高阶-运行Excel宏
    /// </summary>
    public class HighLevel_RunMacro : ExcelAction
    {
        /// <param name="wbPath"></param>
        /// <param name="wsName"></param>
        /// <param name="VBACode"></param>
        /// <param name="MacroName">默认执行名称为"m"的宏</param>
        public HighLevel_RunMacro(string wbPath = null, string wsName = null, string VBACode = null, string MacroName = null)
            : base(wbPath, wsName)
        {
            this.VBACode = VBACode;
            this.MacroName = CheckString(MacroName) ? MacroName : "m";
            Run();
        }

        protected override void Action()
        {
            base.Action();
            //运行宏
            if (!CheckString(MacroName))
            {
                wb.Activate();
                try
                {
                    RunVBA();
                }
                //沒有信任存取VAB專案物件模型
                catch (System.Runtime.InteropServices.COMException come)
                {
                    //插入VBA代码
                    if (!CheckString(VBACode))
                    {
                        try
                        {
                            VBE vbe = app.VBE;
                            VBComponent vbComponent;
                            vbComponent = wb.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                            vbComponent.CodeModule.AddFromString(VBACode);
                        }
                        catch (Exception e)
                        {
                            throw new ActionException("添加vb函數失敗\n" + e.ToString());
                        }
                        RunVBA();
                    }
                    else
                    {
                        throw come;
                    }
                }
            }
        }

        private void RunVBA()
        {
            app.Run($@"'{wbFileName}'!{MacroName}");
        }

        private readonly string VBACode = null;
        private readonly string MacroName = null;
    }
}

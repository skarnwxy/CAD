/** 
  * Copyright (C) Skarn 2022
  * @file     CSharpAndCubeMode                                                       
  * @brief    创建立方体                                                   
  * @author   debugSarn                                                     
  * @date     2022-11-29  
  * @note     
  * @see      
 */

// 用于引用系统的API
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 用于引用SolidWorks相关API
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

// 用于引用Inventor相关API
using Inventor;

namespace CSharpAndCubeMode
{
    public partial class Form1 : Form
    {
        //---------------------------成员变量------------------------------
        // 创建模型的方式（0: 表示通过SolidWorks创建立方体模型；1: 表示通过Inventor创建立方体模型）
        int typeMode = 0;

        // 装配的层级
        int asmLevel = 1;
        AssemblyDocument sldAsmDocCube = null;
        ModelDoc2 invAsmDocCube = null;

        static int asmSum = 0;

        // X、Y、Z方向的立方体的个数
        int xNum = 0, yNum = 0, zNum = 0;

        // 立方体之间的间隙
        int cubeGap = 0;

        // 模型默认的公共路径
        string strCurPath = null;

        // 模板文档的路径
        string strAsmTemplatePath = null;
        string strPartTemplatePath = null;

        string strAsmTargetPath = null;
        string strPartTargetPath = null;

        //---------------------------成员函数------------------------------
        /** 
          * @brief        初始化文档的路径
          * @see          
          * @note          
         */
        private void initDocPath()
        {
            typeMode = comboBox1.SelectedIndex;
            if (typeMode == 1)
            {
                // 模板文档的路径
                strPartTemplatePath = strCurPath + "\\" + "ModeTemplate" + "\\" + "Standard.ipt";
                strAsmTemplatePath = strCurPath + "\\" + "ModeTemplate" + "\\" + "Standard.iam";
                
                // 目标文档的路径
                strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley";
                strPartTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Extrusion";
            }
            else
            {
                strPartTemplatePath = strCurPath + "\\" + "ModeTemplate" + "\\" + "零件.prtdot";
                strAsmTemplatePath = strCurPath + "\\" + "ModeTemplate" + "\\" + "装配.asmdot";

                strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "Assembley";
                strPartTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "Extrusion";
            }
        }

        /** 
          * @brief        构造函数，用于初始化成员变量
          * @see          
          * @note          
         */
        public Form1()
        {
            InitializeComponent();

            // 控件初始化
            button1.Enabled = false;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;

            // 相关路径设置
            strCurPath = System.Environment.CurrentDirectory;
            initDocPath();
        }

        //---------------------------公共函数------------------------------
        /** 
          * @brief        获取TextBox中的值
          * @see          
          * @note          
         */
        private void retrieveTextBoxValue()
        {
            typeMode = comboBox1.SelectedIndex;

            string strValue = comboBox2.Text;
            int.TryParse(strValue, out asmLevel);

            strValue = textBox1.Text;
            int.TryParse(strValue, out xNum);
            strValue = textBox2.Text;
            int.TryParse(strValue, out yNum);
            strValue = textBox3.Text;
            int.TryParse(strValue, out zNum);

            strValue = comboBox3.Text;
            int.TryParse(strValue, out cubeGap);
        }

        ///////////////////////////////SolidWorks///////////////////////////////////
        /** 
          * @brief        根据装配层级创建部件
          * @see          
          * @note          
         */
        private ModelDoc2 createAsmFirstLevel(SldWorks swApp, ref string ostrAsmTargetPathTemp)
        {
            int errors = 0;
            int warinings = 0;

            // 步骤1: 新建父层级
            ++asmSum;
            // 新建文档
            swApp.NewDocument(strAsmTemplatePath, 0, 0, 0);
            ModelDoc2 asmDocTemp = (ModelDoc2)swApp.ActiveDoc;
            // 保存文档
            string strAsmTargetPathTemp = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleyFirst";
            strAsmTargetPathTemp += (asmSum.ToString());
            strAsmTargetPathTemp += ".SLDASM";
            asmDocTemp.Extension.SaveAs(strAsmTargetPathTemp, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, "", ref errors, ref warinings);
            ostrAsmTargetPathTemp = strAsmTargetPathTemp;

            string strDocNameTemp = "AssembleyFirst";
            strDocNameTemp += (asmSum.ToString());
            strDocNameTemp += ".SLDASM";

            // 步骤2: 根据层级生成父层级的子层级
            for (int i = 1; i <= asmLevel; ++i)
            {
                ++asmSum;

                // 新建文档
                swApp.NewDocument(strAsmTemplatePath, 0, 0, 0);
                ModelDoc2 asmDocTemp1 = (ModelDoc2)swApp.ActiveDoc;

                // 保存文档
                string strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "Assembley";
                strAsmTargetPath += (asmSum.ToString());
                strAsmTargetPath += ".SLDASM";
                asmDocTemp1.Extension.SaveAs(strAsmTargetPath, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, "", ref errors, ref warinings);

                // 添加装配
                swApp.OpenDoc(strAsmTargetPath, (int)swDocumentTypes_e.swDocPART);

                AssemblyDoc assemblyDoc = (AssemblyDoc)swApp.ActivateDoc3(strDocNameTemp, true, 0, errors);
                swApp.ActivateDoc(strDocNameTemp);
                Component2 InsertedComponent = assemblyDoc.AddComponent5(strAsmTargetPath, 0, "", false, "", 0, 0, 0);
                InsertedComponent.Select(false);
                assemblyDoc.UnfixComponent();

                invAsmDocCube = asmDocTemp1;
                swApp.CloseDoc(strAsmTargetPath);
            }

            //swApp.CloseDoc(strAsmTargetPathTemp);
            asmDocTemp.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warinings);
            return asmDocTemp;
        }

        private ModelDoc2 createAsmSecondLevel(SldWorks swApp, ref string ostrAsmTargetPathTemp)
        {
            int errors = 0;
            int warinings = 0;

            // 步骤1: 新建父层级
            ++asmSum;
            // 新建文档
            swApp.NewDocument(strAsmTemplatePath, 0, 0, 0);
            ModelDoc2 asmDocTemp = (ModelDoc2)swApp.ActiveDoc;
            // 保存文档
            string strAsmTargetPathTemp = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleySeconcd";
            strAsmTargetPathTemp += (asmSum.ToString());
            strAsmTargetPathTemp += ".SLDASM";
            asmDocTemp.Extension.SaveAs(strAsmTargetPathTemp, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, "", ref errors, ref warinings);
            ostrAsmTargetPathTemp = strAsmTargetPathTemp;

            string strDocNameTemp = "AssembleySeconcd";
            strDocNameTemp += (asmSum.ToString());
            strDocNameTemp += ".SLDASM";

            // 步骤2: 根据层级生成父层级的子层级
            for (int i = 1; i <= asmLevel; ++i)
            {
                string strAsmTargetPath = null;
                ModelDoc2 asmDocTemp1 = createAsmFirstLevel(swApp, ref strAsmTargetPath);

                // 添加装配
                //string strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleyFirst";
                //strAsmTargetPath += (asmSum-asmLevel).ToString();
                //strAsmTargetPath += ".SLDASM";
                swApp.OpenDoc(strAsmTargetPath, (int)swDocumentTypes_e.swDocPART);

                AssemblyDoc assemblyDoc = (AssemblyDoc)swApp.ActivateDoc3(strDocNameTemp, true, 0, errors);
                swApp.ActivateDoc(strDocNameTemp);
                Component2 InsertedComponent = assemblyDoc.AddComponent5(strAsmTargetPath, 0, "", false, "", 0, 0, 0);
                InsertedComponent.Select(false);
                assemblyDoc.UnfixComponent();

                swApp.CloseDoc(strAsmTargetPath);
            }

            //swApp.CloseDoc(strAsmTargetPathTemp);
            asmDocTemp.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warinings);
            return asmDocTemp;
        }

        private ModelDoc2 createAsmThirdLevel(SldWorks swApp, ref string ostrAsmTargetPathTemp)
        {
            int errors = 0;
            int warinings = 0;

            // 步骤1: 新建父层级
            ++asmSum;
            // 新建文档
            swApp.NewDocument(strAsmTemplatePath, 0, 0, 0);
            ModelDoc2 asmDocTemp = (ModelDoc2)swApp.ActiveDoc;
            // 保存文档
            string strAsmTargetPathTemp = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleyThird";
            strAsmTargetPathTemp += (asmSum.ToString());
            strAsmTargetPathTemp += ".SLDASM";
            asmDocTemp.Extension.SaveAs(strAsmTargetPathTemp, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, "", ref errors, ref warinings);
            ostrAsmTargetPathTemp = strAsmTargetPathTemp;

            string strDocNameTemp = "AssembleyThird";
            strDocNameTemp += (asmSum.ToString());
            strDocNameTemp += ".SLDASM";

            // 步骤2: 根据层级生成父层级的子层级
            for (int i = 1; i <= asmLevel; ++i)
            {
                string strAsmTargetPath = null;
                ModelDoc2 asmDocTemp1 = createAsmSecondLevel(swApp, ref strAsmTargetPath);

                // 添加装配
                //string strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleySecond";
                //strAsmTargetPath += (asmSum - asmLevel*3).ToString();
                //strAsmTargetPath += ".SLDASM";
                swApp.OpenDoc(strAsmTargetPath, (int)swDocumentTypes_e.swDocPART);

                AssemblyDoc assemblyDoc = (AssemblyDoc)swApp.ActivateDoc3(strDocNameTemp, true, 0, errors);
                swApp.ActivateDoc(strDocNameTemp);
                Component2 InsertedComponent = assemblyDoc.AddComponent5(ostrAsmTargetPathTemp, 0, "", false, "", 0, 0, 0);
                InsertedComponent.Select(false);
                assemblyDoc.UnfixComponent();

                swApp.CloseDoc(strAsmTargetPath);
            }

            //swApp.CloseDoc(strAsmTargetPathTemp);
            asmDocTemp.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warinings);
            return asmDocTemp;
        }

        private ModelDoc2 createAsmFourthLevel(SldWorks swApp, ref string ostrAsmTargetPathTemp)
        {
            int errors = 0;
            int warinings = 0;

            // 步骤1: 新建父层级
            ++asmSum;
            // 新建文档
            swApp.NewDocument(strAsmTemplatePath, 0, 0, 0);
            ModelDoc2 asmDocTemp = (ModelDoc2)swApp.ActiveDoc;
            // 保存文档
            string strAsmTargetPathTemp = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleyFourth";
            strAsmTargetPathTemp += (asmSum.ToString());
            strAsmTargetPathTemp += ".SLDASM";
            asmDocTemp.Extension.SaveAs(strAsmTargetPathTemp, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, "", ref errors, ref warinings);
            ostrAsmTargetPathTemp = strAsmTargetPathTemp;

            string strDocNameTemp = "AssembleyFourth";
            strDocNameTemp += (asmSum.ToString());
            strDocNameTemp += ".SLDASM";

            // 步骤2: 根据层级生成父层级的子层级
            for (int i = 1; i <= asmLevel; ++i)
            {
                string strAsmTargetPath = null;
                ModelDoc2 asmDocTemp1 = createAsmThirdLevel(swApp, ref strAsmTargetPath);

                // 添加装配
                //string strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleySecond";
                //strAsmTargetPath += (asmSum - asmLevel*3).ToString();
                //strAsmTargetPath += ".SLDASM";
                swApp.OpenDoc(strAsmTargetPath, (int)swDocumentTypes_e.swDocPART);

                AssemblyDoc assemblyDoc = (AssemblyDoc)swApp.ActivateDoc3(strDocNameTemp, true, 0, errors);
                swApp.ActivateDoc(strDocNameTemp);
                Component2 InsertedComponent = assemblyDoc.AddComponent5(strAsmTargetPath, 0, "", false, "", 0, 0, 0);
                InsertedComponent.Select(false);
                assemblyDoc.UnfixComponent();

                swApp.CloseDoc(strAsmTargetPath);
            }

            //swApp.CloseDoc(strAsmTargetPathTemp);
            asmDocTemp.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warinings);
            return asmDocTemp;
        }

        private ModelDoc2 createAsmFifthLevel(SldWorks swApp, ref string ostrAsmTargetPathTemp)
        {
            int errors = 0;
            int warinings = 0;

            // 步骤1: 新建父层级
            ++asmSum;
            // 新建文档
            swApp.NewDocument(strAsmTemplatePath, 0, 0, 0);
            ModelDoc2 asmDocTemp = (ModelDoc2)swApp.ActiveDoc;
            // 保存文档
            string strAsmTargetPathTemp = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleyFifth";
            strAsmTargetPathTemp += (asmSum.ToString());
            strAsmTargetPathTemp += ".SLDASM";
            asmDocTemp.Extension.SaveAs(strAsmTargetPathTemp, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, "", ref errors, ref warinings);
            ostrAsmTargetPathTemp = strAsmTargetPathTemp;

            string strDocNameTemp = "AssembleyFifth";
            strDocNameTemp += (asmSum.ToString());
            strDocNameTemp += ".SLDASM";

            // 步骤2: 根据层级生成父层级的子层级
            for (int i = 1; i <= asmLevel; ++i)
            {
                string strAsmTargetPath = null;
                ModelDoc2 asmDocTemp1 = createAsmFourthLevel(swApp, ref strAsmTargetPath);

                // 添加装配
                //string strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "AssembleySecond";
                //strAsmTargetPath += (asmSum - asmLevel*3).ToString();
                //strAsmTargetPath += ".SLDASM";
                swApp.OpenDoc(strAsmTargetPath, (int)swDocumentTypes_e.swDocPART);

                AssemblyDoc assemblyDoc = (AssemblyDoc)swApp.ActivateDoc3(strDocNameTemp, true, 0, errors);
                swApp.ActivateDoc(strDocNameTemp);
                Component2 InsertedComponent = assemblyDoc.AddComponent5(strAsmTargetPath, 0, "", false, "", 0, 0, 0);
                InsertedComponent.Select(false);
                assemblyDoc.UnfixComponent();

                swApp.CloseDoc(strAsmTargetPath);
            }

            //swApp.CloseDoc(strAsmTargetPathTemp);
            asmDocTemp.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warinings);
            return asmDocTemp;
        }

        /** 
          * @brief        通过SolidWorks中的API创建立方体
          * @see          
          * @note          
         */
        private void createCubeFromSolidWorks()
        {
            // step 1: 创建一个SldWorks实例
            SldWorks swApp = null;
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                swApp = new SldWorks();
                //Type solidWorksApp = System.Type.GetTypeFromProgID("Inventor.Application");
                //swApp = (SldWorks)Activator.CreateInstance(solidWorksApp);
                if (swApp == null)
                {
                    return;
                }
                //swApp.Visible = true;
            }

            int errors = 0;
            int warinings = 0;

            // step 2: 新建装配文档
            ModelDoc2 AsmDoc = null;
            string ostrAsmTargetPathTemp = null;
            if (asmLevel == 1)
            {
                AsmDoc = createAsmFirstLevel(swApp, ref ostrAsmTargetPathTemp);
            }
            else if (asmLevel == 2)
            {
                AsmDoc = createAsmSecondLevel(swApp, ref ostrAsmTargetPathTemp);
            }
            else if (asmLevel == 3)
            {
                AsmDoc = createAsmThirdLevel(swApp, ref ostrAsmTargetPathTemp);
            }
            else if (asmLevel == 4)
            {
                AsmDoc = createAsmFourthLevel(swApp, ref ostrAsmTargetPathTemp);
            }
            else if (asmLevel == 5)
            {
                AsmDoc = createAsmFifthLevel(swApp, ref ostrAsmTargetPathTemp);
            }

            // step 3: 创建立方体(根据指定参数生成立方体)
            int curSum = 0;
            for (int i = 1; i <= zNum; ++i)
            {
                for (int j = 1; j <= yNum; ++j)
                {
                    for (int k = 1; k <= xNum; ++k)
                    {
                        // 3 - 1: 创建零件文档
                        //strCurPath = "C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2021\\templates\\装配.asmdot";
                        var newDoc = swApp.NewDocument(strPartTemplatePath, 0, 0, 0);
                        if (newDoc == null)
                        {
                            MessageBox.Show("零件文档未生成成功，请选择正确的模板.");
                            swApp.ExitApp(); swApp = null;
                            return;
                        }
                        ModelDoc2 PartDoc = (ModelDoc2)swApp.ActiveDoc;

                        // 3 - 2: 创建草图
                        PartDoc.SketchManager.InsertSketch(true); //插入一个草图
                        PartDoc.Extension.SelectByID2("前视基准面", "PLANE", -5.51688046000562E-02, 6.20552978418983E-02, -1.37730216225479E-03, false, 0, null, 0);//选择前视基准面
                        PartDoc.ClearSelection2(true);
                        PartDoc.SketchManager.CreateCenterRectangle(0, 0, 0, -2.56018390836871E-02, 2.57775321181576E-02, 0); //画一个方形

                        PartDoc.ShowNamedView2("*上下二等角轴测", 8);
                        PartDoc.ClearSelection2(true);

                        // 3 - 3: 创建拉伸体
                        PartDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.050, 0.01, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, true, true, true, 0, 0, false);

                        // 3 - 4: 保存模型
                        //string strPartTargetPath = "C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS 2021\\templates\\Demo\\Cubes\\testCubes\\Extrusion";
                        strPartTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "Extrusion";

                        ++curSum;
                        string strIndex = curSum.ToString();
                        strPartTargetPath += strIndex;
                        strPartTargetPath += ".SLDPRT";
                        PartDoc.Extension.SaveAs(strPartTargetPath, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, "", ref errors, ref warinings);

                        // 3 - 5: 打开已有零件(只有把文档打开后才可以将零件添加至装配下???)
                        swApp.OpenDoc(strPartTargetPath, (int)swDocumentTypes_e.swDocPART);

                        // 3 - 6: 将零件添加至装配节点
                        // 激活文档
                        string strDocName = "Assembley";
                        strDocName += (asmSum.ToString());
                        strDocName += ".SLDASM";

                        AssemblyDoc assemblyDoc = (AssemblyDoc)swApp.ActivateDoc3(strDocName, true, 0, errors);
                        swApp.ActivateDoc(strDocName);

                        // 装配零件
                        Component2 InsertedComponent = assemblyDoc.AddComponent5(strPartTargetPath, 0, "", false, "", 0, 0, 0);
                        InsertedComponent.Select(false);
                        assemblyDoc.UnfixComponent();

                        // 关闭零件文档
                        swApp.CloseDoc(strPartTargetPath);

                        // 关闭零件所在的装配文档(关闭前先保存)
                        ModelDoc2 assemblyDocTemp = (ModelDoc2)assemblyDoc;
                        assemblyDocTemp.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, errors, warinings);
                        strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "SolidWorks" + "\\" + "Assembley";
                        strAsmTargetPath += (asmSum.ToString());
                        strAsmTargetPath += ".SLDASM";
                        swApp.CloseDoc(strAsmTargetPath);

                        //3 - 7: 配合:
                        bool boolstatus = invAsmDocCube.Extension.SelectByID2("Plane1", "PLANE", 0, 0, 0, false, 0, null, 0);

                        boolstatus = invAsmDocCube.Extension.SelectByID2("Front Plane@clamp1-1@TempAssembly", "PLANE", 0, 0, 0, true, 0, null, 0);
                        int longstatus = 0;

                        // 重合
                        assemblyDoc.AddMate5(0, 0, false, 0, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus);
                        invAsmDocCube.EditRebuild3();
                        invAsmDocCube.ClearSelection();

                        // 距离配合
                        boolstatus = invAsmDocCube.Extension.SelectByID2("Plane2", "PLANE", 0, 0, 0, false, 0, null, 0);
                        boolstatus = invAsmDocCube.Extension.SelectByID2("Top Plane@clamp1-1@TempAssembly", "PLANE", 0, 0, 0, true, 0, null, 0);
                        assemblyDoc.AddMate5((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignALIGNED, true, 0.01, 0.01, 0.01, 0.01, 0.01, 0, 0, 0, false, false, 0, out longstatus);

                        // 3 - 8: 设置位置
                        var MathUtility = (MathUtility)swApp.GetMathUtility();

                        var swAxisVerX = (MathVector)MathUtility.CreateVector(new double[] { 1, 0, 0 });
                        var swAxisVerY = (MathVector)MathUtility.CreateVector(new double[] { 0, 1, 0 });
                        var swAxisVerZ = (MathVector)MathUtility.CreateVector(new double[] { 0, 0, 1 });

                        int x = k;
                        int y = j - 1 > 0 ? j - 1 : 0;
                        int z = i - 1 > 0 ? i - 1 : 0;

                        double gapTemp = (1 + cubeGap) * 0.1;
                        double gapValue = 5.0 + gapTemp;
                        var swAxisVerM = (MathVector)MathUtility.CreateVector(new double[] { 0.01 * x * gapValue, 0.01 * z * gapValue, 0.01 * y * gapValue });
                        var MathXform = (MathTransform)MathUtility.ComposeTransform(swAxisVerX, swAxisVerY, swAxisVerZ, swAxisVerM, 1);

                        InsertedComponent.Transform2 = MathXform;
                    }
                }
            }

            // step4: 保存装配
            AsmDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, errors, warinings);

            swApp.ExitApp();//退出
            swApp = null;
        }

        ///////////////////////////////Inventor///////////////////////////////////
        /** 
          * @brief        根据装配层级创建部件
          * @see          
          * @note          
         */
        private AssemblyDocument createAsmFirstLevel(Inventor.Application invApp)
        {
            AssemblyDocument asmDocTemp = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, strAsmTemplatePath, true);
            for (int i = 1; i<= asmLevel; ++i)
            {
                AssemblyDocument asmDocTemp1 = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, strAsmTemplatePath, true);

                ++asmSum;
                strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley";
                strAsmTargetPath += (asmSum.ToString());
                strAsmTargetPath += ".iam";
                asmDocTemp1.SaveAs(strAsmTargetPath, false);

                var ocurs = asmDocTemp.ComponentDefinition.Occurrences;
                var tranGeo = invApp.TransientGeometry;
                var transMatrix = tranGeo.CreateMatrix();
                ocurs.Add(strAsmTargetPath, transMatrix);
                asmDocTemp1.Close();

                if (i == asmLevel)
                {
                    sldAsmDocCube = asmDocTemp1;
                }
            }

            return asmDocTemp;
        }

        private AssemblyDocument createAsmSecondLevel(Inventor.Application invApp)
        {
            AssemblyDocument asmDocTemp = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, strAsmTemplatePath, true);
            for (int i = 1; i <= asmLevel; ++i)
            {
                AssemblyDocument asmDocTemp1 = createAsmFirstLevel(invApp);

                ++asmSum;
                strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley";
                strAsmTargetPath += asmSum.ToString();
                strAsmTargetPath += ".iam";
                asmDocTemp1.SaveAs(strAsmTargetPath, false);

                var ocurs = asmDocTemp.ComponentDefinition.Occurrences;
                var tranGeo = invApp.TransientGeometry;
                var transMatrix = tranGeo.CreateMatrix();
                ocurs.Add(strAsmTargetPath, transMatrix);
                asmDocTemp1.Close();
            }
            return asmDocTemp;
        }

        private AssemblyDocument createAsmThirdLevel(Inventor.Application invApp)
        {
            AssemblyDocument asmDocTemp = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, strAsmTemplatePath, true);
            for (int i = 1; i <= asmLevel; ++i)
            {
                AssemblyDocument asmDocTemp1 = createAsmSecondLevel(invApp);

                ++asmSum;
                strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley";
                strAsmTargetPath += asmSum.ToString();
                strAsmTargetPath += ".iam";
                asmDocTemp1.SaveAs(strAsmTargetPath, false);

                var ocurs = asmDocTemp.ComponentDefinition.Occurrences;
                var tranGeo = invApp.TransientGeometry;
                var transMatrix = tranGeo.CreateMatrix();
                ocurs.Add(strAsmTargetPath, transMatrix);
                asmDocTemp1.Close();
            }

            return asmDocTemp;
        }

        private AssemblyDocument createAsmFourthLevel(Inventor.Application invApp)
        {
            AssemblyDocument asmDocTemp = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, strAsmTemplatePath, true);
            for (int i = 1; i <= asmLevel; ++i)
            {
                AssemblyDocument asmDocTemp1 = createAsmThirdLevel(invApp);

                ++asmSum;
                strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley";
                strAsmTargetPath += asmSum.ToString();
                strAsmTargetPath += ".iam";
                asmDocTemp1.SaveAs(strAsmTargetPath, false);

                var ocurs = asmDocTemp.ComponentDefinition.Occurrences;
                var tranGeo = invApp.TransientGeometry;
                var transMatrix = tranGeo.CreateMatrix();
                ocurs.Add(strAsmTargetPath, transMatrix);
                asmDocTemp1.Close();
            }

            return asmDocTemp;
        }

        private AssemblyDocument createAsmFifththLevel(Inventor.Application invApp)
        {
            AssemblyDocument asmDocTemp = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, strAsmTemplatePath, true);
            for (int i = 1; i <= asmLevel; ++i)
            {
                AssemblyDocument asmDocTemp1 = createAsmThirdLevel(invApp);

                ++asmSum;
                strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley";
                strAsmTargetPath += asmSum.ToString();
                strAsmTargetPath += ".iam";
                asmDocTemp1.SaveAs(strAsmTargetPath, false);

                var ocurs = asmDocTemp.ComponentDefinition.Occurrences;
                var tranGeo = invApp.TransientGeometry;
                var transMatrix = tranGeo.CreateMatrix();
                ocurs.Add(strAsmTargetPath, transMatrix);
                asmDocTemp1.Close();
            }

            return asmDocTemp;
        }

        /** 
          * @brief        通过Inventor中的API创建立方体
          * @see          
          * @note          
         */
        private void createCubeFromInventor()
        {
            //string stIinventorApp = "Inventor.Application";

            // step1: 先判断是否存在Inventor程序，如果不存在则创建Inventor对象实例
            Inventor.Application invApp = null;
            try
            {
                invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");

            }
            catch
            {
                Type inventorApp = System.Type.GetTypeFromProgID("Inventor.Application");
                invApp = (Inventor.Application)Activator.CreateInstance(inventorApp);
                if (invApp == null)
                {
                    return;
                }
                //invApp.Visible = true;  // 显示Inventor
            }

            // step2: 创建装配
            // 生成独立的文档
            //string strAsmTemplatePath = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject);
            //AssemblyDocument asmDoc = (AssemblyDocument)invApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, strAsmTemplatePath, true);

            // 根据层级生产装配
            AssemblyDocument asmRootDoc = null;
            if (asmLevel == 1)
            {
                asmRootDoc = createAsmFirstLevel(invApp);
            }
            else if (asmLevel == 2)
            {
                asmRootDoc = createAsmSecondLevel(invApp);
            }
            else if (asmLevel == 3)
            {
                asmRootDoc = createAsmThirdLevel(invApp);

            }
            else if (asmLevel == 4)
            {
                asmRootDoc = createAsmFourthLevel(invApp);

            }
            else if (asmLevel == 5)
            {
                asmRootDoc = createAsmFifththLevel(invApp);
            }

            // step3: 创建立方体(根据指定参数生成立方体)
            int curSum = 0;
            for (int i = 1; i <= zNum; ++i)
            {
                for (int j = 1; j <= yNum; ++j)
                {
                    for (int k = 1; k <= xNum; ++k)
                    {
                        //3 - 1: 创建.ipt零件文件
                        string strPartTemplatePath = invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject);
                        PartDocument partDoc = (PartDocument)invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, strPartTemplatePath, true);

                        //3 - 2: 定义工作平面，1,2,3分别为yz,xz,xy
                        var workplane = partDoc.ComponentDefinition.WorkPlanes[3];

                        //3 - 3： 定义平面草图
                        PlanarSketch profSketch = partDoc.ComponentDefinition.Sketches.Add(workplane);
                        profSketch.Name = "Sketch";

                        //3 - 4: 平面草图，画矩形.通过两个角点确定(参数 两个点对象) - 返回值SketchEntitiesEnumerator，草图实体枚举器
                        // 2d点通过invapp.TransientGeometry.CreatePoint2d实现的
                        SketchEntitiesEnumerator olines = profSketch.SketchLines.AddAsTwoPointRectangle(invApp.TransientGeometry.CreatePoint2d(0, 0), invApp.TransientGeometry.CreatePoint2d(50, 50));

                        //3 - 5: 定义草图轮廓
                        Profile oProfile = profSketch.Profiles.AddForSolid();

                        //3 - 6: 创建拉伸定义
                        var oExtrudeDef = partDoc.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kJoinOperation);

                        //3 - 7: 设置拉伸的方向和尺寸
                        oExtrudeDef.SetDistanceExtent("50 cm ", PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);

                        //3 - 8: 创建拉伸体特征
                        var oExtrude = partDoc.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef);
                        oExtrude.Name = "Extrusion";

                        // 保存拉伸零件
                        string strPartTargetPath = System.IO.Path.Combine(strCurPath, "Extrusion");
                        strPartTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Extrusion";

                        ++curSum;
                        string strIndex = curSum.ToString();
                        strPartTargetPath += strIndex;
                        strPartTargetPath += ".ipt";
                        partDoc.SaveAs(strPartTargetPath, false);

                        // 3 - 9: 将零件文档添加至装配文档中
                        var ocurs = sldAsmDocCube.ComponentDefinition.Occurrences;
                        var tranGeo = invApp.TransientGeometry;

                        // 变换矩阵: 设置位置
                        var transMatrix = tranGeo.CreateMatrix();
                        int x = k - 1 > 0 ? k - 1 : 0;
                        int y = j - 1 > 0 ? j - 1 : 0;
                        int z = i - 1 > 0 ? i - 1 : 0;
                        int gapValue = 50 + cubeGap;
                        transMatrix.SetTranslation(tranGeo.CreateVector(x * gapValue, y * gapValue, z * gapValue));

                        // 装配零件
                        ocurs.Add(strPartTargetPath, transMatrix);

                        // 关闭零件
                        partDoc.Close();
                    }
                }
            }

            // step4: 保存装配
            // 包含实体的装配
            //var strAsmTargetPath = System.IO.Path.Combine(strCurPath, "AssemblyInentor.iam");
            //strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley_Last";
            //strAsmTargetPath += ".iam";
            //sldAsmDocCube.SaveAs(strAsmTargetPath, false);
            sldAsmDocCube.Save();
            sldAsmDocCube.Close();

            // 根装配
            strAsmTargetPath = strCurPath + "\\" + "ModeResult" + "\\" + "Inventor" + "\\" + "Assembley_Root";
            strAsmTargetPath += ".iam";
            asmRootDoc.SaveAs(strAsmTargetPath, false);
        }

        //---------------------------回调函数------------------------------
        /** 
          * @brief        切换类型的回调函数
          * @see          
          * @note          
         */
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 更新文档的路径
            initDocPath();
        }

        /** 
          * @brief        点击确定的回调函数
          * @see          
          * @note          
         */
        private void button1_Click(object sender, EventArgs e)
        {
            // 步骤1: 获取X、Y、Z方向的立方体的个数
            retrieveTextBoxValue();

            int sumValue = xNum * yNum * zNum;
            textBox4.ResetText();
            textBox4.AppendText(sumValue.ToString());

            // 步骤2: 根据不同方向立方体的个数构建立方体
            if (typeMode == 1)  // Inventor
            {
                createCubeFromInventor();
            }
            else // SolidWorks
            {
                createCubeFromSolidWorks();
            }

            this.Hide();
            this.Close();
        }

        /** 
          * @brief        点击取消的回调函数
          * @see          
          * @note          
         */
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            this.Close();
        }

        /** 
          * @brief        修改TextBox的回调
          * @see          
          * @note          
         */
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            retrieveTextBoxValue();
            if (xNum >= 1 && yNum >= 1 && zNum >= 1)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        /** 
          * @brief        修改TextBox的回调
          * @see          
          * @note          
         */
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            retrieveTextBoxValue();
            if (xNum >= 1 && yNum >= 1 && zNum >= 1)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        /** 
          * @brief        修改TextBox的回调
          * @see          
          * @note          
         */
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            retrieveTextBoxValue();
            if (xNum >= 1 && yNum >= 1 && zNum >= 1)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }
    }
}

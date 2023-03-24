using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;

namespace ValveCreator
{
    public partial class MainForm : Form
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Feature swFeature;
        private ModelDocExtension swModelDocExt;
        private FeatureManager swFeatMgr;
        private SelectionMgr swSelMgr;


        private double b;
        private double c;
        private double d;
        private double e;
        private double f;
        private double g;
        private double h;
        private double k;
        private double m;
        private double o;
        private double p;
        private double r;
        private double s;
        private int t;

        private double l;
        private double a;

        private double bMax;
        private double cMin;
        private double cMax;
        private double dMax;
        private double eMax;
        private double fMax;
        private double gMax;
        private double hMax;
        private double kMin;
        private double mMin;
        private double mMax;
        private double oMin;
        private double pMin;
        private double pMax;
        private double rMax;
        private double sMin;
        private double sMax;
        private int tMax;

        // конструктор за замовчуванням з ініціалізацією всіх параметричних розмірів моделі
        public MainForm()
        {
            InitializeComponent();
            b = 27.0;
            c = 49.0;
            d = 30.0;
            e = 38.0;
            f = 25.0;
            g = 15.0;
            h = 10.0;
            k = 58.0;
            m = 35.0;
            o = 76.0;
            p = 19.0;
            r = 12.0;
            s = 120.0;
            t = 4;
            l = k - m - g;
            a = k - e - h;

            updateLimits();
        }

        // метод загрузки вікна програми в якому задаються значення всіх параметрів в поля для їх введення
        private void MainForm_Load(object sender, EventArgs ee)
        {
            buttonSave.Enabled = false;
            textBoxTargetFolder.Text = Directory.GetCurrentDirectory();
            textBoxFileName.Text = "Valve";
            numericUpDownB.Value = (decimal)b;
            numericUpDownC.Value = (decimal)c;
            numericUpDownD.Value = (decimal)d;
            numericUpDownE.Value = (decimal)e;
            numericUpDownF.Value = (decimal)f;
            numericUpDownG.Value = (decimal)g;
            numericUpDownH.Value = (decimal)h;
            numericUpDownK.Value = (decimal)k;
            numericUpDownM.Value = (decimal)m;
            numericUpDownO.Value = (decimal)o;
            numericUpDownP.Value = (decimal)p;
            numericUpDownR.Value = (decimal)r;
            numericUpDownS.Value = (decimal)s;
            numericUpDownT.Value = (decimal)t;
            setBlackLable();
        }

        // метод оновлення підказок з межами можливих значень параметрів
        private void updateToolTips()
        {
            var myToolTip = new System.Windows.Forms.ToolTip();
            myToolTip.SetToolTip(numericUpDownB, "Значення В має бути меншим за " + bMax.ToString());
            myToolTip.SetToolTip(numericUpDownC, "Значення C має бути меншим за " + cMax.ToString() + " та більшим за " + cMin.ToString());
            myToolTip.SetToolTip(numericUpDownD, "Значення D має бути меншим за " + dMax.ToString());
            myToolTip.SetToolTip(numericUpDownE, "Значення E має бути меншим за " + eMax.ToString());
            myToolTip.SetToolTip(numericUpDownF, "Значення F має бути меншим за " + fMax.ToString());
            myToolTip.SetToolTip(numericUpDownG, "Значення G має бути меншим за " + gMax.ToString());
            myToolTip.SetToolTip(numericUpDownH, "Значення H має бути меншим за " + hMax.ToString());
            myToolTip.SetToolTip(numericUpDownK, "Значення K має бути більшим за " + kMin.ToString());
            myToolTip.SetToolTip(numericUpDownM, "Значення M має бути меншим за " + mMax.ToString() + " та більшим за " + mMin.ToString());
            myToolTip.SetToolTip(numericUpDownO, "Значення O має бути більшим за " + oMin.ToString());
            myToolTip.SetToolTip(numericUpDownP, "Значення P має бути не меншим за " + pMax.ToString() + " та більшим за " + pMin.ToString());
            myToolTip.SetToolTip(numericUpDownR, "Значення R має бути меншим за " + rMax.ToString());
            myToolTip.SetToolTip(numericUpDownS, "Значення S має бути меншим за " + sMax.ToString() + " та більшим за " + sMin.ToString());
            myToolTip.SetToolTip(numericUpDownT, "Значення T має бути меншим за " + tMax.ToString() + " та не меншим за 1.");
        }

        // метод встановлює колір тексту підписів полів для введення параметрів на чорний
        private void setBlackLable()
        {
            labelB.ForeColor = Color.Black;
            labelC.ForeColor = Color.Black;
            labelD.ForeColor = Color.Black;
            labelE.ForeColor = Color.Black;
            labelF.ForeColor = Color.Black;
            labelG.ForeColor = Color.Black;
            labelH.ForeColor = Color.Black;
            labelK.ForeColor = Color.Black;
            labelM.ForeColor = Color.Black;
            labelO.ForeColor = Color.Black;
            labelP.ForeColor = Color.Black;
            labelR.ForeColor = Color.Black;
            labelS.ForeColor = Color.Black;
            labelT.ForeColor = Color.Black;
        }

        // перевіряє всі параметри на коректність значень (чи лежать вони в своїх можливих межах)
        private void checkCorrectness()
        {
            bool corr = true;
            setBlackLable();
            if (b >= bMax)
            {
                labelB.ForeColor = Color.Red;
                corr = false;
            }
            if (c <= cMin || c >= cMax)
            {
                labelC.ForeColor = Color.Red;
                corr = false;
            }
            if (d >= dMax)
            {
                labelD.ForeColor = Color.Red;
                corr = false;
            }
            if (e >= eMax)
            {
                labelE.ForeColor = Color.Red;
                corr = false;
            }
            if (f >= fMax)
            {
                labelF.ForeColor = Color.Red;
                corr = false;
            }
            if (g >= gMax)
            {
                labelG.ForeColor = Color.Red;
                corr = false;
            }
            if (h >= hMax)
            {
                labelH.ForeColor = Color.Red;
                corr = false;
            }
            if (k <= kMin)
            {
                labelK.ForeColor = Color.Red;
                corr = false;
            }
            if (m <= mMin || m >= mMax)
            {
                labelM.ForeColor = Color.Red;
                corr = false;
            }
            if (o <= oMin)
            {
                labelO.ForeColor = Color.Red;
                corr = false;
            }
            if (p <= pMin || p > pMax)
            {
                labelP.ForeColor = Color.Red;
                corr = false;
            }
            if (r >= rMax)
            {
                labelR.ForeColor = Color.Red;
                corr = false;
            }
            if (s <= sMin || s >= sMax)
            {
                labelS.ForeColor = Color.Red;
                corr = false;
            }
            if (t >= tMax)
            {
                labelT.ForeColor = Color.Red;
                corr = false;
            }
            buttonCreatePart.Enabled = corr;
        }

        // метод для перерахунку меж можливих значень параметрів
        private void updateLimits()
        {
            bMax = c - Math.Tan((s - 90.0) * Math.PI / 180.0) * h;
            cMin = b + Math.Tan((s - 90.0) * Math.PI / 180.0) * h;
            cMax = o;
            dMax = c;
            eMax = k - h - 2.0;
            fMax = m - 1.0;
            gMax = a + h;
            hMax = k - e - 1.0;
            kMin = e + h + 1.0;
            mMax = k - g - 1.0;
            mMin = f;
            oMin = c + 5.0;
            pMax = f - r / 2;
            pMin = r / 2;
            rMax = p + 1.0;
            sMin = Math.Atan(-b / (h * 2)) / Math.PI * 180.0 + 90.0 + 3;
            sMax = Math.Atan((c - b) / (h * 2)) / Math.PI * 180.0 + 90.0 - 3;
            tMax = (int)(Math.PI / Math.Asin(r / c));

            updateToolTips();
            checkCorrectness();
        }

        // метод для створення моделі клапана
        private void drawPart()
        {
            // переводимо всі діаметри в радіуси, та ділимо всі значення на 1000, щоб перевести в мм
            b /= 2.0; c /= 2.0; d /= 2.0; o /= 2.0; r /= 2.0;
            a = k - e - h;
            l = k - m - g;
            a /= 1000.0; b /= 1000.0; c /= 1000.0; d /= 1000.0; e /= 1000.0; f /= 1000.0; g /= 1000.0;
            h /= 1000.0; k /= 1000.0; m /= 1000.0; l /= 1000.0; o /= 1000.0; p /= 1000.0; r /= 1000.0;

            // змінюємо імена всіх стандартних площин
            swFeature = swModel.FeatureByPositionReverse(3);
            swFeature.Name = "Front";

            swFeature = swModel.FeatureByPositionReverse(2);
            swFeature.Name = "Top";

            swFeature = swModel.FeatureByPositionReverse(1);
            swFeature.Name = "Right";

            // створюємо ось
            swModelDocExt.SelectByID2("Front", "PLANE", 0, 0, 0, true, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Top", "PLANE", 0, 0, 0, true, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModel.InsertAxis2(true);

            // даємо ім'я осі
            swFeature = swModel.FeatureByPositionReverse(0);
            swFeature.Name = "Axis1";

            // створюємо ескіз на площині Front
            swModel.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            swModel.SketchManager.AutoSolve = true;

            // проводимо всі потрібні лінії
            swModel.CreateLine2(0.0, 0.0, 0.0, a, 0.0, 0.0);
            swModel.CreateLine2(a, 0.0, 0.0, a, b, 0.0);
            swModel.CreateLine2(a, b, 0.0, a + h, b + Math.Tan((s - 90.0) * Math.PI / 180.0) * h, 0.0);
            swModel.CreateLine2(a + h, b + Math.Tan((s - 90.0) * Math.PI / 180.0) * h, 0.0, a + h, c, 0.0);
            swModel.CreateLine2(a + h, c, 0.0, a + h - g, c, 0.0);
            swModel.CreateLine2(a + h - g, c, 0.0, a + h - g, o, 0.0);
            swModel.CreateLine2(a + h - g, o, 0.0, a + h - g - l, o, 0.0);
            swModel.CreateLine2(a + h - g - l, o, 0.0, a + h - g - l, c + m - f, 0.0);
            swModel.CreateLine2(a + h - g - l, c + m - f, 0.0, 0.0 - e + f, c, 0.0);
            swModel.CreateLine2(0.0 - e + f, c, 0.0, 0.0 - e, c, 0.0);
            swModel.CreateLine2(0.0 - e, c, 0.0, 0.0 - e, d, 0.0);
            swModel.CreateLine2(0.0 - e, d, 0.0, 0.0, d, 0.0);
            swModel.CreateLine2(0.0, d, 0.0, 0.0, 0.0, 0.0);

            swModel.AutoSolveToggle();
            swModel.InsertSketch2(false);
            // блок визначення розмірів ексізу (автоматичне визначення ескізу)
            {
                bool bValue = false;
                SketchManager swSketchManager = default(SketchManager);
                ModelDocExtension swModelExtension = default(ModelDocExtension);
                int lStatus = 0;
                int lMarkHorizontal = 0;
                int lMarkVertical = 0;
                SelectionMgr swSelectionManager = default(SelectionMgr);

                swModel = (ModelDoc2)swApp.ActiveDoc;
                swModelExtension = swModel.Extension;
                swSketchManager = swModel.SketchManager;
                swSelectionManager = (SelectionMgr)swModel.SelectionManager;

                swModel.ClearSelection2(true);

                // These are the marks expected for the dimension datums
                lMarkHorizontal = 2;
                lMarkVertical = 4;

                swFeature = (Feature)swModel.FirstFeature();

                while (((swFeature != null)))
                {
                    if ((swFeature.GetTypeName() == "ProfileFeature"))
                    {
                        break;
                    }
                    swFeature = (Feature)swFeature.GetNextFeature();
                }

                if (((swFeature != null)))
                {
                    bValue = swFeature.Select2(false, 0);
                    swSketchManager.InsertSketch(false);

                    // OR together the marks for the vertical and horizontal datums;
                    // You cannot select the same point twice with different marks
                    bValue = swModelExtension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, false, lMarkHorizontal | lMarkVertical, null, 0);
                    lStatus = swSketchManager.FullyDefineSketch(true, true, (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Vertical | (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Horizontal, true, 1, null, 1, null, 1, 1);

                    swSketchManager.InsertSketch(true);
                }
            }

            // даємо назву ескізу
            swFeature = swModel.FeatureByPositionReverse(0);
            swFeature.Name = "Sketch1";

            // обираємо есзік і ось для створення бобишки
            swModelDocExt.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swModelDocExt.SelectByID2("Axis1", "AXIS", 0, 0, 0, true, 16, null, (int)swSelectOption_e.swSelectOptionDefault);
            swFeatMgr = swModel.FeatureManager;

            // створюємо бобишку обертом
            swFeatMgr.FeatureRevolve2(true, true, false, false, false, false, 0, 0, 2.0 * Math.PI/*6.28318530718*/, 0, false,
                false, 0.01, 0.01, 0, 0, 0, true, true, true);

            // даємо назву бубишкі
            swFeature = swModel.FeatureByPositionReverse(0);
            swFeature.Name = "Revolve1";

            // на площині Front створюємо ескіз під виріз
            swModel.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, false, 0, null, 0);

            swModel.InsertSketch2(true);
            swModel.CreateCircleByRadius2(0.0 - e + p, 0.0, m, r);
            swModel.InsertSketch2(false);

            // даємо назву ескізу під виріз
            swFeature = swModel.FeatureByPositionReverse(0);
            swFeature.Name = "Sketch2";

            // створення вирізу
            swModelDocExt.SelectByID2("Sketch2", "SKETCH", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
            swFeatMgr = swModel.FeatureManager;
            swFeature = swModel.FeatureManager.FeatureCut3(true, false, true, (int)swEndConditions_e.swEndCondThroughAllBoth,
               (int)swEndConditions_e.swEndCondThroughAllBoth, 0, 0, false, false, false, false, 0, 0, false, false, false, false,
                false, true, true, false, false, false, (int)swEndConditions_e.swEndCondMidPlane, 0, true);

            swFeature = swModel.FeatureByPositionReverse(0);
            swFeature.Name = "Cut1";

            swModel.ClearSelection2(true);

            // створення масиву вирізів
            CircularPatternFeatureData swCircularPatternFeatureData = default(CircularPatternFeatureData);
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swFeatMgr = (FeatureManager)swModel.FeatureManager;

            swModelDocExt.SelectByID2("Cut1", "BODYFEATURE", 0, 0, 0, false, 4, null, 0);
            swModelDocExt.SelectByID2("Axis1", "AXIS", 0, 0, 0, true, 1, null, 0);
            swCircularPatternFeatureData = (CircularPatternFeatureData)swFeatMgr.CreateDefinition((int)swFeatureNameID_e.swFmCirPattern);
            swCircularPatternFeatureData.Direction2 = false;
            swCircularPatternFeatureData.EqualSpacing = true;
            swCircularPatternFeatureData.GeometryPattern = false;
            swCircularPatternFeatureData.ReverseDirection = false;
            swCircularPatternFeatureData.Spacing = 2 * Math.PI;
            swCircularPatternFeatureData.TotalInstances = t;
            swCircularPatternFeatureData.VarySketch = false;
            swFeature = (Feature)swFeatMgr.CreateFeature(swCircularPatternFeatureData);

            // перебудовуємо модель
            swModel.ForceRebuild3(true);

            // повертаємо значення параметрів назад
            b *= 2.0; c *= 2.0; d *= 2.0; o *= 2.0; r *= 2.0;
            a *= 1000.0; b *= 1000.0; c *= 1000.0; d *= 1000.0; e *= 1000.0; f *= 1000.0; g *= 1000.0;
            h *= 1000.0; k *= 1000.0; m *= 1000.0; l *= 1000.0; o *= 1000.0; p *= 1000.0; r *= 1000.0;
        }

        // метод реакції програми на натиск кнопки "Створення деталі"
        private void buttonCreatePart_Click(object sender, EventArgs e)
        {
            // Якщо солід не відкритий, то відкриваємо його, інакше використвуємо вже відкритий солід
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch (Exception)
            {
                swApp = new SldWorks();
            }
            swApp.Visible = true;

            // створюємо нову деталь
            string defaultPartTemplate;
            defaultPartTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExt = swModel.Extension;
            swSelMgr = (SelectionMgr)swModel.SelectionManager;

            drawPart();

            buttonSave.Enabled = true;
        }

        // метод реакції програми на натиск кнопки "..."
        private void buttonSelPath_Click(object sender, EventArgs e)
        {
            // виклик стандратного вікна вибору шляху
            FolderBrowserDialog target = new FolderBrowserDialog();
            target.ShowDialog();
            textBoxTargetFolder.Text = target.SelectedPath;
        }

        // метод реакції програми на натиск кнопки "зберігти"
        private void buttonSave_Click(object sender, EventArgs e)
        {
            // перевіряємо задане ім'я файлу на правильність
            string fileName = textBoxFileName.Text;
            for (int i = 0; i < fileName.Length; i++)
            {
                if (fileName[i] == '\\' || fileName[i] == '/' || fileName[i] == ':'
                    || fileName[i] == '*' || fileName[i] == '\"' || fileName[i] == '?'
                    || fileName[i] == '<' || fileName[i] == '>' || fileName[i] == '|')
                {
                    MessageBox.Show("Ім\'я файлу не має містити наступні символи:\n\\ / : * ? \" < > |",
                        "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            //зберігаємо файл по можливості
            string path = textBoxTargetFolder.Text + '\\' + fileName + ".SLDPRT";
            try
            {
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swModel.SaveAs(path);
            }
            catch (Exception)
            {
                MessageBox.Show("Невдалося зберегти файл",
                        "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // методи з закінченням на ValueChanged відповідають на зміни параметрів моделі та перевіряють коректність вводу даних
        private void numericUpDownB_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownB.Value >= bMax)
            {
                MessageBox.Show("Значення В має бути меншим за " + bMax.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownB.Value = (decimal)b;
                return;
            }
            b = (double)numericUpDownB.Value;
            updateLimits();
        }

        private void numericUpDownC_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownC.Value <= cMin || (double)numericUpDownC.Value >= cMax)
            {
                string errMsg = "Значення C має бути меншим за " + cMax.ToString() +
                    " та більшим за " + cMin.ToString();
                MessageBox.Show(errMsg, "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownC.Value = (decimal)c;
                return;
            }
            c = (double)numericUpDownC.Value;
            updateLimits();
        }

        private void numericUpDownD_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownD.Value >= dMax)
            {
                MessageBox.Show("Значення D має бути меншим за " + dMax.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownD.Value = (decimal)d;
                return;
            }
            d = (double)numericUpDownD.Value;
            updateLimits();
        }

        private void numericUpDownE_ValueChanged(object sender, EventArgs ee)
        {
            if ((double)numericUpDownE.Value >= eMax)
            {
                MessageBox.Show("Значення E має бути меншим за " + eMax.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownE.Value = (decimal)e;
                return;
            }
            e = (double)numericUpDownE.Value;
            a = k - e - h;
            updateLimits();
        }

        private void numericUpDownF_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownF.Value >= fMax)
            {
                MessageBox.Show("Значення F має бути меншим за " + fMax.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownF.Value = (decimal)f;
                return;
            }
            f = (double)numericUpDownF.Value;
            updateLimits();
        }

        private void numericUpDownG_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownG.Value >= gMax)
            {
                MessageBox.Show("Значення G має бути меншим за " + gMax.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownG.Value = (decimal)g;
                return;
            }
            g = (double)numericUpDownG.Value;
            l = k - m - g;
            updateLimits();
        }

        private void numericUpDownH_ValueChanged(object sender, EventArgs ee)
        {
            if ((double)numericUpDownH.Value >= hMax)
            {
                MessageBox.Show("Значення H має бути меншим за " + hMax.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownH.Value = (decimal)h;
                return;
            }
            h = (double)numericUpDownH.Value;
            a = k - e - h;
            updateLimits();
        }

        private void numericUpDownK_ValueChanged(object sender, EventArgs ee)
        {
            if ((double)numericUpDownK.Value <= kMin)
            {
                MessageBox.Show("Значення K має бути більшим за " + kMin.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownK.Value = (decimal)k;
                return;
            }
            k = (double)numericUpDownK.Value;
            a = k - e - h;
            l = k - m - g;
            updateLimits();
        }

        private void numericUpDownM_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownM.Value <= mMin || (double)numericUpDownM.Value >= mMax)
            {
                string errMsg = "Значення M має бути меншим за " + mMax.ToString() +
                    " та більшим за " + mMin.ToString();
                MessageBox.Show(errMsg, "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownM.Value = (decimal)m;
                return;
            }
            m = (double)numericUpDownM.Value;
            l = k - m - g;
            updateLimits();
        }

        private void numericUpDownO_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownO.Value <= oMin)
            {
                MessageBox.Show("Значення O має бути більшим за " + oMin.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownO.Value = (decimal)o;
                return;
            }
            o = (double)numericUpDownO.Value;
            updateLimits();
        }

        private void numericUpDownP_ValueChanged(object sender, EventArgs ee)
        {
            if ((double)numericUpDownP.Value <= pMin || (double)numericUpDownP.Value > pMax)
            {
                string errMsg = "Значення P має бути не меншим за " + pMax.ToString() +
                    " та більшим за " + pMin.ToString();
                MessageBox.Show(errMsg, "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownP.Value = (decimal)p;
                return;
            }
            p = (double)numericUpDownP.Value;
            updateLimits();
        }

        private void numericUpDownR_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownR.Value >= rMax)
            {
                MessageBox.Show("Значення R має бути меншим за " + rMax.ToString(), "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownR.Value = (decimal)r;
                return;
            }
            r = (double)numericUpDownR.Value;
            updateLimits();
        }

        private void numericUpDownS_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDownS.Value <= sMin || (double)numericUpDownS.Value >= sMax)
            {
                string errMsg = "Значення S має бути меншим за " + sMax.ToString() +
                    " та більшим за " + sMin.ToString();
                MessageBox.Show(errMsg, "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownS.Value = (decimal)s;
                return;
            }
            s = (double)numericUpDownS.Value;
            updateLimits();
        }

        private void numericUpDownT_ValueChanged(object sender, EventArgs e)
        {
            if ((int)numericUpDownT.Value < 1 || (int)numericUpDownT.Value > tMax)
            {
                MessageBox.Show("Значення T має бути меншим за " + tMax.ToString() + " та не меншим за 1.", "Невірне значення", MessageBoxButtons.OK, MessageBoxIcon.Error);
                numericUpDownT.Value = (decimal)t;
                return;
            }
            t = (int)numericUpDownT.Value;
            updateLimits();
        }

        // Рекція на сняття курсору з поля для вводу
        private void numericUpDown_Leave(object sender, EventArgs e)
        {
            checkCorrectness();
        }

        // методи з закінченням на Enter відповідають за зміну кольору підписів параметрів на помаранчевий
        private void numericUpDownB_Enter(object sender, EventArgs e)
        {
            labelC.ForeColor = Color.DarkOrange;
            labelS.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownC_Enter(object sender, EventArgs e)
        {
            labelB.ForeColor = Color.DarkOrange;
            labelD.ForeColor = Color.DarkOrange;
            labelO.ForeColor = Color.DarkOrange;
            labelS.ForeColor = Color.DarkOrange;
            labelT.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownD_Enter(object sender, EventArgs e)
        {
        }

        private void numericUpDownE_Enter(object sender, EventArgs e)
        {
            labelH.ForeColor = Color.DarkOrange;
            labelK.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownF_Enter(object sender, EventArgs e)
        {
            labelM.ForeColor = Color.DarkOrange;
            labelP.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownG_Enter(object sender, EventArgs e)
        {
            labelM.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownH_Enter(object sender, EventArgs e)
        {
            labelB.ForeColor = Color.DarkOrange;
            labelC.ForeColor = Color.DarkOrange;
            labelE.ForeColor = Color.DarkOrange;
            labelG.ForeColor = Color.DarkOrange;
            labelK.ForeColor = Color.DarkOrange;
            labelS.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownK_Enter(object sender, EventArgs e)
        {
            labelE.ForeColor = Color.DarkOrange;
            labelH.ForeColor = Color.DarkOrange;
            labelM.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownM_Enter(object sender, EventArgs e)
        {
            labelF.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownO_Enter(object sender, EventArgs e)
        {
            labelC.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownP_Enter(object sender, EventArgs e)
        {
            labelR.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownR_Enter(object sender, EventArgs e)
        {
            labelP.ForeColor = Color.DarkOrange;
            labelT.ForeColor = Color.DarkOrange;
        }

        private void numericUpDownS_Enter(object sender, EventArgs e)
        {
        }

        private void numericUpDownT_Enter(object sender, EventArgs e)
        {
        }
    }
}
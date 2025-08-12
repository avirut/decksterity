using decksterity.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    [ComVisible(true)]
    public class DecksterityRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DecksterityRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("decksterity.DecksterityRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI; 
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "Blank": return Resources.Blank;
                case "HarveyBall0": return Resources.HarveyBall0;
                case "HarveyBall1": return Resources.HarveyBall1;
                case "HarveyBall2": return Resources.HarveyBall2;
                case "HarveyBall3": return Resources.HarveyBall3;
                case "HarveyBall4": return Resources.HarveyBall4;
                case "HarveyBallCustom": return Resources.HarveyBallCustom;
                case "ArrowNW": return Resources.ArrowNW;
                case "ArrowN": return Resources.ArrowN;
                case "ArrowNE": return Resources.ArrowNE;
                case "ArrowW": return Resources.ArrowW;
                case "ArrowE": return Resources.ArrowE;
                case "ArrowSW": return Resources.ArrowSW;
                case "ArrowS": return Resources.ArrowS;
                case "ArrowSE": return Resources.ArrowSE;
                case "IconCheck": return Resources.IconCheck;
                case "IconCross": return Resources.IconCross;
                case "IconQuestion": return Resources.IconQuestion;
                case "IconPlus": return Resources.IconPlus;
                case "IconMinus": return Resources.IconMinus;
                case "IconEllipsis": return Resources.IconEllipsis;
                case "StoplightRed": return Resources.StoplightRed;
                case "StoplightAmber": return Resources.StoplightAmber;
                case "StoplightGreen": return Resources.StoplightGreen;
                default: return Resources.Blank;
            }
        }

        // Ribbon onAction callback headers generated from XML
        public void HarveyBall0(Office.IRibbonControl control) { ElementHelper.InsertElement("⭘"); }
        public void HarveyBall1(Office.IRibbonControl control) { ElementHelper.InsertElement("◔"); }
        public void HarveyBall2(Office.IRibbonControl control) { ElementHelper.InsertElement("◑"); }
        public void HarveyBall3(Office.IRibbonControl control) { ElementHelper.InsertElement("◕"); }
        public void HarveyBall4(Office.IRibbonControl control) { ElementHelper.InsertElement("●"); }
        public void StoplightRed(Office.IRibbonControl control) { ElementHelper.InsertElement("●", 0xab0e04); }
        public void StoplightAmber(Office.IRibbonControl control) { ElementHelper.InsertElement("●", 0xe2ad00); }
        public void StoplightGreen(Office.IRibbonControl control) { ElementHelper.InsertElement("●", 0x007748); }
        public void IconCheck(Office.IRibbonControl control) { ElementHelper.InsertElement("✔"); }
        public void IconPlus(Office.IRibbonControl control) { ElementHelper.InsertElement("➕"); }
        public void IconQuestion(Office.IRibbonControl control) { ElementHelper.InsertElement("❓"); }
        public void IconCross(Office.IRibbonControl control) { ElementHelper.InsertElement("✘"); }
        public void IconMinus(Office.IRibbonControl control) { ElementHelper.InsertElement("➖"); }
        public void IconEllipsis(Office.IRibbonControl control) { ElementHelper.InsertElement("…"); }
        public void ArrowNW(Office.IRibbonControl control) { ElementHelper.InsertElement("🡼"); }
        public void ArrowW(Office.IRibbonControl control) { ElementHelper.InsertElement("🡸"); }
        public void ArrowSW(Office.IRibbonControl control) { ElementHelper.InsertElement("🡿"); }
        public void ArrowN(Office.IRibbonControl control) { ElementHelper.InsertElement("🡹"); }
        public void ArrowS(Office.IRibbonControl control) { ElementHelper.InsertElement("🡻"); }
        public void ArrowNE(Office.IRibbonControl control) { ElementHelper.InsertElement("🡽"); }
        public void ArrowE(Office.IRibbonControl control) { ElementHelper.InsertElement("🡺"); }
        public void ArrowSE(Office.IRibbonControl control) { ElementHelper.InsertElement("🡾"); }
        public void AlignLeft(Office.IRibbonControl control) { AlignmentHelper.AlignLeft(); }
        public void AlignBottom(Office.IRibbonControl control) { AlignmentHelper.AlignBottom(); }
        public void AlignCenter(Office.IRibbonControl control) { AlignmentHelper.AlignCenter(); }
        public void AlignMiddle(Office.IRibbonControl control) { AlignmentHelper.AlignMiddle(); }
        public void AlignRight(Office.IRibbonControl control) { AlignmentHelper.AlignRight(); }
        public void AlignTop(Office.IRibbonControl control) { AlignmentHelper.AlignTop(); }
        public void ResizeAndSpaceEvenHorizontal(Office.IRibbonControl control) { AlignmentHelper.ResizeAndSpaceEvenly("evenhorizontal"); }
        public void ResizeAndSpaceEvenHorizontalPreserveFirst(Office.IRibbonControl control) { AlignmentHelper.ResizeAndSpaceEvenly("evenhorizontalpreservefirst"); }
        public void ResizeAndSpaceEvenHorizontalPreserveLast(Office.IRibbonControl control) { AlignmentHelper.ResizeAndSpaceEvenly("evenhorizontalpreservelast"); }
        public void ResizeAndSpaceEvenVertical(Office.IRibbonControl control) { AlignmentHelper.ResizeAndSpaceEvenly("evenvertical"); }
        public void ResizeAndSpaceEvenVerticalPreserveFirst(Office.IRibbonControl control) { AlignmentHelper.ResizeAndSpaceEvenly("evenverticalpreservefirst"); }
        public void ResizeAndSpaceEvenVerticalPreserveLast(Office.IRibbonControl control) { AlignmentHelper.ResizeAndSpaceEvenly("evenverticalpreservelast"); }
        public void DistributeHorizontally(Office.IRibbonControl control) { AlignmentHelper.DistributeHorizontally(); }
        public void DistributeVertically(Office.IRibbonControl control) { AlignmentHelper.DistributeVertically(); }
        public void SameHeight(Office.IRibbonControl control) { AlignmentHelper.SameHeight(); }
        public void SameWidth(Office.IRibbonControl control) { AlignmentHelper.SameWidth(); }
        public void PrimaryAlignLeft(Office.IRibbonControl control) { AlignmentHelper.PrimaryAlignLeft(); }
        public void PrimaryAlignBottom(Office.IRibbonControl control) { AlignmentHelper.PrimaryAlignBottom(); }
        public void PrimaryAlignCenter(Office.IRibbonControl control) { AlignmentHelper.PrimaryAlignCenter(); }
        public void PrimaryAlignMiddle(Office.IRibbonControl control) { AlignmentHelper.PrimaryAlignMiddle(); }
        public void PrimaryAlignRight(Office.IRibbonControl control) { AlignmentHelper.PrimaryAlignRight(); }
        public void PrimaryAlignTop(Office.IRibbonControl control) { AlignmentHelper.PrimaryAlignTop(); }
        public void ObjectsSwapPositionCentered(Office.IRibbonControl control) { AlignmentHelper.ObjectsSwapPositionCentered(); }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }


        #endregion
    }
}

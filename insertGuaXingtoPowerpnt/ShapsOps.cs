using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows.Forms;
using PowerPnt = Microsoft.Office.Interop.PowerPoint;
using WinWord = Microsoft.Office.Interop.Word;


namespace insertGuaXingtoPowerpnt
{
    class ShapsOps<T> : WinWord.Shape, PowerPnt.Shape
    {

        string name;
        WinWord.Shape docShape = null;
        PowerPnt.Shape pptShape = null;
        WinWord.ShapeRange docShapeRng = null;
        PowerPnt.ShapeRange pptShapeRng = null;
        object shapsOps;
        
        
        public ShapsOps(T shape, string typeName)
        {
            shapsOps = shape;
            Name = typeName;
            switch (typeName)
            {
                case "PowerPnt.Shape":
                    pptShape = shape as PowerPnt.Shape;
                    //shapsOps = pptShape;
                    break;
                case "WinWord.Shape":
                    docShape = shape as WinWord.Shape;
                    //shapsOps = docShape;
                    break;
                case "PowerPnt.ShapeRange":
                    pptShapeRng = (PowerPnt.ShapeRange)shape;
                    //shapsOps = pptShapeRng;
                    break;
                case "WinWord.ShapeRange":
                    docShapeRng = (WinWord.ShapeRange)shape;
                    //shapsOps = docShapeRng;
                    break;
                default:
                    break;
            }
        }
        public Type getType { get => shapsOps.GetType(); }
        internal virtual bool isShapeRnginMeShape(PowerPnt.ShapeRange pptSpRng)
        {
            if (pptSpRng == null)
            {
                MessageBox.Show("所給的shape有誤");
                return false;
            }//pptShape is Me ,outside the pptSp
            if (pptShape.Left <= pptSpRng.Left && pptShape.Top <= pptSpRng.Top &&
                pptShape.Left+pptShape.Width >= pptSpRng.Left+pptSpRng.Width &&
                pptShape.Top+pptShape.Height >= pptSpRng.Top+pptSpRng.Height)
            {
                return true;
            }
            return false;
        }

        internal virtual bool isShapeContainsMeShapeRng(PowerPnt.Shape pptSp)
        {
            if (pptSp == null)
            {
                MessageBox.Show("所給的shape有誤");
                return false;
            }//pptShape is Me ,inside the pptSp

            foreach (PowerPnt.Shape item in pptShapeRng)
            {
                if (item.Left >= pptSp.Left && item.Top >= pptSp.Top &&
                    item.Left + item.Width <= pptSp.Left + pptSp.Width &&
                    item.Top + item.Height <= pptSp.Top + pptSp.Height)
                {
                    return true;
                }                
            }
            return false;
        }
        internal virtual bool isShapeContainsMeShape(PowerPnt.Shape pptSp)
        {
            if (pptSp == null)
            {
                MessageBox.Show("所給的shape有誤");
                return false;
            }//pptShape is Me ,inside the pptSp
            if (pptShape.Left >= pptSp.Left && pptShape.Top >= pptSp.Top &&
                pptShape.Left + pptShape.Width <= pptSp.Left + pptSp.Width &&
                pptShape.Top + pptShape.Height <= pptSp.Top + pptSp.Height)
            {
                return true;
            }
            return false;
        }
        public void Apply()
        {
            pptShape.Apply();
        }

        public void Delete()
        {
            pptShape.Delete();
        }

        public void Flip(MsoFlipCmd FlipCmd)
        {
            pptShape.Flip(FlipCmd);
        }

        public void IncrementLeft(float Increment)
        {
            pptShape.IncrementLeft(Increment);
        }

        public void IncrementRotation(float Increment)
        {
            pptShape.IncrementRotation(Increment);
        }

        public void IncrementTop(float Increment)
        {
            pptShape.IncrementTop(Increment);
        }

        public void PickUp()
        {
            pptShape.PickUp();
        }

        public void RerouteConnections()
        {
            pptShape.RerouteConnections();
        }

        public void ScaleHeight(float Factor, MsoTriState RelativeToOriginalSize, MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft)
        {
            pptShape.ScaleHeight(Factor, RelativeToOriginalSize, fScale);
        }

        public void ScaleWidth(float Factor, MsoTriState RelativeToOriginalSize, MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft)
        {
            pptShape.ScaleWidth(Factor, RelativeToOriginalSize, fScale);
        }

        public void SetShapesDefaultProperties()
        {
            pptShape.SetShapesDefaultProperties();
        }

        public PowerPnt.ShapeRange Ungroup()
        {
            return pptShape.Ungroup();
        }

        public void ZOrder(MsoZOrderCmd ZOrderCmd)
        {
            pptShape.ZOrder(ZOrderCmd);
        }

        public void Cut()
        {
            pptShape.Cut();
        }

        public void Copy()
        {
            pptShape.Copy();
        }

        public void Select(MsoTriState Replace = MsoTriState.msoTrue)
        {
            pptShape.Select(Replace);
        }

        public PowerPnt.ShapeRange Duplicate()
        {
            return pptShape.Duplicate();
        }

        public void Export(string PathName, PpShapeFormat Filter, int ScaleWidth = 0, int ScaleHeight = 0, PpExportMode ExportMode = PpExportMode.ppRelativeToSlide)
        {
            pptShape.Export(PathName, Filter, ScaleWidth, ScaleHeight, ExportMode);
        }

        public void CanvasCropLeft(float Increment)
        {
            pptShape.CanvasCropLeft(Increment);
        }

        public void CanvasCropTop(float Increment)
        {
            pptShape.CanvasCropTop(Increment);
        }

        public void CanvasCropRight(float Increment)
        {
            pptShape.CanvasCropRight(Increment);
        }

        public void CanvasCropBottom(float Increment)
        {
            pptShape.CanvasCropBottom(Increment);
        }

        public void ConvertTextToSmartArt(SmartArtLayout Layout)
        {
            pptShape.ConvertTextToSmartArt(Layout);
        }

        public void PickupAnimation()
        {
            pptShape.PickupAnimation();
        }

        public void ApplyAnimation()
        {
            pptShape.ApplyAnimation();
        }

        public void UpgradeMedia()
        {
            pptShape.UpgradeMedia();
        }

        public dynamic Application => pptShape.Application;

        public int Creator => pptShape.Creator;

        public dynamic Parent => pptShape.Parent;

        public PowerPnt.Adjustments Adjustments => pptShape.Adjustments;

        public MsoAutoShapeType AutoShapeType { get => pptShape.AutoShapeType; set => pptShape.AutoShapeType = value; }
        public MsoBlackWhiteMode BlackWhiteMode { get => pptShape.BlackWhiteMode; set => pptShape.BlackWhiteMode = value; }

        public PowerPnt.CalloutFormat Callout => pptShape.Callout;

        public int ConnectionSiteCount => pptShape.ConnectionSiteCount;

        public MsoTriState Connector => pptShape.Connector;

        public PowerPnt.ConnectorFormat ConnectorFormat => pptShape.ConnectorFormat;

        public PowerPnt.FillFormat Fill => pptShape.Fill;

        public PowerPnt.GroupShapes GroupItems => pptShape.GroupItems;

        public float Height { get => pptShape.Height; set => pptShape.Height = value; }

        public MsoTriState HorizontalFlip => pptShape.HorizontalFlip;

        public float Left { get => pptShape.Left; set => pptShape.Left = value; }

        public PowerPnt.LineFormat Line => pptShape.Line;

        public MsoTriState LockAspectRatio { get => pptShape.LockAspectRatio; set => pptShape.LockAspectRatio = value; }
        //public string Name { get => pptShape.Name; set => pptShape.Name = value; }
        public string Name { get => name; set => name = value; }

        public PowerPnt.ShapeNodes Nodes => pptShape.Nodes;

        public float Rotation { get => pptShape.Rotation; set => pptShape.Rotation = value; }

        public PowerPnt.PictureFormat PictureFormat => pptShape.PictureFormat;

        public PowerPnt.ShadowFormat Shadow => pptShape.Shadow;

        public PowerPnt.TextEffectFormat TextEffect => pptShape.TextEffect;

        public PowerPnt.TextFrame TextFrame => pptShape.TextFrame;

        public PowerPnt.ThreeDFormat ThreeD => pptShape.ThreeD;

        public float Top { get => pptShape.Top; set => pptShape.Top = value; }

        public MsoShapeType Type => pptShape.Type;

        public MsoTriState VerticalFlip => pptShape.VerticalFlip;

        public dynamic Vertices => pptShape.Vertices;

        public MsoTriState Visible { get => pptShape.Visible; set => pptShape.Visible = value; }
        public float Width { get => pptShape.Width; set => pptShape.Width = value; }

        public int ZOrderPosition => pptShape.ZOrderPosition;

        public OLEFormat OLEFormat => pptShape.OLEFormat;

        public LinkFormat LinkFormat => pptShape.LinkFormat;

        public PlaceholderFormat PlaceholderFormat => pptShape.PlaceholderFormat;

        public AnimationSettings AnimationSettings => pptShape.AnimationSettings;

        public ActionSettings ActionSettings => pptShape.ActionSettings;

        public Tags Tags => pptShape.Tags;

        public PpMediaType MediaType => pptShape.MediaType;

        public MsoTriState HasTextFrame => pptShape.HasTextFrame;

        public SoundFormat SoundFormat => pptShape.SoundFormat;

        public Script Script => pptShape.Script;

        public string AlternativeText { get => pptShape.AlternativeText; set => pptShape.AlternativeText = value; }

        public MsoTriState HasTable => pptShape.HasTable;

        public Table Table => pptShape.Table;

        public MsoTriState HasDiagram => pptShape.HasDiagram;

        public Diagram Diagram => pptShape.Diagram;

        public MsoTriState HasDiagramNode => pptShape.HasDiagramNode;

        public PowerPnt.DiagramNode DiagramNode => pptShape.DiagramNode;

        public MsoTriState Child => pptShape.Child;

        public PowerPnt.Shape ParentGroup => pptShape.ParentGroup;

        public PowerPnt.CanvasShapes CanvasItems => pptShape.CanvasItems;

        public int Id => pptShape.Id;

        public string RTF { set => pptShape.RTF = value; }

        public CustomerData CustomerData => pptShape.CustomerData;

        public PowerPnt.TextFrame2 TextFrame2 => pptShape.TextFrame2;

        public MsoTriState HasChart => pptShape.HasChart;

        public MsoShapeStyleIndex ShapeStyle { get => pptShape.ShapeStyle; set => pptShape.ShapeStyle = value; }
        public MsoBackgroundStyleIndex BackgroundStyle { get => pptShape.BackgroundStyle; set => pptShape.BackgroundStyle = value; }

        public SoftEdgeFormat SoftEdge => pptShape.SoftEdge;

        public GlowFormat Glow => pptShape.Glow;

        public ReflectionFormat Reflection => pptShape.Reflection;

        public Chart Chart => pptShape.Chart;

        public MsoTriState HasSmartArt => pptShape.HasSmartArt;

        public SmartArt SmartArt => pptShape.SmartArt;

        public string Title { get => pptShape.Title; set => pptShape.Title = value; }

        public MediaFormat MediaFormat => pptShape.MediaFormat;

        WinWord.Shape WinWord.Shape.Duplicate()
        {
            return docShape.Duplicate();
        }

        public void Select(ref object Replace)
        {
            docShape.Select(ref Replace);
        }

        WinWord.ShapeRange WinWord.Shape.Ungroup()
        {
            return docShape.Ungroup();
        }

        public WinWord.InlineShape ConvertToInlineShape()
        {
            return docShape.ConvertToInlineShape();
        }

        public WinWord.Frame ConvertToFrame()
        {
            return docShape.ConvertToFrame();
        }

        public void Activate()
        {
            docShape.Activate();
        }

        WinWord.Application WinWord.Shape.Application => docShape.Application;

        WinWord.Adjustments WinWord.Shape.Adjustments => docShape.Adjustments;

        WinWord.CalloutFormat WinWord.Shape.Callout => docShape.Callout;

        WinWord.ConnectorFormat WinWord.Shape.ConnectorFormat => docShape.ConnectorFormat;

        WinWord.FillFormat WinWord.Shape.Fill => docShape.Fill;

        WinWord.GroupShapes WinWord.Shape.GroupItems => docShape.GroupItems;

        WinWord.LineFormat WinWord.Shape.Line => docShape.Line;

        WinWord.ShapeNodes WinWord.Shape.Nodes => docShape.Nodes;

        WinWord.PictureFormat WinWord.Shape.PictureFormat => docShape.PictureFormat;

        WinWord.ShadowFormat WinWord.Shape.Shadow => docShape.Shadow;

        WinWord.TextEffectFormat WinWord.Shape.TextEffect => docShape.TextEffect;

        WinWord.TextFrame WinWord.Shape.TextFrame => docShape.TextFrame;

        WinWord.ThreeDFormat WinWord.Shape.ThreeD => docShape.ThreeD;

        public WinWord.Hyperlink Hyperlink => docShape.Hyperlink;

        public WinWord.WdRelativeHorizontalPosition RelativeHorizontalPosition { get => docShape.RelativeHorizontalPosition; set => docShape.RelativeHorizontalPosition = value; }
        public WinWord.WdRelativeVerticalPosition RelativeVerticalPosition { get => docShape.RelativeVerticalPosition; set => docShape.RelativeVerticalPosition = value; }
        public int LockAnchor { get => docShape.LockAnchor; set => docShape.LockAnchor = value; }

        public WinWord.WrapFormat WrapFormat => docShape.WrapFormat;

        WinWord.OLEFormat WinWord.Shape.OLEFormat => docShape.OLEFormat;

        public WinWord.Range Anchor => docShape.Anchor;

        WinWord.LinkFormat WinWord.Shape.LinkFormat => docShape.LinkFormat;

        IMsoDiagram WinWord.Shape.Diagram => docShape.Diagram;

        WinWord.DiagramNode WinWord.Shape.DiagramNode => docShape.DiagramNode;

        WinWord.Shape WinWord.Shape.ParentGroup => docShape.ParentGroup;

        WinWord.CanvasShapes WinWord.Shape.CanvasItems => docShape.CanvasItems;

        public int ID => docShape.ID;

        public int LayoutInCell { get => docShape.LayoutInCell; set => docShape.LayoutInCell = value; }

        WinWord.Chart WinWord.Shape.Chart => docShape.Chart;

        public float LeftRelative { get => docShape.LeftRelative; set => docShape.LeftRelative = value; }
        public float TopRelative { get => docShape.TopRelative; set => docShape.TopRelative = value; }
        public float WidthRelative { get => docShape.WidthRelative; set => docShape.WidthRelative = value; }
        public float HeightRelative { get => docShape.HeightRelative; set => docShape.HeightRelative = value; }
        public WinWord.WdRelativeHorizontalSize RelativeHorizontalSize { get => docShape.RelativeHorizontalSize; set => docShape.RelativeHorizontalSize = value; }
        public WinWord.WdRelativeVerticalSize RelativeVerticalSize { get => docShape.RelativeVerticalSize; set => docShape.RelativeVerticalSize = value; }

        WinWord.SoftEdgeFormat WinWord.Shape.SoftEdge => docShape.SoftEdge;

        WinWord.GlowFormat WinWord.Shape.Glow => docShape.Glow;

        WinWord.ReflectionFormat WinWord.Shape.Reflection => docShape.Reflection;

        Microsoft.Office.Core.TextFrame2 WinWord.Shape.TextFrame2 => docShape.TextFrame2;

        public int AnchorID => docShape.AnchorID;

        public int EditID => docShape.EditID;
    }

}

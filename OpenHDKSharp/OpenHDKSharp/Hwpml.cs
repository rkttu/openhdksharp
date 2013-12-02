/*
 * Hwpml.cs
 *
 * Copyright (C) 2013 Junghyun Nam <rkttu@rkttu.com>
 * Based on Hodong Kim's libhwp project. <cogniti@gmail.com>
 * https://github.com/cogniti
 * 
 * This library is free software: you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as published
 * by the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This library is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

// TODO: parent = x as TParent를 value = x as TParent로 변경, as IHwpmlElement<TParent>로 검사

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.XPath;

namespace OpenHDKSharp.Markup
{
    public struct HwpUnit : IComparable, IFormattable, IComparable<int>, IEquatable<int>, IComparable<HwpUnit>, IEquatable<HwpUnit>
    {
        public HwpUnit(int value)
        {
            this.Value = value;
        }

        public int Value;

        public double ToPoint()
        { return this.Value / 100.0; }

        public override string ToString()
        { return String.Concat(this.Value, " hwpunit"); }

        public int CompareTo(object obj)
        { return this.Value.CompareTo(obj); }

        public string ToString(string format, IFormatProvider formatProvider)
        { return String.Concat(this.Value.ToString(format, formatProvider), " hwpunit"); }

        public int CompareTo(int other)
        { return this.Value.CompareTo(other); }

        public bool Equals(int other)
        { return this.Value.Equals(other); }

        public int CompareTo(HwpUnit other)
        { return this.Value.CompareTo(other.Value); }

        public bool Equals(HwpUnit other)
        { return this.Value.Equals(other.Value); }
    }

    public enum FontType : byte
    {
        rep,
        ttf,
        hft
    }

    public enum TextArtShapeFontType : byte
    {
        Unknown,
        ttf,
        hft
    }

    public enum LineType1 : byte
    {
        Solid,
        Dash,
        Dot,
        DashDot,
        DashDotDot,
        LongDash,
        Circle,
        DoubleSlim,
        SlimThick,
        ThickSlim,
        SlimThickSlim,
        None
    }

    public enum LineType2 : byte
    {
        Solid,
        Dash,
        Dot,
        DashDot,
        DashDotDot,
        LongDash,
        Circle,
        DoubleSlim,
        SlimThick,
        ThickSlim,
        SlimThickSlim
    }

    public enum LineType3 : byte
    {
        Solid,
        Dot,
        Thick,
        Dash,
        DashDot,
        DashDotDot
    }

    public struct LineWidth : IComparable, IFormattable, IComparable<float>, IEquatable<float>, IComparable<LineWidth>, IEquatable<LineWidth>
    {
        public LineWidth(float value)
        {
            this.Value = value;
        }

        public LineWidth(string value)
        {
            this.Value = Single.Parse(value.Replace("mm", String.Empty).Trim());
        }

        public float Value;

        public override string ToString()
        { return String.Concat(this.Value, "mm"); }

        public int CompareTo(object obj)
        { return this.Value.CompareTo(obj); }

        public string ToString(string format, IFormatProvider formatProvider)
        { return String.Concat(this.Value.ToString(format, formatProvider), "mm"); }

        public int CompareTo(float other)
        { return this.Value.CompareTo(other); }

        public bool Equals(float other)
        { return this.Value.Equals(other); }

        public int CompareTo(LineWidth other)
        { return this.Value.CompareTo(other.Value); }

        public bool Equals(LineWidth other)
        { return this.Value.Equals(other.Value); }
    }

    public struct RGBColor : IComparable, IFormattable, IComparable<long>, IEquatable<long>, IComparable<RGBColor>, IEquatable<RGBColor>
    {
        public RGBColor(long value)
        {
            this.Value = value;
        }

        public long Value;

        public string ToRGBString()
        { return String.Concat("#", this.Value.ToString("X6")); }

        public string ToARGBString()
        { return String.Concat("#", this.Value.ToString("X8")); }

        public override string ToString()
        { return this.ToARGBString(); }

        public int CompareTo(object obj)
        { return this.Value.CompareTo(obj); }

        public string ToString(string format, IFormatProvider formatProvider)
        { return this.Value.ToString(format, formatProvider); }

        public int CompareTo(long other)
        { return this.Value.CompareTo(other); }

        public bool Equals(long other)
        { return this.Value.Equals(other); }

        public int CompareTo(RGBColor other)
        { return this.Value.CompareTo(other.Value); }

        public bool Equals(RGBColor other)
        { return this.Value.Equals(other.Value); }
    }

    public enum NumberType1 : byte
    {
        Digit,
        CircledDigit,
        RomanCapital,
        RomanSmall,
        LatinCapital,
        LatinSmall,
        CircledLatinCapital,
        CircledLatinSmall,
        HangulSyllable,
        CircledHangulSyllable,
        HangulJamo,
        CircledHangulJamo,
        HangulPhonetic,
        Ideograph,
        CircledIdeograph
    }

    public enum NumberType2 : byte
    {
        Digit,
        CircledDigit,
        RomanCapital,
        RomanSmall,
        LatinCapital,
        LatinSmall,
        CircledLatinCapital,
        CircledLatinSmall,
        HangulSyllable,
        CircledHangulSyllable,
        HangulJamo,
        CircledHangulJamo,
        HangulPhonetic,
        Ideograph,
        CircledIdeograph,
        DecagonCircle,
        DecagonCircleHanja,
        Symbol,
        UserChar
    }

    public enum AlignmentType1 : byte
    {
        Justify,
        Left,
        Right,
        Center,
        Distribute,
        DistributeSpace
    }

    public enum AlignmentType2 : byte
    {
        Left,
        Center,
        Right
    }

    public enum ArrowType : byte
    {
        Normal,
        Arrow,
        Spear,
        ConcaveArrow,
        EmptyDiamond,
        EmptyCircle,
        EmptyBox,
        FilledDiamond,
        FilledCircle,
        FilledBox
    }

    public enum ArrowSize : byte
    {
        SmallSmall,
        SmallMedium,
        SmallLarge,
        MediumSmall,
        MediumMedium,
        MediumLarge,
        LargeSmall,
        LargeMedium,
        LargeLarge
    }

    public enum LangType : byte
    {
        Hangul,
        Latin,
        Hanja,
        Japanese,
        Other,
        Symbol,
        User
    }

    public enum HatchStyle : byte
    {
        Horizontal,
        Vertical,
        BackSlash,
        Slash,
        Cross,
        CrossDiagonal
    }

    public enum InfillMode : byte
    {
        Tile,
        TileHorzTop,
        TileHorzBottom,
        TileVertLeft,
        TileVertRight,
        Total,
        Center,
        CenterTop,
        CenterBottom,
        LeftCenter,
        LeftTop,
        LeftBottom,
        RightCenter,
        RightTop,
        RightBottom,
        Zoom
    }

    public enum LineWrapType : byte
    {
        Break,
        Squeeze,
        Keep
    }

    public enum TextWrapType : byte
    {
        Square,
        Tight,
        Through,
        TopAndBottom,
        BehindText,
        InFrontOfText
    }

    public enum FieldType : byte
    {
        Clickhere,
        Hyperlink,
        Bookmark,
        Formula,
        Summery,
        UserInfo,
        Date,
        DocDate,
        Path,
        Crossref,
        Mailmerge,
        Memo,
        RevisionChange,
        RevisionSign,
        RevisionDelete,
        RevisionAttach,
        RevisionClipping,
        RevisionSawtooth,
        RevisionThinking,
        RevisionPraise,
        RevisionLine,
        RevisionSimpleChange,
        RevisionHyperlink,
        RevisionLineAttach,
        RevisionLineLink,
        RevisionLineTransfer,
        RevisionRightmove,
        RevisionLeftmove,
        RevisionTransfer,
        RevisionSplit
    }

    public interface IHwpmlElement
    {
        IHwpmlElement Parent { get; set; }
        string ElementName { get; }
        string ToString();
    }

    public interface ITextElement : IHwpmlElement
    {
        string Content { get; set; }
    }

    public interface IHwpmlElement<TParent> : IHwpmlElement
        where TParent: class, IHwpmlElement
    {
        new TParent Parent { get; set; }
    }
    
    public sealed class HwpmlDocument : IHwpmlElement
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this; }
            set { }
        }

        public string ElementName { get { return "#HwpmlDocument"; } }
        public HwpmlElement Hwpml;

        public static HwpmlDocument Load(IXPathNavigable target)
        {
            if (target == null)
                return null;

            HwpmlDocument item = new HwpmlDocument();

            XPathNavigator nav = target.CreateNavigator();
            item.Hwpml = HwpmlElement.Create(nav.SelectSingleNode("HWPML"), item);
            return item;
        }

        public static HwpmlDocument Load(string fileName)
        {
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            doc.Load(fileName);
            return Load(doc);
        }
    }

    public enum HwpmlStyle2 : byte
    {
        embed,
        export
    }

    public sealed class HwpmlElement : IHwpmlElement, IHwpmlElement<HwpmlDocument>
    {
        public HwpmlElement()
            : base()
        {
            this.Version = "2.8";
            this.SubVersion = "8.0.0.0";
            this.Style2 = HwpmlStyle2.embed;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HwpmlDocument)value; }
        }

        public string ElementName { get { return "HWPML"; } }
        public HwpmlDocument Parent { get; set; }

        public HeadElement Head;
        public BodyElement Body;
        public TailElement Tail;

        public string Version;
        public string SubVersion;
        public HwpmlStyle2 Style2;

        public static HwpmlElement Create(IXPathNavigable target, HwpmlDocument parent)
        {
            if (target == null)
                return null;

            HwpmlElement item = new HwpmlElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.Head = HeadElement.Create(nav.SelectSingleNode("HEAD"), item);
            item.Body = BodyElement.Create(nav.SelectSingleNode("BODY"), item);
            item.Tail = TailElement.Create(nav.SelectSingleNode("TAIL"), item);

            item.Version = String.Concat(nav.SelectSingleNode("@Version"));
            item.SubVersion = String.Concat(nav.SelectSingleNode("@SubVersion"));

            switch (String.Concat(nav.SelectSingleNode("@Style2")).ToUpperInvariant())
            {
                case "EMBED":
                    item.Style2 = HwpmlStyle2.embed;
                    break;
                case "EXPORT":
                    item.Style2 = HwpmlStyle2.export;
                    break;
            }

            return item;
        }
    }

    public sealed class HeadElement : IHwpmlElement, IHwpmlElement<HwpmlElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HwpmlElement)value; }
        }

        public string ElementName { get { return "HEAD"; } }
        public HwpmlElement Parent { get; set; }

        public DocSummaryElement DocSummary;
        public DocSettingElement DocSetting;
        public MappingTableElement MappingTable;
        public CompatibleDocumentElement CompatibleDocument;

        public int SecCnt;

        public static HeadElement Create(IXPathNavigable target, HwpmlElement parent)
        {
            if (target == null)
                return null;

            HeadElement item = new HeadElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.DocSummary = DocSummaryElement.Create(nav.SelectSingleNode("DOCSUMMARY"), item);
            item.DocSetting = DocSettingElement.Create(nav.SelectSingleNode("DOCSETTING"), item);
            item.MappingTable = MappingTableElement.Create(nav.SelectSingleNode("MAPPINGTABLE"), item);
            item.CompatibleDocument = CompatibleDocumentElement.Create(nav.SelectSingleNode("COMPATIBLEDOCUMENT"), item);

            try { item.SecCnt = Int32.Parse(String.Concat(nav.SelectSingleNode("@SECCNT"))); }
            catch { item.SecCnt = 0; }

            return item;
        }
    }

    public sealed class DocSummaryElement : IHwpmlElement, IHwpmlElement<HeadElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HeadElement)value; }
        }

        public string ElementName { get { return "DOCSUMMARY"; } }
        public HeadElement Parent { get; set; }

        public string Title;
        public string Subject;
        public string Author;
        public string Date;
        public string Keywords;
        public string Comments;
        public ForbiddenStringElement ForbiddenString;

        public static DocSummaryElement Create(IXPathNavigable target, HeadElement parent)
        {
            if (target == null)
                return null;

            DocSummaryElement item = new DocSummaryElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.Title = String.Concat(nav.SelectSingleNode("TITLE"));
            item.Subject = String.Concat(nav.SelectSingleNode("SUBJECT"));
            item.Author = String.Concat(nav.SelectSingleNode("AUTHOR"));
            item.Date = String.Concat(nav.SelectSingleNode("DATE"));
            item.Keywords = String.Concat(nav.SelectSingleNode("KEYWORDS"));
            item.Comments = String.Concat(nav.SelectSingleNode("COMMENTS"));
            item.ForbiddenString = ForbiddenStringElement.Create(nav.SelectSingleNode("FORBIDDENSTRING"), item);

            return item;
        }
    }

    public sealed class ForbiddenStringElement : List<ForbiddenElement>, IHwpmlElement, IHwpmlElement<DocSummaryElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (DocSummaryElement)value; }
        }

        public string ElementName { get { return "FORBIDDENSTRING"; } }
        public DocSummaryElement Parent { get; set; }

        public static ForbiddenStringElement Create(IXPathNavigable target, DocSummaryElement parent)
        {
            if (target == null)
                return null;

            ForbiddenStringElement item = new ForbiddenStringElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("FORBIDDEN"))
                item.Add(ForbiddenElement.Create(each, item));

            return item;
        }
    }

    public sealed class ForbiddenElement : IHwpmlElement, IHwpmlElement<ForbiddenStringElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ForbiddenStringElement)value; }
        }

        public string ElementName { get { return "FORBIDDEN"; } }
        public ForbiddenStringElement Parent { get; set; }

        public string Id;

        public static ForbiddenElement Create(IXPathNavigable target, ForbiddenStringElement parent)
        {
            if (target == null)
                return null;

            ForbiddenElement item = new ForbiddenElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.Id = String.Concat(nav.SelectSingleNode("@id"));

            return item;
        }
    }

    public sealed class DocSettingElement : IHwpmlElement, IHwpmlElement<HeadElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HeadElement)value; }
        }

        public string ElementName { get { return "DOCSETTING"; } }
        public HeadElement Parent { get; set; }

        public BeginNumberElement BeginNumber;
        public CaretPosElement CaretPos;

        public static DocSettingElement Create(IXPathNavigable target, HeadElement parent)
        {
            if (target == null)
                return null;

            DocSettingElement item = new DocSettingElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.BeginNumber = BeginNumberElement.Create(nav.SelectSingleNode("BEGINNUMBER"), item);
            item.CaretPos = CaretPosElement.Create(nav.SelectSingleNode("CARETPOS"), item);

            return item;
        }
    }

    public sealed class BeginNumberElement : IHwpmlElement, IHwpmlElement<DocSettingElement>
    {
        public BeginNumberElement()
            : base()
        {
            this.Page = 1;
            this.Footnote = 1;
            this.Endnote = 1;
            this.Picture = 1;
            this.Table = 1;
            this.Equation = 1;
            this.TotalPage = 1;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (DocSettingElement)value; }
        }

        public string ElementName { get { return "BEGINNUMBER"; } }
        public DocSettingElement Parent { get; set; }

        public int Page;
        public int Footnote;
        public int Endnote;
        public int Picture;
        public int Table;
        public int Equation;
        public int TotalPage;

        public static BeginNumberElement Create(IXPathNavigable target, DocSettingElement parent)
        {
            if (target == null)
                return null;

            BeginNumberElement item = new BeginNumberElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Page = Int32.Parse(String.Concat(nav.SelectSingleNode("@Page"))); }
            catch { }

            try { item.Footnote = Int32.Parse(String.Concat(nav.SelectSingleNode("@Footnote"))); }
            catch { }

            try { item.Endnote = Int32.Parse(String.Concat(nav.SelectSingleNode("@Endnote"))); }
            catch { }

            try { item.Picture = Int32.Parse(String.Concat(nav.SelectSingleNode("@Picture"))); }
            catch { }

            try { item.Table = Int32.Parse(String.Concat(nav.SelectSingleNode("@Table"))); }
            catch { }

            try { item.Equation = Int32.Parse(String.Concat(nav.SelectSingleNode("@Equation"))); }
            catch { }

            try { item.TotalPage = Int32.Parse(String.Concat(nav.SelectSingleNode("@TotalPage"))); }
            catch { }

            return item;
        }
    }

    public sealed class CaretPosElement : IHwpmlElement, IHwpmlElement<DocSettingElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (DocSettingElement)value; }
        }

        public string ElementName { get { return "CARETPOS"; } }
        public DocSettingElement Parent { get; set; }

        public string List;
        public string Para;
        public string Pos;

        public static CaretPosElement Create(IXPathNavigable target, DocSettingElement parent)
        {
            if (target == null)
                return null;

            CaretPosElement item = new CaretPosElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.List = String.Concat(nav.SelectSingleNode("@List"));
            item.Para = String.Concat(nav.SelectSingleNode("@Para"));
            item.Pos = String.Concat(nav.SelectSingleNode("@Pos"));

            return item;
        }
    }

    public sealed class MappingTableElement : IHwpmlElement, IHwpmlElement<HeadElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HeadElement)value; }
        }

        public string ElementName { get { return "MAPPINGTABLE"; } }
        public HeadElement Parent { get; set; }

        public BinDataListElement BinDataList;
        public FaceNameListElement FaceNameList;
        public BorderFillListElement BorderFillList;
        public CharShapeListElement CharShapeList;
        public TabDefListElement TabDefList;
        public NumberingListElement NumberingList;
        public BulletListElement BulletList;
        public ParaShapeListElement ParaShapeList;
        public StyleListElement StyleList;
        public MemoShapeListElement MemoShapeList;

        public static MappingTableElement Create(IXPathNavigable target, HeadElement parent)
        {
            if (target == null)
                return null;

            MappingTableElement item = new MappingTableElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.BinDataList = BinDataListElement.Create(nav.SelectSingleNode("BINDATALIST"), item);
            item.FaceNameList = FaceNameListElement.Create(nav.SelectSingleNode("FACENAMELIST"), item);
            item.BorderFillList = BorderFillListElement.Create(nav.SelectSingleNode("BORDERFILLLIST"), item);
            item.CharShapeList = CharShapeListElement.Create(nav.SelectSingleNode("CHARSHAPELIST"), item);
            item.TabDefList = TabDefListElement.Create(nav.SelectSingleNode("TABDEFLIST"), item);
            item.NumberingList = NumberingListElement.Create(nav.SelectSingleNode("NUMBERINGLIST"), item);
            item.BulletList = BulletListElement.Create(nav.SelectSingleNode("BULLETLIST"), item);
            item.ParaShapeList = ParaShapeListElement.Create(nav.SelectSingleNode("PARASHAPELIST"), item);
            item.StyleList = StyleListElement.Create(nav.SelectSingleNode("STYLELIST"), item);
            item.MemoShapeList = MemoShapeListElement.Create(nav.SelectSingleNode("MEMOSHAPELIST"), item);

            return item;
        }
    }

    public sealed class BinDataListElement : List<BinItemElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "BINDATALIST"; } }
        public MappingTableElement Parent { get; set; }

        public static BinDataListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            BinDataListElement item = new BinDataListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("BINITEM"))
                item.Add(BinItemElement.Create(each, item));

            return item;
        }
    }

    public enum BinItemType : byte
    {
        Link,
        Embedding,
        Storage
    }

    public enum BinItemFormat : byte
    {
        jpg,
        bmp,
        gif,
        ole
    }

    public sealed class BinItemElement : IHwpmlElement, IHwpmlElement<BinDataListElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BinDataListElement)value; }
        }

        public string ElementName { get { return "BINITEM"; } }
        public BinDataListElement Parent { get; set; }

        public BinItemType Type;
        public string APath;
        public string RPath;
        public string BinData;
        public BinItemFormat Format;

        public static BinItemElement Create(IXPathNavigable target, BinDataListElement parent)
        {
            if (target == null)
                return null;

            BinItemElement item = new BinItemElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (BinItemType)Enum.Parse(typeof(BinItemType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            item.APath = String.Concat(nav.SelectSingleNode("@APath"));
            item.RPath = String.Concat(nav.SelectSingleNode("@RPath"));
            item.BinData = String.Concat(nav.SelectSingleNode("@BinData"));

            try { item.Format = (BinItemFormat)Enum.Parse(typeof(BinItemFormat), String.Concat(nav.SelectSingleNode("@Format")), true); }
            catch { }

            return item;
        }
    }

    public sealed class FaceNameListElement : List<FontFaceElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "FACENAMELIST"; } }
        public MappingTableElement Parent { get; set; }

        public static FaceNameListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            FaceNameListElement item = new FaceNameListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("FONTFACE"))
                item.Add(FontFaceElement.Create(each, item));

            return item;
        }
    }

    public sealed class FontFaceElement : List<FontElement>, IHwpmlElement, IHwpmlElement<FaceNameListElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FaceNameListElement)value; }
        }

        public string ElementName { get { return "FONTFACE"; } }
        public FaceNameListElement Parent { get; set; }

        public LangType Lang;

        public static FontFaceElement Create(IXPathNavigable target, FaceNameListElement parent)
        {
            if (target == null)
                return null;

            FontFaceElement item = new FontFaceElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("FONT"))
                item.Add(FontElement.Create(each, item));

            try { item.Lang = (LangType)Enum.Parse(typeof(LangType), String.Concat(nav.SelectSingleNode("@Lang")), true); }
            catch { }

            return item;
        }
    }

    public sealed class FontElement : IHwpmlElement, IHwpmlElement<FontFaceElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FontFaceElement)value; }
        }

        public string ElementName { get { return "FONT"; } }
        public FontFaceElement Parent { get; set; }

        public SubstFontElement SubstFont;
        public TypeInfoElement TypeInfo;

        public string Id;
        public FontType Type;
        public string Name;

        public static FontElement Create(IXPathNavigable target, FontFaceElement parent)
        {
            if (target == null)
                return null;

            FontElement item = new FontElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.SubstFont = SubstFontElement.Create(nav.SelectSingleNode("SUBSTFONT"), item);
            item.TypeInfo = TypeInfoElement.Create(nav.SelectSingleNode("TYPEINFO"), item);

            item.Id = String.Concat(nav.SelectSingleNode("@Id"));
            
            try { item.Type = (FontType)Enum.Parse(typeof(FontType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));

            return item;
        }
    }

    public sealed class SubstFontElement : IHwpmlElement, IHwpmlElement<FontElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FontElement)value; }
        }

        public string ElementName { get { return "SUBSTFONT"; } }
        public FontElement Parent { get; set; }

        public FontType Type;
        public string Name;

        public static SubstFontElement Create(IXPathNavigable target, FontElement parent)
        {
            if (target == null)
                return null;

            SubstFontElement item = new SubstFontElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (FontType)Enum.Parse(typeof(FontType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));

            return item;
        }
    }

    public sealed class TypeInfoElement : IHwpmlElement, IHwpmlElement<FontElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FontElement)value; }
        }

        public string ElementName { get { return "TYPEINFO"; } }
        public FontElement Parent { get; set; }

        public string FamilyType;
        public string SerifStyle;
        public string Weight;
        public string Proportion;
        public string Contrast;
        public string StrokeVariation;
        public string ArmStyle;
        public string Letterform;
        public string Midline;
        public string XHeight;

        public static TypeInfoElement Create(IXPathNavigable target, FontElement parent)
        {
            if (target == null)
                return null;

            TypeInfoElement item = new TypeInfoElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.FamilyType = String.Concat(nav.SelectSingleNode("@FamilyType"));
            item.SerifStyle = String.Concat(nav.SelectSingleNode("@SerifStyle"));
            item.Weight = String.Concat(nav.SelectSingleNode("@Weight"));
            item.Proportion = String.Concat(nav.SelectSingleNode("@Proportion"));
            item.Contrast = String.Concat(nav.SelectSingleNode("@Contrast"));
            item.StrokeVariation = String.Concat(nav.SelectSingleNode("@StrokeVariation"));
            item.ArmStyle = String.Concat(nav.SelectSingleNode("@ArmStyle"));
            item.Letterform = String.Concat(nav.SelectSingleNode("@LetterForm"));
            item.Midline = String.Concat(nav.SelectSingleNode("@Midline"));
            item.XHeight = String.Concat(nav.SelectSingleNode("@XHeight"));

            return item;
        }
    }

    public sealed class BorderFillListElement : List<BorderFillElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "BORDERFILLLIST"; } }
        public MappingTableElement Parent { get; set; }

        public static BorderFillListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            BorderFillListElement item = new BorderFillListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("BORDERFILL"))
                item.Add(BorderFillElement.Create(each, item));

            return item;
        }
    }

    public sealed class BorderFillElement : IHwpmlElement, IHwpmlElement<BorderFillListElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BorderFillListElement)value; }
        }

        public string ElementName { get { return "BORDERFILL"; } }
        public BorderFillListElement Parent { get; set; }

        public LeftBorderElement LeftBorder;
        public RightBorderElement RightBorder;
        public TopBorderElement TopBorder;
        public BottomBorderElement BottomBorder;
        public DiagonalBorderElement Diagonal;
        public FillBrushElement FillBrush;

        public int Id;
        public bool ThreeD;
        public bool Shadow;
        public int Slash;
        public int BackSlash;
        public int CrookedSlash;
        public int CounterSlash;
        public int CounterBackSlash;
        public int BreakCellSeparateLine;

        public static BorderFillElement Create(IXPathNavigable target, BorderFillListElement parent)
        {
            if (target == null)
                return null;

            BorderFillElement item = new BorderFillElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.LeftBorder = LeftBorderElement.Create(nav.SelectSingleNode("LEFTBORDER"), item);
            item.RightBorder = RightBorderElement.Create(nav.SelectSingleNode("RIGHTBORDER"), item);
            item.TopBorder = TopBorderElement.Create(nav.SelectSingleNode("TOPBORDER"), item);
            item.BottomBorder = BottomBorderElement.Create(nav.SelectSingleNode("BOTTOMBORDER"), item);
            item.Diagonal = DiagonalBorderElement.Create(nav.SelectSingleNode("DIAGONAL"), item);
            item.FillBrush = FillBrushElement.Create(nav.SelectSingleNode("FILLBRUSH"), item);

            try { item.Id = Int32.Parse(String.Concat(nav.SelectSingleNode("@Id"))); }
            catch { }

            try { item.ThreeD = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ThreeD"))); }
            catch { }

            try { item.Shadow = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Shadow"))); }
            catch { }

            try { item.Slash = Int32.Parse(String.Concat(nav.SelectSingleNode("@Slash"))); }
            catch { }

            try { item.BackSlash = Int32.Parse(String.Concat(nav.SelectSingleNode("@BackSlash"))); }
            catch { }

            try { item.CrookedSlash = Int32.Parse(String.Concat(nav.SelectSingleNode("@CrookedSlash"))); }
            catch { }

            try { item.CounterSlash = Int32.Parse(String.Concat(nav.SelectSingleNode("@CounterSlash"))); }
            catch { }

            try { item.CounterBackSlash = Int32.Parse(String.Concat(nav.SelectSingleNode("@CounterBackSlash"))); }
            catch { }

            try { item.BreakCellSeparateLine = Int32.Parse(String.Concat(nav.SelectSingleNode("@BreakCellSeparateLine"))); }
            catch { }

            return item;
        }
    }

    public sealed class LeftBorderElement : IHwpmlElement, IHwpmlElement<BorderFillElement>
    {
        public LeftBorderElement()
            : base()
        {
            this.Type = LineType1.Solid;
            this.Width = new LineWidth("0.12mm");
            this.Color = new RGBColor(0L);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BorderFillElement)value; }
        }

        public string ElementName { get { return "LEFTBORDER"; } }
        public BorderFillElement Parent { get; set; }

        public LineType1 Type;
        public LineWidth Width;
        public RGBColor Color;

        public static LeftBorderElement Create(IXPathNavigable target, BorderFillElement parent)
        {
            if (target == null)
                return null;

            LeftBorderElement item = new LeftBorderElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Width = new LineWidth(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class RightBorderElement : IHwpmlElement, IHwpmlElement<BorderFillElement>
    {
        public RightBorderElement()
            : base()
        {
            this.Type = LineType1.Solid;
            this.Width = new LineWidth("0.12mm");
            this.Color = new RGBColor(0L);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BorderFillElement)value; }
        }

        public string ElementName { get { return "RIGHTBORDER"; } }
        public BorderFillElement Parent { get; set; }

        public LineType1 Type;
        public LineWidth Width;
        public RGBColor Color;

        public static RightBorderElement Create(IXPathNavigable target, BorderFillElement parent)
        {
            if (target == null)
                return null;

            RightBorderElement item = new RightBorderElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Width = new LineWidth(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class TopBorderElement : IHwpmlElement, IHwpmlElement<BorderFillElement>
    {
        public TopBorderElement()
            : base()
        {
            this.Type = LineType1.Solid;
            this.Width = new LineWidth("0.12mm");
            this.Color = new RGBColor(0L);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BorderFillElement)value; }
        }

        public string ElementName { get { return "TOPBORDER"; } }
        public BorderFillElement Parent { get; set; }

        public LineType1 Type;
        public LineWidth Width;
        public RGBColor Color;

        public static TopBorderElement Create(IXPathNavigable target, BorderFillElement parent)
        {
            if (target == null)
                return null;

            TopBorderElement item = new TopBorderElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Width = new LineWidth(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class BottomBorderElement : IHwpmlElement, IHwpmlElement<BorderFillElement>
    {
        public BottomBorderElement()
            : base()
        {
            this.Type = LineType1.Solid;
            this.Width = new LineWidth("0.12mm");
            this.Color = new RGBColor(0L);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BorderFillElement)value; }
        }

        public string ElementName { get { return "BOTTOMBORDER"; } }
        public BorderFillElement Parent { get; set; }

        public LineType1 Type;
        public LineWidth Width;
        public RGBColor Color;

        public static BottomBorderElement Create(IXPathNavigable target, BorderFillElement parent)
        {
            if (target == null)
                return null;

            BottomBorderElement item = new BottomBorderElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Width = new LineWidth(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class DiagonalBorderElement : IHwpmlElement, IHwpmlElement<BorderFillElement>
    {
        public DiagonalBorderElement()
            : base()
        {
            this.Type = LineType1.Solid;
            this.Width = new LineWidth("0.12mm");
            this.Color = new RGBColor(0L);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BorderFillElement)value; }
        }

        public string ElementName { get { return "DIAGONALBORDER"; } }
        public BorderFillElement Parent { get; set; }

        public LineType1 Type;
        public LineWidth Width;
        public RGBColor Color;

        public static DiagonalBorderElement Create(IXPathNavigable target, BorderFillElement parent)
        {
            if (target == null)
                return null;

            DiagonalBorderElement item = new DiagonalBorderElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Width = new LineWidth(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class FillBrushElement : IHwpmlElement, IHwpmlElement<BorderFillElement>, IHwpmlElement<DrawingObjectElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as BorderFillElement) != null)
                    return;
                if ((this.parent = value as DrawingObjectElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        BorderFillElement IHwpmlElement<BorderFillElement>.Parent
        {
            get { return this.parent as BorderFillElement; }
            set { this.parent = value; } 
        }

        DrawingObjectElement IHwpmlElement<DrawingObjectElement>.Parent
        {
            get { return this.parent as DrawingObjectElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "FILLBRUSH"; } }

        public WindowBrushElement WindowBrush;
        public GradationElement Gradation;
        public ImageBrushElement ImageBrush;

        public static FillBrushElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            FillBrushElement item = new FillBrushElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.WindowBrush = WindowBrushElement.Create(nav.SelectSingleNode("WINDOWBRUSH"), item);
            item.Gradation = GradationElement.Create(nav.SelectSingleNode("GRADATION"), item);
            item.ImageBrush = ImageBrushElement.Create(nav.SelectSingleNode("IMAGEBRUSH"), item);

            return item;
        }
    }

    public sealed class WindowBrushElement : IHwpmlElement, IHwpmlElement<FillBrushElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FillBrushElement)value; }
        }

        public string ElementName { get { return "WINDOWBRUSH"; } }
        public FillBrushElement Parent { get; set; }

        public RGBColor FaceColor;
        public RGBColor HatchColor;
        public HatchStyle HatchStyle;
        public string Alpha;

        public static WindowBrushElement Create(IXPathNavigable target, FillBrushElement parent)
        {
            if (target == null)
                return null;

            WindowBrushElement item = new WindowBrushElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.FaceColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@FaceColor")))); }
            catch { }

            try { item.HatchColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@HatchColor")))); }
            catch { }

            try { item.HatchStyle = (HatchStyle)Enum.Parse(typeof(HatchStyle), String.Concat(nav.SelectSingleNode("@HatchStyle")), true); }
            catch { }

            item.Alpha = String.Concat(nav.SelectSingleNode("@Alpha"));

            return item;
        }
    }

    public enum GradationType : byte
    {
        Linear,
        Radial,
        Conical,
        Square
    }

    public sealed class GradationElement : List<ColorElement>, IHwpmlElement, IHwpmlElement<FillBrushElement>
    {
        public GradationElement()
            : base()
        {
            this.Angle = 90d;
            this.Step = 50;
            this.ColorNum = 2;
            this.StepCenter = 50;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FillBrushElement)value; }
        }

        public string ElementName { get { return "GRADATION"; } }
        public FillBrushElement Parent { get; set; }

        public GradationType Type;
        public double Angle;
        public double CenterX;
        public double CenterY;
        public int Step;
        public int ColorNum;
        public int StepCenter;
        public string Alpha;

        public static GradationElement Create(IXPathNavigable target, FillBrushElement parent)
        {
            if (target == null)
                return null;

            GradationElement item = new GradationElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("COLOR"))
                item.Add(ColorElement.Create(each, item));

            try { item.Type = (GradationType)Enum.Parse(typeof(GradationType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Angle = Double.Parse(String.Concat(nav.SelectSingleNode("@Angle"))); }
            catch { }

            try { item.CenterX = Double.Parse(String.Concat(nav.SelectSingleNode("@CenterX"))); }
            catch { }

            try { item.CenterY = Double.Parse(String.Concat(nav.SelectSingleNode("@CenterY"))); }
            catch { }

            try { item.Step = Int32.Parse(String.Concat(nav.SelectSingleNode("@Step"))); }
            catch { }

            try { item.ColorNum = Int32.Parse(String.Concat(nav.SelectSingleNode("@ColorNum"))); }
            catch { }

            try { item.StepCenter = Int32.Parse(String.Concat(nav.SelectSingleNode("@StepCenter"))); }
            catch { }

            item.Alpha = String.Concat(nav.SelectSingleNode("@Alpha"));

            return item;
        }
    }

    public sealed class ColorElement : IHwpmlElement, IHwpmlElement<GradationElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (GradationElement)value; }
        }

        public string ElementName { get { return "COLOR"; } }
        public GradationElement Parent { get; set; }

        public RGBColor Value;

        public static ColorElement Create(IXPathNavigable target, GradationElement parent)
        {
            if (target == null)
                return null;

            ColorElement item = new ColorElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Value = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Value")))); }
            catch { }

            return item;
        }
    }

    public sealed class ImageBrushElement : IHwpmlElement, IHwpmlElement<FillBrushElement>
    {
        public ImageBrushElement()
            : base()
        {
            this.Mode = InfillMode.Tile;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FillBrushElement)value; }
        }

        public string ElementName { get { return "IMAGEBRUSH"; } }
        public FillBrushElement Parent { get; set; }

        public ImageElement Image;

        public InfillMode Mode;

        public static ImageBrushElement Create(IXPathNavigable target, FillBrushElement parent)
        {
            if (target == null)
                return null;

            ImageBrushElement item = new ImageBrushElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.Image = ImageElement.Create(nav.SelectSingleNode("IMAGE"), item);

            try { item.Mode = (InfillMode)Enum.Parse(typeof(InfillMode), String.Concat(nav.SelectSingleNode("@Mode")), true); }
            catch { }

            return item;
        }
    }

    public enum ImageEffect : byte
    {
        RealPic,
        GrayScale,
        BlackWhite
    }

    public sealed class ImageElement : IHwpmlElement, IHwpmlElement<ImageBrushElement>, IHwpmlElement<PictureElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as ImageBrushElement) != null)
                    return;
                if ((this.parent = value as PictureElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        ImageBrushElement IHwpmlElement<ImageBrushElement>.Parent
        {
            get { return this.parent as ImageBrushElement; }
            set { this.parent = value; }
        }

        PictureElement IHwpmlElement<PictureElement>.Parent
        {
            get { return this.parent as PictureElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "IMAGE"; } }

        public int Bright;
        public int Contrast;
        public ImageEffect Effect;
        public string BinItem;
        public string Alpha;

        public static ImageElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ImageElement item = new ImageElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Bright = Int32.Parse(String.Concat(nav.SelectSingleNode("@Bright"))); }
            catch { }

            try { item.Contrast = Int32.Parse(String.Concat(nav.SelectSingleNode("@Contrast"))); }
            catch { }

            try { item.Effect = (ImageEffect)Enum.Parse(typeof(ImageEffect), String.Concat(nav.SelectSingleNode("@Effect")), true); }
            catch { }

            item.BinItem = String.Concat(nav.SelectSingleNode("@BinItem"));
            item.Alpha = String.Concat(nav.SelectSingleNode("@Alpha"));

            return item;
        }
    }

    public sealed class CharShapeListElement : List<CharShapeElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "CHARSHAPELIST"; } }
        public MappingTableElement Parent { get; set; }

        public static CharShapeListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            CharShapeListElement item = new CharShapeListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("CHARSHAPE"))
                item.Add(CharShapeElement.Create(each, item));

            return item;
        }
    }

    public sealed class CharShapeElement : IHwpmlElement, IHwpmlElement<CharShapeListElement>
    {
        public CharShapeElement()
            : base()
        {
            this.Height = new HwpUnit(1000);
            this.ShadeColor = new RGBColor(4294967295L);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeListElement)value; }
        }

        public string ElementName { get { return "CHARSHAPE"; } }
        public CharShapeListElement Parent { get; set; }

        public FontIdElement FontId;
        public RatioElement Ratio;
        public CharSpacingElement CharSpacing;
        public RelSizeElement RelSize;
        public CharOffsetElement CharOffset;
        public bool Italic;
        public bool Bold;
        public UnderlineElement Underline;
        public OutlineElement Outline;
        public ShadowElement Shadow;
        public bool Emboss;
        public bool Engrave;
        public bool Superscript;
        public bool Subscript;

        public int Id;
        public HwpUnit Height;
        public RGBColor TextColor;
        public RGBColor ShadeColor;
        public bool UseFontSpace;
        public bool UseKerning;
        public int SymMark;
        public int BorderFillId;

        public static CharShapeElement Create(IXPathNavigable target, CharShapeListElement parent)
        {
            if (target == null)
                return null;

            CharShapeElement item = new CharShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.FontId = FontIdElement.Create(nav.SelectSingleNode("FONTID"), item);
            item.Ratio = RatioElement.Create(nav.SelectSingleNode("RATIO"), item);
            item.CharSpacing = CharSpacingElement.Create(nav.SelectSingleNode("CHARSPACING"), item);
            item.RelSize = RelSizeElement.Create(nav.SelectSingleNode("RELSIZE"), item);
            item.CharOffset = CharOffsetElement.Create(nav.SelectSingleNode("CHAROFFSET"), item);
            item.Italic = nav.SelectSingleNode("ITALIC") != null;
            item.Bold = nav.SelectSingleNode("BOLD") != null;
            item.Underline = UnderlineElement.Create(nav.SelectSingleNode("UNDERLINE"), item);
            item.Outline = OutlineElement.Create(nav.SelectSingleNode("OUTLINE"), item);
            item.Shadow = ShadowElement.Create(nav.SelectSingleNode("SHADOW"), item);
            item.Emboss = nav.SelectSingleNode("EMBOSS") != null;
            item.Engrave = nav.SelectSingleNode("ENGRAVE") != null;
            item.Superscript = nav.SelectSingleNode("SUPERSCRIPT") != null;
            item.Subscript = nav.SelectSingleNode("SUBSCRIPT") != null;

            try { item.Id = Int32.Parse(String.Concat(nav.SelectSingleNode("@Id"))); }
            catch { }

            try { item.Height = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Height")))); }
            catch { }

            try { item.TextColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@TextColor")))); }
            catch { }

            try { item.ShadeColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@ShadeColor")))); }
            catch { }

            try { item.UseFontSpace = Boolean.Parse(String.Concat(nav.SelectSingleNode("@UseFontSpace"))); }
            catch { }

            try { item.UseKerning = Boolean.Parse(String.Concat(nav.SelectSingleNode("@UseKerning"))); }
            catch { }

            try { item.SymMark = Int32.Parse(String.Concat(nav.SelectSingleNode("@SymMark"))); }
            catch { }

            try { item.BorderFillId = Int32.Parse(String.Concat(nav.SelectSingleNode("@BorderFillId"))); }
            catch { }

            return item;
        }
    }

    public sealed class FontIdElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "FONTID"; } }
        public CharShapeElement Parent { get; set; }

        public string Hangul;
        public string Latin;
        public string Hanja;
        public string Japanese;
        public string Other;
        public string Symbol;
        public string User;

        public static FontIdElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            FontIdElement item = new FontIdElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Hangul = String.Concat(nav.SelectSingleNode("@Hangul"));
            item.Latin = String.Concat(nav.SelectSingleNode("@Latin"));
            item.Hanja = String.Concat(nav.SelectSingleNode("@Hanja"));
            item.Japanese = String.Concat(nav.SelectSingleNode("@Japanese"));
            item.Other = String.Concat(nav.SelectSingleNode("@Other"));
            item.Symbol = String.Concat(nav.SelectSingleNode("@Symbol"));
            item.User = String.Concat(nav.SelectSingleNode("@User"));

            return item;
        }
    }

    public sealed class RatioElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        public RatioElement()
            : base()
        {
            this.Hangul = 100d;
            this.Latin = 100d;
            this.Hanja = 100d;
            this.Japanese = 100d;
            this.Other = 100d;
            this.Symbol = 100d;
            this.User = 100d;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "RATIO"; } }
        public CharShapeElement Parent { get; set; }

        public double Hangul;
        public double Latin;
        public double Hanja;
        public double Japanese;
        public double Other;
        public double Symbol;
        public double User;

        public static RatioElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            RatioElement item = new RatioElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Hangul = Double.Parse(String.Concat(nav.SelectSingleNode("@Hangul"))); }
            catch { }

            try { item.Latin = Double.Parse(String.Concat(nav.SelectSingleNode("@Latin"))); }
            catch { }

            try { item.Hanja = Double.Parse(String.Concat(nav.SelectSingleNode("@Hanja"))); }
            catch { }

            try { item.Japanese = Double.Parse(String.Concat(nav.SelectSingleNode("@Japanese"))); }
            catch { }

            try { item.Other = Double.Parse(String.Concat(nav.SelectSingleNode("@Other"))); }
            catch { }

            try { item.Symbol = Double.Parse(String.Concat(nav.SelectSingleNode("@Symbol"))); }
            catch { }

            try { item.User = Double.Parse(String.Concat(nav.SelectSingleNode("@User"))); }
            catch { }

            return item;
        }
    }

    public sealed class CharSpacingElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "CHARSPACING"; } }
        public CharShapeElement Parent { get; set; }

        public double Hangul;
        public double Latin;
        public double Hanja;
        public double Japanese;
        public double Other;
        public double Symbol;
        public double User;

        public static CharSpacingElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            CharSpacingElement item = new CharSpacingElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Hangul = Double.Parse(String.Concat(nav.SelectSingleNode("@Hangul"))); }
            catch { }

            try { item.Latin = Double.Parse(String.Concat(nav.SelectSingleNode("@Latin"))); }
            catch { }

            try { item.Hanja = Double.Parse(String.Concat(nav.SelectSingleNode("@Hanja"))); }
            catch { }

            try { item.Japanese = Double.Parse(String.Concat(nav.SelectSingleNode("@Japanese"))); }
            catch { }

            try { item.Other = Double.Parse(String.Concat(nav.SelectSingleNode("@Other"))); }
            catch { }

            try { item.Symbol = Double.Parse(String.Concat(nav.SelectSingleNode("@Symbol"))); }
            catch { }

            try { item.User = Double.Parse(String.Concat(nav.SelectSingleNode("@User"))); }
            catch { }

            return item;
        }
    }

    public sealed class RelSizeElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        public RelSizeElement()
            : base()
        {
            this.Hangul = 100d;
            this.Latin = 100d;
            this.Hanja = 100d;
            this.Japanese = 100d;
            this.Other = 100d;
            this.Symbol = 100d;
            this.User = 100d;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "RELSIZE"; } }
        public CharShapeElement Parent { get; set; }

        public double Hangul;
        public double Latin;
        public double Hanja;
        public double Japanese;
        public double Other;
        public double Symbol;
        public double User;

        public static RelSizeElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            RelSizeElement item = new RelSizeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Hangul = Double.Parse(String.Concat(nav.SelectSingleNode("@Hangul"))); }
            catch { }

            try { item.Latin = Double.Parse(String.Concat(nav.SelectSingleNode("@Latin"))); }
            catch { }

            try { item.Hanja = Double.Parse(String.Concat(nav.SelectSingleNode("@Hanja"))); }
            catch { }

            try { item.Japanese = Double.Parse(String.Concat(nav.SelectSingleNode("@Japanese"))); }
            catch { }

            try { item.Other = Double.Parse(String.Concat(nav.SelectSingleNode("@Other"))); }
            catch { }

            try { item.Symbol = Double.Parse(String.Concat(nav.SelectSingleNode("@Symbol"))); }
            catch { }

            try { item.User = Double.Parse(String.Concat(nav.SelectSingleNode("@User"))); }
            catch { }

            return item;
        }
    }

    public sealed class CharOffsetElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "CHAROFFSET"; } }
        public CharShapeElement Parent { get; set; }

        public double Hangul;
        public double Latin;
        public double Hanja;
        public double Japanese;
        public double Other;
        public double Symbol;
        public double User;

        public static CharOffsetElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            CharOffsetElement item = new CharOffsetElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Hangul = Double.Parse(String.Concat(nav.SelectSingleNode("@Hangul"))); }
            catch { }

            try { item.Latin = Double.Parse(String.Concat(nav.SelectSingleNode("@Latin"))); }
            catch { }

            try { item.Hanja = Double.Parse(String.Concat(nav.SelectSingleNode("@Hanja"))); }
            catch { }

            try { item.Japanese = Double.Parse(String.Concat(nav.SelectSingleNode("@Japanese"))); }
            catch { }

            try { item.Other = Double.Parse(String.Concat(nav.SelectSingleNode("@Other"))); }
            catch { }

            try { item.Symbol = Double.Parse(String.Concat(nav.SelectSingleNode("@Symbol"))); }
            catch { }

            try { item.User = Double.Parse(String.Concat(nav.SelectSingleNode("@User"))); }
            catch { }

            return item;
        }
    }

    public enum UnderlineType : byte
    {
        Bottom,
        Center,
        Top
    }

    public sealed class UnderlineElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        public UnderlineElement()
            : base()
        {
            this.Type = UnderlineType.Bottom;
            this.Shape = LineType2.Solid;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "UNDERLINE"; } }
        public CharShapeElement Parent { get; set; }

        public UnderlineType Type;
        public LineType2 Shape;
        public RGBColor Color;

        public static UnderlineElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            UnderlineElement item = new UnderlineElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (UnderlineType)Enum.Parse(typeof(UnderlineType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Shape = (LineType2)Enum.Parse(typeof(LineType2), String.Concat(nav.SelectSingleNode("@Shape")), true); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public enum StrikeoutType : byte
    {
        None,
        Continuous
    }

    public sealed class StrikeoutElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        public StrikeoutElement()
            : base()
        {
            this.Type = StrikeoutType.Continuous;
            this.Shape = LineType2.Solid;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "STRIKEOUT"; } }
        public CharShapeElement Parent { get; set; }

        public StrikeoutType Type;
        public LineType2 Shape;
        public RGBColor Color;

        public static StrikeoutElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            StrikeoutElement item = new StrikeoutElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (StrikeoutType)Enum.Parse(typeof(StrikeoutType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Shape = (LineType2)Enum.Parse(typeof(LineType2), String.Concat(nav.SelectSingleNode("@Shape")), true); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class OutlineElement : IHwpmlElement, IHwpmlElement<CharShapeElement>
    {
        public OutlineElement()
            : base()
        {
            this.Type = LineType3.Solid;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharShapeElement)value; }
        }

        public string ElementName { get { return "OUTLINE"; } }
        public CharShapeElement Parent { get; set; }

        public LineType3 Type;

        public static OutlineElement Create(IXPathNavigable target, CharShapeElement parent)
        {
            if (target == null)
                return null;

            OutlineElement item = new OutlineElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (LineType3)Enum.Parse(typeof(LineType3), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            return item;
        }
    }

    public enum ShadowType : byte
    {
        Drop,
        Cont
    }

    public sealed class ShadowElement : IHwpmlElement, IHwpmlElement<CharShapeElement>, IHwpmlElement<DrawingObjectElement>, IHwpmlElement<TextArtShapeElement>
    {
        public ShadowElement()
            : base()
        {
            this.OffsetX = 10;
            this.OffsetY = 10;
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = null; return; }
                if ((this.parent = value as CharShapeElement) != null)
                    return;
                if ((this.parent = value as DrawingObjectElement) != null)
                    return;
                if ((this.parent = value as TextArtShapeElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        CharShapeElement IHwpmlElement<CharShapeElement>.Parent
        {
            get { return this.parent as CharShapeElement; }
            set { this.parent = value; }
        }

        DrawingObjectElement IHwpmlElement<DrawingObjectElement>.Parent
        {
            get { return this.parent as DrawingObjectElement; }
            set { this.parent = value; }
        }

        TextArtShapeElement IHwpmlElement<TextArtShapeElement>.Parent
        {
            get { return this.parent as TextArtShapeElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "SHADOW"; } }

        public ShadowType Type;
        public RGBColor Color;
        public int OffsetX;
        public int OffsetY;
        public string Alpha;

        public static ShadowElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ShadowElement item = new ShadowElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (ShadowType)Enum.Parse(typeof(ShadowType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            try { item.OffsetX = Int32.Parse(String.Concat(nav.SelectSingleNode("@OffsetX"))); }
            catch { }

            try { item.OffsetY = Int32.Parse(String.Concat(nav.SelectSingleNode("@OffsetY"))); }
            catch { }

            item.Alpha = String.Concat(nav.SelectSingleNode("@Alpha"));

            return item;
        }
    }

    public sealed class TabDefListElement : List<TabDefElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "TABDEFLIST"; } }
        public MappingTableElement Parent { get; set; }

        public static TabDefListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            TabDefListElement item = new TabDefListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("TABDEF"))
                item.Add(TabDefElement.Create(each, item));

            return item;
        }
    }

    public sealed class TabDefElement : IHwpmlElement, IHwpmlElement<TabDefListElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TabDefListElement)value; }
        }

        public string ElementName { get { return "TABDEF"; } }
        public TabDefListElement Parent { get; set; }

        public TabItemElement TabItem;

        public int Id;
        public bool AutoTabLeft;
        public bool AutoTabRight;

        public static TabDefElement Create(IXPathNavigable target, TabDefListElement parent)
        {
            if (target == null)
                return null;

            TabDefElement item = new TabDefElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.TabItem = TabItemElement.Create(nav.SelectSingleNode("TABITEM"), item);

            try { item.AutoTabLeft = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AutoTabLeft"))); }
            catch { }

            try { item.AutoTabRight = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AutoTabRight"))); }
            catch { }

            return item;
        }
    }

    public enum TabItemType : byte
    {
        Left,
        Right,
        Center,
        Decimal
    }

    public sealed class TabItemElement : IHwpmlElement, IHwpmlElement<TabDefElement>
    {
        public TabItemElement()
            : base()
        {
            this.Type = TabItemType.Left;
            this.Leader = LineType2.Solid;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TabDefElement)value; }
        }

        public string ElementName { get { return "TABITEM"; } }
        public TabDefElement Parent { get; set; }

        public HwpUnit Pos;
        public TabItemType Type;
        public LineType2 Leader;

        public static TabItemElement Create(IXPathNavigable target, TabDefElement parent)
        {
            if (target == null)
                return null;

            TabItemElement item = new TabItemElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Pos = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Pos")))); }
            catch { }

            try { item.Type = (TabItemType)Enum.Parse(typeof(TabItemType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }
            
            try { item.Leader = (LineType2)Enum.Parse(typeof(LineType2), String.Concat(nav.SelectSingleNode("@Leader")), true); }
            catch { }

            return item;
        }
    }

    public sealed class NumberingListElement : List<NumberingElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "NUMBERINGLIST"; } }
        public MappingTableElement Parent { get; set; }

        public static NumberingListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            NumberingListElement item = new NumberingListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("NUMBERING"))
                item.Add(NumberingElement.Create(each, item));

            return item;
        }
    }

    public sealed class NumberingElement : IHwpmlElement, IHwpmlElement<NumberingListElement>
    {
        public NumberingElement()
            : base()
        {
            this.Id = 1;
            this.Start = 1;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (NumberingListElement)value; }
        }

        public string ElementName { get { return "NUMBERING"; } }
        public NumberingListElement Parent { get; set; }

        public ParaHeadElement ParaHead;

        public int Id;
        public int Start;

        public static NumberingElement Create(IXPathNavigable target, NumberingListElement parent)
        {
            if (target == null)
                return null;

            NumberingElement item = new NumberingElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaHead = ParaHeadElement.Create(nav.SelectSingleNode("PARAHEAD"), item);

            try { item.Id = Int32.Parse(String.Concat(nav.SelectSingleNode("@Id"))); }
            catch { }

            try { item.Start = Int32.Parse(String.Concat(nav.SelectSingleNode("@Start"))); }
            catch { }

            return item;
        }
    }

    public enum ParaHeadAlignment : byte
    {
        Left,
        Center,
        Right
    }

    public enum ParaHeadTextOffsetType : byte
    {
        percent,
        hwpunit
    }

    public sealed class ParaHeadElement : IHwpmlElement, IHwpmlElement<NumberingElement>, IHwpmlElement<BulletElement>
    {
        public ParaHeadElement()
            : base()
        {
            this.Level = 1;
            this.Alignment = ParaHeadAlignment.Left;
            this.UseInstWidth = true;
            this.AutoIndent = true;
            this.TextOffsetType = ParaHeadTextOffsetType.percent;
            this.TextOffset = 50;
            this.NumFormat = NumberType1.Digit;
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as NumberingElement) != null)
                    return;
                if ((this.parent = value as BulletElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        NumberingElement IHwpmlElement<NumberingElement>.Parent
        {
            get { return this.Parent as NumberingElement; }
            set { this.Parent = value; }
        }

        BulletElement IHwpmlElement<BulletElement>.Parent
        {
            get { return this.Parent as BulletElement; }
            set { this.Parent = value; }
        }

        public string ElementName { get { return "PARAHEAD"; } }

        public string InnerText;

        public int Level;
        public ParaHeadAlignment Alignment;
        public bool UseInstWidth;
        public bool AutoIndent;
        public HwpUnit WidthAdjust;
        public ParaHeadTextOffsetType TextOffsetType;
        public int TextOffset;
        public NumberType1 NumFormat;
        public int CharShape;

        public Type GetTextOffsetValueType()
        {
            if (this.TextOffsetType == ParaHeadTextOffsetType.percent)
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetTextOffsetValueAsPercentage(ref int result)
        {
            if (!this.GetTextOffsetValueType().Equals(typeof(int)))
                return false;

            result = this.TextOffset;
            return true;
        }

        public bool TryGetTextOffsetValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetTextOffsetValueType().Equals(typeof(HwpUnit)))
                return false;

            result = new HwpUnit(this.TextOffset);
            return true;
        }

        public static ParaHeadElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ParaHeadElement item = new ParaHeadElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.InnerText = String.Concat(target);

            try { item.Level = Int32.Parse(String.Concat(nav.SelectSingleNode("@Level"))); }
            catch { }

            try { item.Alignment = (ParaHeadAlignment)Enum.Parse(typeof(ParaHeadAlignment), String.Concat(nav.SelectSingleNode("@Alignment")), true); }
            catch { }

            try { item.UseInstWidth = Boolean.Parse(String.Concat(nav.SelectSingleNode("@UseInstWidth"))); }
            catch { }

            try { item.AutoIndent = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AutoIndent"))); }
            catch { }

            try { item.WidthAdjust = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@WidthAdjust")))); }
            catch { }

            try { item.TextOffsetType = (ParaHeadTextOffsetType)Enum.Parse(typeof(ParaHeadTextOffsetType), String.Concat(nav.SelectSingleNode("@TextOffsetType")), true); }
            catch { }

            try { item.TextOffset = Int32.Parse(String.Concat(nav.SelectSingleNode("@TextOffset"))); }
            catch { }

            try { item.NumFormat = (NumberType1)Enum.Parse(typeof(NumberType1), String.Concat(nav.SelectSingleNode("@NumFormat")), true); }
            catch { }

            try { item.CharShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@CharShape"))); }
            catch { }

            return item;
        }
    }

    public sealed class BulletListElement : List<BulletElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "BULLETLIST"; } }
        public MappingTableElement Parent { get; set; }

        public static BulletListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            BulletListElement item = new BulletListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("BULLET"))
                item.Add(BulletElement.Create(each, item));

            return item;
        }
    }

    public sealed class BulletElement : IHwpmlElement, IHwpmlElement<BulletListElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BulletListElement)value; }
        }

        public string ElementName { get { return "BULLET"; } }
        public BulletListElement Parent { get; set; }

        public ParaHeadElement ParaHead;

        public int Id;
        public string Char;
        public bool Image;

        public static BulletElement Create(IXPathNavigable target, BulletListElement parent)
        {
            if (target == null)
                return null;

            BulletElement item = new BulletElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaHead = ParaHeadElement.Create(nav.SelectSingleNode("PARAHEAD"), item);

            try { item.Id = Int32.Parse(String.Concat(nav.SelectSingleNode("@Id"))); }
            catch { }

            item.Char = String.Concat(nav.SelectSingleNode("@Char"));

            try { item.Image = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Image"))); }
            catch { }

            return item;
        }
    }

    public sealed class ParaShapeListElement : List<ParaShapeElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "PARASHAPELIST"; } }
        public MappingTableElement Parent { get; set; }

        public static ParaShapeListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            ParaShapeListElement item = new ParaShapeListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("PARASHAPE"))
                item.Add(ParaShapeElement.Create(each, item));

            return item;
        }
    }

    public enum ParaShapeVerAlign : byte
    {
        Baseline,
        Top,
        Center,
        Bottom
    }

    public enum ParaShapeHeadingType : byte
    {
        None,
        Outline,
        Number,
        Bullet
    }

    public enum ParaShapeBreakLatinWord : byte
    {
        KeepWord,
        Hyphenation,
        BreakWord
    }

    public sealed class ParaShapeElement : IHwpmlElement, IHwpmlElement<ParaShapeListElement>
    {
        public ParaShapeElement()
            : base()
        {
            this.Align = AlignmentType1.Justify;
            this.VerAlign = ParaShapeVerAlign.Baseline;
            this.HeadingType = ParaShapeHeadingType.None;
            this.BreakLatinWord = ParaShapeBreakLatinWord.KeepWord;
            this.BreakNonLatinWord = true;
            this.SnapToGrid = true;
            this.LineWrap = LineWrapType.Break;
            this.AutoSpaceEAsianEng = true;
            this.AutoSpaceEAsianNum = true;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ParaShapeListElement)value; }
        }

        public string ElementName { get { return "PARASHAPE"; } }
        public ParaShapeListElement Parent { get; set; }

        public ParaMarginElement ParaMargin;
        public ParaBorderElement ParaBorder;

        public int Id;
        public AlignmentType1 Align;
        public ParaShapeVerAlign VerAlign;
        public ParaShapeHeadingType HeadingType;
        public int Heading;
        public int Level;
        public int TabDef;
        public ParaShapeBreakLatinWord BreakLatinWord;
        public bool BreakNonLatinWord;
        public double Condense;
        public bool WindowOrphan;
        public bool KeepWithNext;
        public bool KeepLines;
        public bool PageBreakBefore;
        public bool FontLineHeight;
        public bool SnapToGrid;
        public LineWrapType LineWrap;
        public bool AutoSpaceEAsianEng;
        public bool AutoSpaceEAsianNum;

        public static ParaShapeElement Create(IXPathNavigable target, ParaShapeListElement parent)
        {
            if (target == null)
                return null;

            ParaShapeElement item = new ParaShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaMargin = ParaMarginElement.Create(nav.SelectSingleNode("PARAMARGIN"), item);
            item.ParaBorder = ParaBorderElement.Create(nav.SelectSingleNode("PARABORDER"), item);

            try { item.Id = Int32.Parse(String.Concat(nav.SelectSingleNode("@Id"))); }
            catch { }

            try { item.Align = (AlignmentType1)Enum.Parse(typeof(AlignmentType1), String.Concat(nav.SelectSingleNode("@Align")), true); }
            catch { }

            try { item.VerAlign = (ParaShapeVerAlign)Enum.Parse(typeof(ParaShapeVerAlign), String.Concat(nav.SelectSingleNode("@VerAlign")), true); }
            catch { }

            try { item.HeadingType = (ParaShapeHeadingType)Enum.Parse(typeof(ParaShapeHeadingType), String.Concat(nav.SelectSingleNode("@HeadingType")), true); }
            catch { }

            try { item.Heading = Int32.Parse(String.Concat(nav.SelectSingleNode("@Heading"))); }
            catch { }

            try { item.Level = Int32.Parse(String.Concat(nav.SelectSingleNode("@Level"))); }
            catch { }

            try { item.TabDef = Int32.Parse(String.Concat(nav.SelectSingleNode("@TabDef"))); }
            catch { }

            try { item.BreakLatinWord = (ParaShapeBreakLatinWord)Enum.Parse(typeof(ParaShapeBreakLatinWord), String.Concat(nav.SelectSingleNode("@BreakLatinWord")), true); }
            catch { }

            try { item.Condense = Double.Parse(String.Concat(nav.SelectSingleNode("@Condense"))); }
            catch { }

            try { item.WindowOrphan = Boolean.Parse(String.Concat(nav.SelectSingleNode("@WindowOrphan"))); }
            catch { }

            try { item.KeepWithNext = Boolean.Parse(String.Concat(nav.SelectSingleNode("@KeepWithNext"))); }
            catch { }

            try { item.KeepLines = Boolean.Parse(String.Concat(nav.SelectSingleNode("@KeepLines"))); }
            catch { }

            try { item.PageBreakBefore = Boolean.Parse(String.Concat(nav.SelectSingleNode("@PageBreakBefore"))); }
            catch { }

            try { item.FontLineHeight = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FontLineHeight"))); }
            catch { }

            try { item.SnapToGrid = Boolean.Parse(String.Concat(nav.SelectSingleNode("@SnapToGrid"))); }
            catch { }

            try { item.LineWrap = (LineWrapType)Enum.Parse(typeof(LineWrapType), String.Concat(nav.SelectSingleNode("@LineWrap")), true); }
            catch { }

            try { item.AutoSpaceEAsianEng = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AutoSpaceEAsianEng"))); }
            catch { }

            try { item.AutoSpaceEAsianNum = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AutoSpaceEAsianNum"))); }
            catch { }

            return item;
        }
    }

    public enum ParaMarginLineSpacingType : byte
    {
        Percent,
        Fixed,
        BetweenLines,
        AtLeast
    }

    public sealed class ParaMarginElement : IHwpmlElement, IHwpmlElement<ParaShapeElement>
    {
        public ParaMarginElement()
            : base()
        {
            this.LineSpacingType = ParaMarginLineSpacingType.Percent;
            this.LineSpacing = "160";
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ParaShapeElement)value; }
        }

        public string ElementName { get { return "PARAMARGIN"; } }
        public ParaShapeElement Parent { get; set; }

        public string Indent;
        public string Left;
        public string Right;
        public string Prev;
        public string Next;
        public ParaMarginLineSpacingType LineSpacingType;
        public string LineSpacing;

        public Type GetIndentValueType()
        {
            if (this.Indent.EndsWith("ch", StringComparison.OrdinalIgnoreCase))
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetIndentValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetIndentValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.Indent));
                return true;
            }
            catch { return false; }
        }

        public bool TryGetIndentValueAsCharCount(ref int result)
        {
            if (!this.GetIndentValueType().Equals(typeof(int)))
                return false;

            try
            {
                result = Int32.Parse(this.Indent);
                return true;
            }
            catch { return false; }
        }

        public bool IndicateIsIndent()
        {
            try { return Int32.Parse(this.Indent) > 0; }
            catch { return false; }
        }

        public bool IndicateIsNormal()
        {
            try { return Int32.Parse(this.Indent) == 0; }
            catch { return false; }
        }

        public bool IndicateIsOutdent()
        {
            try { return Int32.Parse(this.Indent) < 0; }
            catch { return false; }
        }

        public Type GetLeftValueType()
        {
            if (this.Left.EndsWith("ch", StringComparison.OrdinalIgnoreCase))
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetLeftValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetLeftValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.Left));
                return true;
            }
            catch { return false; }
        }

        public bool TryGetLeftValueAsCharCount(ref int result)
        {
            if (!this.GetLeftValueType().Equals(typeof(int)))
                return false;

            try
            {
                result = Int32.Parse(this.Left);
                return true;
            }
            catch { return false; }
        }

        public Type GetRightValueType()
        {
            if (this.Right.EndsWith("ch", StringComparison.OrdinalIgnoreCase))
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetRightValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetRightValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.Right));
                return true;
            }
            catch { return false; }
        }

        public bool TryGetRightValueAsCharCount(ref int result)
        {
            if (!this.GetRightValueType().Equals(typeof(int)))
                return false;

            try
            {
                result = Int32.Parse(this.Right);
                return true;
            }
            catch { return false; }
        }

        public Type GetPrevValueType()
        {
            if (this.Prev.EndsWith("ch", StringComparison.OrdinalIgnoreCase))
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetPrevValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetPrevValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.Prev));
                return true;
            }
            catch { return false; }
        }

        public bool TryGetPrevValueAsCharCount(ref int result)
        {
            if (!this.GetPrevValueType().Equals(typeof(int)))
                return false;

            try
            {
                result = Int32.Parse(this.Prev);
                return true;
            }
            catch { return false; }
        }

        public Type GetNextValueType()
        {
            if (this.Next.EndsWith("ch", StringComparison.OrdinalIgnoreCase))
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetNextValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetNextValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.Next));
                return true;
            }
            catch { return false; }
        }

        public bool TryGetNextValueAsCharCount(ref int result)
        {
            if (!this.GetNextValueType().Equals(typeof(int)))
                return false;

            try
            {
                result = Int32.Parse(this.Next);
                return true;
            }
            catch { return false; }
        }

        public Type GetLineSpacingValueType()
        {
            if (this.LineSpacingType.Equals(ParaMarginLineSpacingType.Percent))
                return typeof(double);

            if (this.LineSpacing.EndsWith("ch", StringComparison.OrdinalIgnoreCase))
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetLineSpacingValueAsPercentage(ref double result)
        {
            if (!this.GetLineSpacingValueType().Equals(typeof(double)))
                return false;

            try
            {
                result = Double.Parse(this.LineSpacing);
                return true;
            }
            catch { return false; }
        }

        public bool TryGetLineSpacingValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetLineSpacingValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.LineSpacing));
                return true;
            }
            catch { return false; }
        }

        public bool TryGetLineSpacingValueAsCharCount(ref int result)
        {
            if (!this.GetLineSpacingValueType().Equals(typeof(int)))
                return false;

            try
            {
                result = Int32.Parse(this.LineSpacing);
                return true;
            }
            catch { return false; }
        }

        public static ParaMarginElement Create(IXPathNavigable target, ParaShapeElement parent)
        {
            if (target == null)
                return null;

            ParaMarginElement item = new ParaMarginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Indent = String.Concat(nav.SelectSingleNode("@Indent"));
            item.Left = String.Concat(nav.SelectSingleNode("@Left"));
            item.Right = String.Concat(nav.SelectSingleNode("@Right"));
            item.Prev = String.Concat(nav.SelectSingleNode("@Prev"));
            item.Next = String.Concat(nav.SelectSingleNode("@Next"));

            try { item.LineSpacingType = (ParaMarginLineSpacingType)Enum.Parse(typeof(ParaMarginLineSpacingType), String.Concat(nav.SelectSingleNode("@LineSpacingType")), true); }
            catch { }

            item.LineSpacing = String.Concat(nav.SelectSingleNode("@LineSpacing"));

            return item;
        }
    }

    public sealed class ParaBorderElement : IHwpmlElement, IHwpmlElement<ParaShapeElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ParaShapeElement)value; }
        }

        public string ElementName { get { return "PARABORDER"; } }
        public ParaShapeElement Parent { get; set; }

        public int BorderFill;
        public HwpUnit OffsetLeft;
        public HwpUnit OffsetRight;
        public HwpUnit OffsetTop;
        public HwpUnit OffsetBottom;
        public bool Connect;
        public bool IgnoreMargin;

        public static ParaBorderElement Create(IXPathNavigable target, ParaShapeElement parent)
        {
            if (target == null)
                return null;

            ParaBorderElement item = new ParaBorderElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.BorderFill = Int32.Parse(String.Concat(nav.SelectSingleNode("@BorderFill"))); }
            catch { }

            try { item.OffsetLeft = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@OffsetLeft")))); }
            catch { }

            try { item.OffsetRight = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@OffsetRight")))); }
            catch { }

            try { item.OffsetTop = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@OffsetTop")))); }
            catch { }

            try { item.OffsetBottom = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@OffsetBottom")))); }
            catch { }

            try { item.Connect = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Connect"))); }
            catch { }

            try { item.IgnoreMargin = Boolean.Parse(String.Concat(nav.SelectSingleNode("@IgnoreMargin"))); }
            catch { }

            return item;
        }
    }

    public sealed class StyleListElement : List<StyleElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "STYLELIST"; } }
        public MappingTableElement Parent { get; set; }

        public static StyleListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            StyleListElement item = new StyleListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("STYLE"))
                item.Add(StyleElement.Create(each, item));

            return item;
        }
    }

    public enum StyleType : byte
    {
        Para,
        Char
    }

    public sealed class StyleElement : IHwpmlElement, IHwpmlElement<StyleListElement>
    {
        public StyleElement()
            : base()
        {
            this.Type = StyleType.Para;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (StyleListElement)value; }
        }

        public string ElementName { get { return "STYLE"; } }
        public StyleListElement Parent { get; set; }

        public string Id;
        public StyleType Type;
        public string Name;
        public string EngName;
        public int ParaShape;
        public int CharShape;
        public string NextStyle;
        public string LangId;
        public string LockForm;

        public static StyleElement Create(IXPathNavigable target, StyleListElement parent)
        {
            if (target == null)
                return null;

            StyleElement item = new StyleElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Id = String.Concat(nav.SelectSingleNode("@Id"));

            try { item.Type = (StyleType)Enum.Parse(typeof(StyleType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));
            item.EngName = String.Concat(nav.SelectSingleNode("@EngName"));
            
            try { item.ParaShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@ParaShape"))); }
            catch { }

            try { item.CharShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@CharShape"))); }
            catch { }

            item.NextStyle = String.Concat(nav.SelectSingleNode("@NextStyle"));
            item.LangId = String.Concat(nav.SelectSingleNode("@LangId"));
            item.LockForm = String.Concat(nav.SelectSingleNode("@LockForm"));

            return item;
        }
    }

    public sealed class MemoShapeListElement : List<MemoElement>, IHwpmlElement, IHwpmlElement<MappingTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MappingTableElement)value; }
        }

        public string ElementName { get { return "MEMOSHAPELIST"; } }
        public MappingTableElement Parent { get; set; }

        public static MemoShapeListElement Create(IXPathNavigable target, MappingTableElement parent)
        {
            if (target == null)
                return null;

            MemoShapeListElement item = new MemoShapeListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("@MEMO"))
                item.Add(MemoElement.Create(each, item));

            return item;
        }
    }

    public sealed class MemoElement : IHwpmlElement, IHwpmlElement<MemoShapeListElement>
    {
        public MemoElement() : base()
        {
            this.LineType = LineType1.None;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (MemoShapeListElement)value; }
        }

        public string ElementName { get { return "MEMO"; } }
        public MemoShapeListElement Parent { get; set; }

        public string Id;
        public int Width;
        public LineType1 LineType;
        public RGBColor LineColor;
        public RGBColor FillColor;
        public RGBColor ActiveColor;
        public string MemoType;

        public static MemoElement Create(IXPathNavigable target, MemoShapeListElement parent)
        {
            if (target == null)
                return null;

            MemoElement item = new MemoElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Id = String.Concat(nav.SelectSingleNode("@Id"));
            
            try { item.Width = Int32.Parse(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.LineType = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@LineType")), true); }
            catch { }

            try { item.LineColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@LineColor")))); }
            catch { }

            try { item.FillColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@FillColor")))); }
            catch { }

            try { item.ActiveColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@ActiveColor")))); }
            catch { }

            item.MemoType = String.Concat(nav.SelectSingleNode("@MemoType"));

            return item;
        }
    }

    public sealed class BodyElement : List<SectionElement>, IHwpmlElement, IHwpmlElement<HwpmlElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HwpmlElement)value; }
        }

        public string ElementName { get { return "BODY"; } }
        public HwpmlElement Parent { get; set; }

        public static BodyElement Create(IXPathNavigable target, HwpmlElement parent)
        {
            if (target == null)
                return null;

            BodyElement item = new BodyElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("SECTION"))
                item.Add(SectionElement.Create(each, item));

            return item;
        }
    }

    public sealed class SectionElement : List<PElement>, IHwpmlElement, IHwpmlElement<BodyElement>
    {
        public SectionElement()
            : base()
        {
            this.Text = new List<string>();
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BodyElement)value; }
        }

        public string ElementName { get { return "SECTION"; } }
        public BodyElement Parent { get; set; }

        public IList<string> Text;

        public static SectionElement Create(IXPathNavigable target, BodyElement parent)
        {
            if (target == null)
                return null;

            SectionElement item = new SectionElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.SelectChildren(XPathNodeType.All))
            {
                if (each.CreateNavigator().Name.Equals("P"))
                    item.Add(PElement.Create(each, item));
                else
                    item.Text.Add(each.ToString());
            }

            return item;
        }
    }

    public sealed class PElement : List<TextElement>, IHwpmlElement, IHwpmlElement<SectionElement>, IHwpmlElement<ParaListElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as SectionElement) != null)
                    return;
                if ((this.parent = value as ParaListElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        SectionElement IHwpmlElement<SectionElement>.Parent
        {
            get { return this.parent as SectionElement; }
            set { this.parent = value; }
        }

        ParaListElement IHwpmlElement<ParaListElement>.Parent
        {
            get { return this.parent as ParaListElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "P"; } }

        public int ParaShape;
        public string Style;
        public string InstId;
        public bool PageBreak;
        public bool ColumnBreak;

        private static PElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            PElement item = new PElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("TEXT"))
                item.Add(TextElement.Create(each, item));

            try { item.ParaShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@ParaShape"))); }
            catch { }

            item.Style = String.Concat(nav.SelectSingleNode("@Style"));
            item.InstId = String.Concat(nav.SelectSingleNode("@InstId"));
            
            try { item.PageBreak = Boolean.Parse(String.Concat(nav.SelectSingleNode("@PageBreak"))); }
            catch { }

            try { item.ColumnBreak = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ColumnBreak"))); }
            catch { }

            return item;
        }

        public static PElement Create(IXPathNavigable target, SectionElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static PElement Create(IXPathNavigable target, ParaListElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public sealed class TextElement : List<object>, IHwpmlElement, IHwpmlElement<PElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (PElement)value; }
        }

        public string ElementName { get { return "TEXT"; } }
        public PElement Parent { get; set; }

        public int CharShape;

        public static TextElement Create(IXPathNavigable target, PElement parent)
        {
            if (target == null)
                return null;

            TextElement item = new TextElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.SelectChildren(XPathNodeType.Element))
                item.Add(CreateElement(each, item));

            try { item.CharShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@CharShape"))); }
            catch { }

            return item;
        }

        public static IHwpmlElement CreateElement(IXPathNavigable target, TextElement parent)
        {
            XPathNavigator nav = target.CreateNavigator();
            switch (nav.Name.ToUpperInvariant().Trim())
            {
                case "CHAR":
                    return CharElement.Create(target, parent);
                case "SECDEF":
                    return SecDefElement.Create(target, parent);
                case "COLDEF":
                    return ColDefElement.Create(target, parent);
                case "TABLE":
                    return TableElement.Create(target, parent);
                case "PICTURE":
                    return PictureElement.Create(target, parent);
                case "CONTAINER":
                    return ContainerElement.Create(target, parent);
                case "OLE":
                    return OleElement.Create(target, parent);
                case "EQUATION":
                    return EquationElement.Create(target, parent);
                case "TEXTART":
                    return TextArtElement.Create(target, parent);
                case "LINE":
                    return LineElement.Create(target, parent);
                case "RECTANGLE":
                    return RectangleElement.Create(target, parent);
                case "ELLIPSE":
                    return EllipseElement.Create(target, parent);
                case "ARC":
                    return ArcElement.Create(target, parent);
                case "POLYGON":
                    return PolygonElement.Create(target, parent);
                case "CURVE":
                    return CurveElement.Create(target, parent);
                case "CONNECTLINE":
                    return ConnectLineElement.Create(target, parent);
                case "UNKNOWNOBJECT":
                    return UnknownObjectElement.Create(target, parent);
                case "FIELDBEGIN":
                    return FieldBeginElement.Create(target, parent);
                case "FIELDEND":
                    return FieldEndElement.Create(target, parent);
                case "BOOKMARK":
                    return BookmarkElement.Create(target, parent);
                case "HEADER":
                    return HeaderElement.Create(target, parent);
                case "FOOTER":
                    return FooterElement.Create(target, parent);
                case "FOOTNOTE":
                    return FootnoteElement.Create(target, parent);
                case "ENDNOTE":
                    return EndnoteElement.Create(target, parent);
                case "AUTONUM":
                    return AutoNumElement.Create(target, parent);
                case "NEWNUM":
                    return NewNumElement.Create(target, parent);
                case "PAGENUMCTRL":
                    return PageNumCtrlElement.Create(target, parent);
                case "PAGEHIDING":
                    return PageHidingElement.Create(target, parent);
                case "PAGENUM":
                    return PageNumElement.Create(target, parent);
                case "INDEXMARK":
                    return IndexMarkElement.Create(target, parent);
                case "COMPOSE":
                    return ComposeElement.Create(target, parent);
                case "DUTMAL":
                    return DutmalElement.Create(target, parent);
                case "HIDDENCOMMENT":
                    return HiddenCommentElement.Create(target, parent);
                case "BUTTON":
                    return ButtonElement.Create(target, parent);
                case "RADIOBUTTON":
                    return RadioButtonElement.Create(target, parent);
                case "CHECKBUTTON":
                    return CheckButtonElement.Create(target, parent);
                case "COMBOBOX":
                    return ComboBoxElement.Create(target, parent);
                case "EDIT":
                    return EditElement.Create(target, parent);
                case "LISTBOX":
                    return ListBoxElement.Create(target, parent);
                case "SCROLLBAR":
                    return ScrollBarElement.Create(target, parent);
            }

            return GenericTextElement.Create(target, parent);
        }
    }

    public sealed class CharElement : List<object>, IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "CHAR"; } }
        public TextElement Parent { get; set; }

        public string Style;

        public override string ToString()
        {
            StringBuilder buffer = new StringBuilder();

            foreach (object each in this)
                buffer.Append(each.ToString());

            return buffer.ToString();
        }

        public static CharElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            CharElement item = new CharElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.SelectChildren(XPathNodeType.All))
                item.Add(CreateElement(each, item));

            item.Style = String.Concat(nav.SelectSingleNode("@Style"));

            return item;
        }

        public static object CreateElement(IXPathNavigable target, CharElement parent)
        {
            XPathNavigator nav = target.CreateNavigator();

            switch (nav.Name.Trim().ToUpperInvariant())
            {
                case "TAB":
                    return TabElement.Create(target, parent);
                case "LINEBREAK":
                    return LineBreakElement.Create(target, parent);
                case "HYPEN":
                    return HypenElement.Create(target, parent);
                case "NBSPACE":
                    return NbSpaceElement.Create(target, parent);
                case "FWSPACE":
                    return FwSpaceElement.Create(target, parent);
                case "TITLEMARK":
                    return TitleMarkElement.Create(target, parent);
                case "MARKPENBEGIN":
                    return MarkPenBeginElement.Create(target, parent);
                case "MARKPENEND":
                    return MarkPenEndElement.Create(target, parent);
            }

            return GenericCharacterElement.Create(target, parent);
        }
    }

    public sealed class MarkPenBeginElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "MARKPENBEGIN"; } }
        public CharElement Parent { get; set; }

        public RGBColor Color;

        public override string ToString()
        {
            return String.Empty;
        }

        public static MarkPenBeginElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            MarkPenBeginElement item = new MarkPenBeginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class MarkPenEndElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "MARKPENEND"; } }
        public CharElement Parent { get; set; }

        public override string ToString()
        {
            return String.Empty;
        }

        public static MarkPenEndElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            MarkPenEndElement item = new MarkPenEndElement();
            item.Parent = parent;
            return item;
        }
    }

    public sealed class TitleMarkElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "TITLEMARK"; } }
        public CharElement Parent { get; set; }

        public bool Ignore;

        public override string ToString()
        {
            return String.Empty;
        }

        public static TitleMarkElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            TitleMarkElement item = new TitleMarkElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Ignore = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Ignore"))); }
            catch { }

            return item;
        }
    }

    public sealed class TabElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }
        
        public string ElementName { get { return "TAB"; } }
        public CharElement Parent { get; set; }

        public override string ToString()
        {
            return "\t";
        }

        public static TabElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            TabElement item = new TabElement();
            item.Parent = parent;
            return item;
        }
    }

    public sealed class LineBreakElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "LINEBREAK"; } }
        public CharElement Parent { get; set; }

        public override string ToString()
        {
            return Environment.NewLine;
        }

        public static LineBreakElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            LineBreakElement item = new LineBreakElement();
            item.Parent = parent;
            return item;
        }
    }

    public sealed class HypenElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "HYPEN"; } }
        public CharElement Parent { get; set; }

        public override string ToString()
        {
            return '\u2015'.ToString();
        }

        public static HypenElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            HypenElement item = new HypenElement();
            item.Parent = parent;
            return item;
        }
    }

    public sealed class NbSpaceElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "NBSPACE"; } }
        public CharElement Parent { get; set; }

        public override string ToString()
        {
            return '\u3000'.ToString();
        }

        public static NbSpaceElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            NbSpaceElement item = new NbSpaceElement();
            item.Parent = parent;
            return item;
        }
    }

    public sealed class FwSpaceElement : IHwpmlElement, IHwpmlElement<CharElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "FWSPACE"; } }
        public CharElement Parent { get; set; }

        public override string ToString()
        {
            return '\u0020'.ToString();
        }

        public static FwSpaceElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            FwSpaceElement item = new FwSpaceElement();
            item.Parent = parent;
            return item;
        }
    }

    public sealed class GenericCharacterElement : IHwpmlElement, IHwpmlElement<CharElement>, ITextElement
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CharElement)value; }
        }

        public string ElementName { get { return "#GenericCharacterElement"; } }
        public CharElement Parent { get; set; }

        public string Content { get; set; }

        public override string ToString()
        {
            return this.Content;
        }

        public static GenericCharacterElement Create(IXPathNavigable target, CharElement parent)
        {
            if (target == null)
                return null;

            GenericCharacterElement item = new GenericCharacterElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Content = (nav.NodeType == XPathNodeType.Text ? nav.ToString() : nav.InnerXml);

            return item;
        }
    }

    public sealed class GenericTextElement : IHwpmlElement, IHwpmlElement<TextElement>, ITextElement
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "#GenericText"; } }
        public TextElement Parent { get; set; }

        public string Content { get; set; }

        public override string ToString()
        {
            return this.Content;
        }

        public static GenericTextElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            GenericTextElement item = new GenericTextElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Content = (nav.NodeType == XPathNodeType.Text ? nav.ToString() : nav.InnerXml);

            return item;
        }
    }

    public sealed class SecDefElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public SecDefElement()
            : base()
        {
            this.ExtMasterPages = new List<ExtMasterPageElement>();
            this.TabStop = "8000";
            this.OutlineShape = 1;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "SECDEF"; } }
        public TextElement Parent { get; set; }

        public ParameterSetElement ParameterSet;
        public StartNumberElement StartNumber;
        public HideElement Hide;
        public PageDefElement PageDef;
        public FootnoteShapeElement FootnoteShape;
        public EndnoteShapeElement EndnoteShape;
        public PageBorderFillElement PageBorderFill;
        public MasterPageElement MasterPage;
        public IList<ExtMasterPageElement> ExtMasterPages;

        public int TextDirection;
        public HwpUnit SpaceColumns;
        public string TabStop;
        public int OutlineShape;
        public HwpUnit LineGrid;
        public HwpUnit CharGrid;
        public bool FirstBorder;
        public bool FirstFill;
        public int ExtMasterpageCount;

        public string MemoShapeId;
        public int TextVerticalWidthHead;

        public Type GetTabStopValueType()
        {
            if (this.TabStop.EndsWith("ch", StringComparison.OrdinalIgnoreCase))
                return typeof(int);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetTabStopValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetTabStopValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.TabStop));
                return true;
            }
            catch { return false; }
        }

        public bool TryGetTabStopValueAsCharCount(ref int result)
        {
            if (!this.GetTabStopValueType().Equals(typeof(int)))
                return false;

            try
            {
                result = Int32.Parse(this.TabStop);
                return true;
            }
            catch { return false; }
        }

        public bool IndicateUseLineGrid()
        { return this.LineGrid.Value.Equals(0); }

        public bool IndicateUseCharGrid()
        { return this.CharGrid.Value.Equals(0); }

        public static SecDefElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            SecDefElement item = new SecDefElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParameterSet = ParameterSetElement.Create(nav.SelectSingleNode("PARAMETERSET"), item);
            item.StartNumber = StartNumberElement.Create(nav.SelectSingleNode("STARTNUMBER"), item);
            item.Hide = HideElement.Create(nav.SelectSingleNode("HIDE"), item);
            item.PageDef = PageDefElement.Create(nav.SelectSingleNode("PAGEDEF"), item);
            item.FootnoteShape = FootnoteShapeElement.Create(nav.SelectSingleNode("FOOTNOTESHAPE"), item);
            item.EndnoteShape = EndnoteShapeElement.Create(nav.SelectSingleNode("ENDNOTESHAPE"), item);
            item.PageBorderFill = PageBorderFillElement.Create(nav.SelectSingleNode("PAGEBORDERFILL"), item);
            item.MasterPage = MasterPageElement.Create(nav.SelectSingleNode("MASTERPAGE"), item);

            foreach (IXPathNavigable each in nav.Select("EXT_MASTERPAGE"))
                item.ExtMasterPages.Add(ExtMasterPageElement.Create(each, item));

            try { item.TextDirection = Int32.Parse(String.Concat(nav.SelectSingleNode("@TextDirection"))); }
            catch { }

            try { item.SpaceColumns = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@SpaceColumns")))); }
            catch { }

            item.TabStop = String.Concat(nav.SelectSingleNode("@TabStop"));

            try { item.OutlineShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@OutlineShape"))); }
            catch { }

            try { item.LineGrid = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@LineGrid")))); }
            catch { }

            try { item.CharGrid = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@CharGrid")))); }
            catch { }

            try { item.FirstBorder = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FirstBorder"))); }
            catch { }

            try { item.FirstFill = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FirstFill"))); }
            catch { }

            try { item.ExtMasterpageCount = Int32.Parse(String.Concat(nav.SelectSingleNode("@ExtMasterpageCount"))); }
            catch { }

            item.MemoShapeId = String.Concat(nav.SelectSingleNode("@MemoShapeId"));

            try { item.TextVerticalWidthHead = Int32.Parse(String.Concat(nav.SelectSingleNode("@TextVerticalWidthHead"))); }
            catch { }

            return item;
        }
    }

    public sealed class ParameterSetElement : List<ItemElement>, IHwpmlElement, IHwpmlElement<SecDefElement>, IHwpmlElement<ItemElement>, IHwpmlElement<ColDefElement>, IHwpmlElement<ShapeComponentElement>, IHwpmlElement<FormObjectElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as SecDefElement) != null)
                    return;
                if ((this.parent = value as ItemElement) != null)
                    return;
                if ((this.parent = value as ColDefElement) != null)
                    return;
                if ((this.parent = value as ShapeComponentElement) != null)
                    return;
                if ((this.parent = value as FormObjectElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        SecDefElement IHwpmlElement<SecDefElement>.Parent
        {
            get { return this.parent as SecDefElement; }
            set { this.parent = value; }
        }

        ItemElement IHwpmlElement<ItemElement>.Parent
        {
            get { return this.parent as ItemElement; }
            set { this.parent = value; }
        }

        ColDefElement IHwpmlElement<ColDefElement>.Parent
        {
            get { return this.parent as ColDefElement; }
            set { this.parent = value; }
        }

        ShapeComponentElement IHwpmlElement<ShapeComponentElement>.Parent
        {
            get { return this.parent as ShapeComponentElement; }
            set { this.parent = value; }
        }

        FormObjectElement IHwpmlElement<FormObjectElement>.Parent
        {
            get { return this.parent as FormObjectElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "PARAMETERSET"; } }

        public string SetId;

        public static ParameterSetElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ParameterSetElement item = new ParameterSetElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("ITEM"))
                item.Add(ItemElement.Create(each, item));

            item.SetId = String.Concat(nav.SelectSingleNode("@SetId"));

            return item;
        }
    }

    public sealed class ParameterArrayElement : List<ItemElement>, IHwpmlElement, IHwpmlElement<ItemElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ItemElement)value; }
        }

        public string ElementName { get { return "PARAMETERARRAY"; } }
        public ItemElement Parent { get; set; }

        public static ParameterArrayElement Create(IXPathNavigable target, ItemElement parent)
        {
            if (target == null)
                return null;

            ParameterArrayElement item = new ParameterArrayElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("ITEM"))
                item.Add(ItemElement.Create(each, item));

            return item;
        }
    }

    public enum ItemType : byte
    {
        Bstr,
        Integer,
        Set,
        Array,
        BinData
    }

    public sealed class ItemElement : List<object>, IHwpmlElement, IHwpmlElement<ParameterSetElement>, IHwpmlElement<ParameterArrayElement>, ITextElement
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as ParameterSetElement) != null)
                    return;
                if ((this.parent = value as ParameterArrayElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        ParameterSetElement IHwpmlElement<ParameterSetElement>.Parent
        {
            get { return this.parent as ParameterSetElement; }
            set { this.parent = value; }
        }

        ParameterArrayElement IHwpmlElement<ParameterArrayElement>.Parent
        {
            get { return this.parent as ParameterArrayElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "ITEM"; } }

        public string Content { get; set; }

        public string ItemId;
        public ItemType Type;

        public int ToInt32()
        {
            if (this.Type.Equals(ItemType.Integer))
                return Int32.Parse(this.Content);
            else
                return default(int);
        }

        public byte[] ToByteArray()
        {
            if (this.Type.Equals(ItemType.BinData))
                return Convert.FromBase64String(this.Content);
            else
                return new byte[] { };
        }

        public IEnumerable<ParameterSetElement> ToParameterSet()
        {
            if (this.Type.Equals(ItemType.Set))
                yield break;

            foreach (object o in this)
                if (o is ParameterSetElement)
                    yield return o as ParameterSetElement;
        }

        public IEnumerable<ParameterArrayElement> ToParameterArray()
        {
            if (this.Type.Equals(ItemType.Array))
                yield break;

            foreach (object o in this)
                if (o is ParameterArrayElement)
                    yield return o as ParameterArrayElement;
        }

        public override string ToString()
        {
            if (this.Type.Equals(ItemType.Bstr))
                return this.Content;
            else
                return String.Empty;
        }

        public static ItemElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ItemElement item = new ItemElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("PARAMETERSET"))
                item.Add(ParameterSetElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("PARAMETERARRAY"))
                item.Add(ParameterArrayElement.Create(each, item));

            foreach (IXPathNavigable each in nav.SelectChildren(XPathNodeType.All))
                if (each.CreateNavigator().NodeType == XPathNodeType.Text)
                    item.Add(ItemElement.Create(each, item));

            item.ItemId = String.Concat(nav.SelectSingleNode("@ItemId"));

            try { item.Type = (ItemType)Enum.Parse(typeof(ItemType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            return item;
        }
    }

    public sealed class StartNumberElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        public StartNumberElement()
            : base()
        {
            this.PageStartsOn = "Both";
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "STARTNUMBER"; } }
        public SecDefElement Parent { get; set; }

        public string PageStartsOn;
        public int Page;
        public int Figure;
        public int Table;
        public int Equation;

        public static StartNumberElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            StartNumberElement item = new StartNumberElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.PageStartsOn = String.Concat(nav.SelectSingleNode("@PageStartsOn"));

            try { item.Page = Int32.Parse(String.Concat(nav.SelectSingleNode("@Page"))); }
            catch { }

            try { item.Figure = Int32.Parse(String.Concat(nav.SelectSingleNode("@Figure"))); }
            catch { }

            try { item.Table = Int32.Parse(String.Concat(nav.SelectSingleNode("@Table"))); }
            catch { }

            try { item.Equation = Int32.Parse(String.Concat(nav.SelectSingleNode("@Equation"))); }
            catch { }

            return item;
        }
    }

    public sealed class HideElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "HIDE"; } }
        public SecDefElement Parent { get; set; }

        public bool Header;
        public bool Footer;
        public bool MasterPage;
        public bool Border;
        public bool Fill;
        public bool PageNumPos;
        public bool EmptyLine;

        public static HideElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            HideElement item = new HideElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Header = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Header"))); }
            catch { }

            try { item.Footer = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Footer"))); }
            catch { }

            try { item.MasterPage = Boolean.Parse(String.Concat(nav.SelectSingleNode("@MasterPage"))); }
            catch { }

            try { item.Border = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Border"))); }
            catch { }

            try { item.Fill = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Fill"))); }
            catch { }

            try { item.PageNumPos = Boolean.Parse(String.Concat(nav.SelectSingleNode("@PageNumPos"))); }
            catch { }

            try { item.EmptyLine = Boolean.Parse(String.Concat(nav.SelectSingleNode("@EmptyLine"))); }
            catch { }

            return item;
        }
    }

    public enum PageDefGutterType : byte
    {
        LeftOnly,
        LeftRight,
        TopBottom
    }

    public sealed class PageDefElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        public PageDefElement()
            : base()
        {
            this.Width = new HwpUnit(59528);
            this.Height = new HwpUnit(84188);
            this.GutterType = PageDefGutterType.LeftOnly;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "PAGEDEF"; } }
        public SecDefElement Parent { get; set; }

        public PageMarginElement PageMargin;

        public int Landscape;
        public HwpUnit Width;
        public HwpUnit Height;
        public PageDefGutterType GutterType;

        public static PageDefElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            PageDefElement item = new PageDefElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.PageMargin = PageMarginElement.Create(nav.SelectSingleNode("PAGEMARGIN"), item);

            try { item.Landscape = Int32.Parse(String.Concat(nav.SelectSingleNode("@Landscape"))); }
            catch { }

            try { item.Width = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Width")))); }
            catch { }

            try { item.Height = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Height")))); }
            catch { }

            try { item.GutterType = (PageDefGutterType)Enum.Parse(typeof(PageDefGutterType), String.Concat(nav.SelectSingleNode("@GutterType")), true); }
            catch { }

            return item;
        }
    }

    public sealed class PageMarginElement : IHwpmlElement, IHwpmlElement<PageDefElement>
    {
        public PageMarginElement()
            : base()
        {
            this.Left = new HwpUnit(8504);
            this.Right = new HwpUnit(8504);
            this.Top = new HwpUnit(5668);
            this.Bottom = new HwpUnit(4252);
            this.Header = new HwpUnit(4252);
            this.Footer = new HwpUnit(4252);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (PageDefElement)value; }
        }
    
        public string ElementName { get { return "PAGEMARGIN"; } }
        public PageDefElement Parent { get; set; }

        public HwpUnit Left;
        public HwpUnit Right;
        public HwpUnit Top;
        public HwpUnit Bottom;
        public HwpUnit Header;
        public HwpUnit Footer;
        public HwpUnit Gutter;

        public static PageMarginElement Create(IXPathNavigable target, PageDefElement parent)
        {
            if (target == null)
                return null;

            PageMarginElement item = new PageMarginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Left = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Left")))); }
            catch { }

            try { item.Right = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Right")))); }
            catch { }

            try { item.Top = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Top")))); }
            catch { }

            try { item.Bottom = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Bottom")))); }
            catch { }

            try { item.Header = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Header")))); }
            catch { }

            try { item.Footer = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Footer")))); }
            catch { }

            try { item.Gutter = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Gutter")))); }
            catch { }

            return item;
        }
    }

    public sealed class FootnoteShapeElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "FOOTNOTESHAPE"; } }
        public SecDefElement Parent { get; set; }

        public AutoNumFormatElement AutoNumFormat;
        public NoteLineElement NoteLine;
        public NoteSpacingElement NoteSpacing;
        public NoteNumberingElement NoteNumbering;
        public NotePlacementElement NotePlacement;

        public static FootnoteShapeElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            FootnoteShapeElement item = new FootnoteShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.AutoNumFormat = AutoNumFormatElement.Create(nav.SelectSingleNode("AUTONUMFORMAT"), item);
            item.NoteLine = NoteLineElement.Create(nav.SelectSingleNode("NOTELINEELEMENT"), item);
            item.NoteSpacing = NoteSpacingElement.Create(nav.SelectSingleNode("NOTESPACING"), item);
            item.NoteNumbering = NoteNumberingElement.Create(nav.SelectSingleNode("NOTENUMBERING"), item);
            item.NotePlacement = NotePlacementElement.Create(nav.SelectSingleNode("NOTEPLACEMENT"), item);

            return item;
        }
    }

    public sealed class EndnoteShapeElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "ENDNOTESHAPE"; } }
        public SecDefElement Parent { get; set; }

        public AutoNumFormatElement AutoNumFormat;
        public NoteLineElement NoteLine;
        public NoteSpacingElement NoteSpacing;
        public NoteNumberingElement NoteNumbering;
        public NotePlacementElement NotePlacement;

        public static EndnoteShapeElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            EndnoteShapeElement item = new EndnoteShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.AutoNumFormat = AutoNumFormatElement.Create(nav.SelectSingleNode("AUTONUMFORMAT"), item);
            item.NoteLine = NoteLineElement.Create(nav.SelectSingleNode("NOTELINEELEMENT"), item);
            item.NoteSpacing = NoteSpacingElement.Create(nav.SelectSingleNode("NOTESPACING"), item);
            item.NoteNumbering = NoteNumberingElement.Create(nav.SelectSingleNode("NOTENUMBERING"), item);
            item.NotePlacement = NotePlacementElement.Create(nav.SelectSingleNode("NOTEPLACEMENT"), item);

            return item;
        }
    }

    public sealed class AutoNumFormatElement : IHwpmlElement, IHwpmlElement<FootnoteShapeElement>, IHwpmlElement<EndnoteShapeElement>, IHwpmlElement<AutoNumElement>, IHwpmlElement<NewNumElement>
    {
        public AutoNumFormatElement()
            : base()
        {
            this.Type = NumberType2.Digit;
            this.SuffixChar = ")";
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as FootnoteShapeElement) != null)
                    return;
                if ((this.parent = value as EndnoteShapeElement) != null)
                    return;
                if ((this.parent = value as AutoNumElement) != null)
                    return;
                if ((this.parent = value as NewNumElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        FootnoteShapeElement IHwpmlElement<FootnoteShapeElement>.Parent
        {
            get { return this.parent as FootnoteShapeElement; }
            set { this.parent = value; }
        }

        EndnoteShapeElement IHwpmlElement<EndnoteShapeElement>.Parent
        {
            get { return this.parent as EndnoteShapeElement; }
            set { this.parent = value; }
        }

        AutoNumElement IHwpmlElement<AutoNumElement>.Parent
        {
            get { return this.parent as AutoNumElement; }
            set { this.parent = value; }
        }

        NewNumElement IHwpmlElement<NewNumElement>.Parent
        {
            get { return this.parent as NewNumElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "AUTONUMFORMAT"; } }

        public NumberType2 Type;
        public string UserChar;
        public string PrefixChar;
        public string SuffixChar;
        public bool Superscript;

        public static AutoNumFormatElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            AutoNumFormatElement item = new AutoNumFormatElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (NumberType2)Enum.Parse(typeof(NumberType2), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            item.UserChar = String.Concat(nav.SelectSingleNode("@UserChar"));
            item.PrefixChar = String.Concat(nav.SelectSingleNode("@PrefixChar"));
            item.SuffixChar = String.Concat(nav.SelectSingleNode("@SuffixChar"));

            try { item.Superscript = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Superscript"))); }
            catch { }

            return item;
        }
    }

    public enum NoteLineLength : byte
    {
        Empty = 0,
        _5cm = 5,
        _2cm = 2,
        AThirdOfColumn = 3,
        Column = 1
    }

    public sealed class NoteLineElement : IHwpmlElement, IHwpmlElement<FootnoteShapeElement>, IHwpmlElement<EndnoteShapeElement>
    {
        public NoteLineElement()
            : base()
        {
            this.Type = LineType1.Solid;
            this.Width = new LineWidth("0.12mm");
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as FootnoteShapeElement) != null)
                    return;
                if ((this.parent = value as EndnoteShapeElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        FootnoteShapeElement IHwpmlElement<FootnoteShapeElement>.Parent
        {
            get { return this.parent as FootnoteShapeElement; }
            set { this.parent = value; }
        }

        EndnoteShapeElement IHwpmlElement<EndnoteShapeElement>.Parent
        {
            get { return this.parent as EndnoteShapeElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "NOTELINE"; } }

        public string Length;
        public LineType1 Type;
        public LineWidth Width;
        public RGBColor Color;

        public Type GetLengthValueType()
        {
            if (this.Length.Equals("0", StringComparison.OrdinalIgnoreCase) ||
                this.Length.Equals("5cm", StringComparison.OrdinalIgnoreCase) ||
                this.Length.Equals("2cm", StringComparison.OrdinalIgnoreCase) ||
                this.Length.Equals("Column/3", StringComparison.OrdinalIgnoreCase) ||
                this.Length.Equals("Column", StringComparison.OrdinalIgnoreCase))
                return typeof(NoteLineLength);
            else
                return typeof(HwpUnit);
        }

        public bool TryGetLengthValueAsNoteLineLength(ref NoteLineLength result)
        {
            if (!this.GetLengthValueType().Equals(typeof(NoteLineLength)))
                return false;

            if (this.Length.Equals("0", StringComparison.OrdinalIgnoreCase))
            {
                result = NoteLineLength.Empty;
                return true;
            }

            if (this.Length.Equals("5cm", StringComparison.OrdinalIgnoreCase))
            {
                result = NoteLineLength._5cm;
                return true;
            }

            if (this.Length.Equals("2cm", StringComparison.OrdinalIgnoreCase))
            {
                result = NoteLineLength._2cm;
                return true;
            }

            if (this.Length.Equals("Column/3", StringComparison.OrdinalIgnoreCase))
            {
                result = NoteLineLength.AThirdOfColumn;
                return true;
            }

            if (this.Length.Equals("Column", StringComparison.OrdinalIgnoreCase))
            {
                result = NoteLineLength.Column;
                return true;
            }

            return false;
        }

        public bool TryGetLengthValueAsHwpUnit(ref HwpUnit result)
        {
            if (!this.GetLengthValueType().Equals(typeof(HwpUnit)))
                return false;

            try
            {
                result = new HwpUnit(Int32.Parse(this.Length));
                return true;
            }
            catch { return false; }
        }

        public static NoteLineElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            NoteLineElement item = new NoteLineElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Length = String.Concat(nav.SelectSingleNode("@Length"));
            
            try { item.Type = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Width = new LineWidth(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class NoteSpacingElement : IHwpmlElement, IHwpmlElement<FootnoteShapeElement>, IHwpmlElement<EndnoteShapeElement>
    {
        public NoteSpacingElement()
            : base()
        {
            this.AboveLine = new HwpUnit(567);
            this.BelowLine = new HwpUnit(567);
            this.BetweenNotes = new HwpUnit(850);
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as FootnoteShapeElement) != null)
                    return;
                if ((this.parent = value as EndnoteShapeElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        FootnoteShapeElement IHwpmlElement<FootnoteShapeElement>.Parent
        {
            get { return this.parent as FootnoteShapeElement; }
            set { this.parent = value; }
        }

        EndnoteShapeElement IHwpmlElement<EndnoteShapeElement>.Parent
        {
            get { return this.parent as EndnoteShapeElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "NOTESPACING"; } }

        public HwpUnit AboveLine;
        public HwpUnit BelowLine;
        public HwpUnit BetweenNotes;

        public static NoteSpacingElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            NoteSpacingElement item = new NoteSpacingElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.AboveLine = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@AboveLine")))); }
            catch { }

            try { item.BelowLine = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@BelowLine")))); }
            catch { }

            try { item.BetweenNotes = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@BetweenNotes")))); }
            catch { }

            return item;
        }
    }

    public enum NoteNumberingType : byte
    {
        Continuous,
        OnSection,
        OnPage
    }

    public sealed class NoteNumberingElement : IHwpmlElement, IHwpmlElement<FootnoteShapeElement>, IHwpmlElement<EndnoteShapeElement>
    {
        public NoteNumberingElement()
            : base()
        {
            this.Type = NoteNumberingType.Continuous;
            this.NewNumber = 1;
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as FootnoteShapeElement) != null)
                    return;
                if ((this.parent = value as EndnoteShapeElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        FootnoteShapeElement IHwpmlElement<FootnoteShapeElement>.Parent
        {
            get { return this.parent as FootnoteShapeElement; }
            set { this.parent = value; }
        }

        EndnoteShapeElement IHwpmlElement<EndnoteShapeElement>.Parent
        {
            get { return this.parent as EndnoteShapeElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "NOTENUMBERING"; } }

        public NoteNumberingType Type;
        public int NewNumber;

        public static NoteNumberingElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            NoteNumberingElement item = new NoteNumberingElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (NoteNumberingType)Enum.Parse(typeof(NoteNumberingType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }
            
            try { item.NewNumber = Int32.Parse(String.Concat(nav.SelectSingleNode("@NewNumber"))); }
            catch { }

            return item;
        }
    }

    public enum NotePlace : byte
    {
        EachColumn,
        MergedColumn,
        RightMostColumn,
        EndOfDocument,
        EndOfSection
    }

    public sealed class NotePlacementElement : IHwpmlElement, IHwpmlElement<FootnoteShapeElement>, IHwpmlElement<EndnoteShapeElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as FootnoteShapeElement) != null)
                    return;
                if ((this.parent = value as EndnoteShapeElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        FootnoteShapeElement IHwpmlElement<FootnoteShapeElement>.Parent
        {
            get { return this.parent as FootnoteShapeElement; }
            set { this.parent = value; }
        }

        EndnoteShapeElement IHwpmlElement<EndnoteShapeElement>.Parent
        {
            get { return this.parent as EndnoteShapeElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "NOTEPLACEMENT"; } }

        public NotePlace Place;
        public bool BeneathText;

        public static NotePlacementElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            NotePlacementElement item = new NotePlacementElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            if (parent is FootnoteShapeElement)
                item.Place = NotePlace.EachColumn;

            if (parent is EndnoteShapeElement)
                item.Place = NotePlace.EndOfDocument;

            try { item.Place = (NotePlace)Enum.Parse(typeof(NotePlace), String.Concat(nav.SelectSingleNode("@Place")), true); }
            catch { }

            try { item.BeneathText = Boolean.Parse(String.Concat(nav.SelectSingleNode("@BeneathText"))); }
            catch { }

            return item;
        }
    }

    public enum PageBorderFillType : byte
    {
        Both,
        Even,
        Odd
    }

    public enum PageBorderFillArea : byte
    {
        Paper,
        Page,
        Border
    }

    public sealed class PageBorderFillElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        public PageBorderFillElement()
            : base()
        {
            this.Type = PageBorderFillType.Both;
            this.FillArea = PageBorderFillArea.Paper;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "PAGEBORDERFILL"; } }
        public SecDefElement Parent { get; set; }

        public PageOffsetElement PageOffset;

        public PageBorderFillType Type;
        public int BorderFill;
        public bool TextBorder;
        public bool HeaderInside;
        public bool FooterInside;
        public PageBorderFillArea FillArea;

        public static PageBorderFillElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            PageBorderFillElement item = new PageBorderFillElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.PageOffset = PageOffsetElement.Create(nav.SelectSingleNode("PAGEOFFSET"), item);

            try { item.Type = (PageBorderFillType)Enum.Parse(typeof(PageBorderFillType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }
            
            try { item.BorderFill = Int32.Parse(String.Concat(nav.SelectSingleNode("@BorderFill"))); }
            catch { }

            try { item.TextBorder = Boolean.Parse(String.Concat(nav.SelectSingleNode("@TextBorder"))); }
            catch { }

            try { item.HeaderInside = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HeaderInside"))); }
            catch { }

            try { item.FooterInside = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FooterInside"))); }
            catch { }

            try { item.FillArea = (PageBorderFillArea)Enum.Parse(typeof(PageBorderFillArea), String.Concat(nav.SelectSingleNode("@FillArea")), true); }
            catch { }

            return item;
        }
    }

    public sealed class PageOffsetElement : IHwpmlElement, IHwpmlElement<PageBorderFillElement>
    {
        public PageOffsetElement()
            : base()
        {
            this.Left = new HwpUnit(1417);
            this.Right = new HwpUnit(1417);
            this.Top = new HwpUnit(1417);
            this.Bottom = new HwpUnit(1417);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (PageBorderFillElement)value; }
        }

        public string ElementName { get { return "PAGEOFFSET"; } }
        public PageBorderFillElement Parent { get; set; }

        public HwpUnit Left;
        public HwpUnit Right;
        public HwpUnit Top;
        public HwpUnit Bottom;

        public static PageOffsetElement Create(IXPathNavigable target, PageBorderFillElement parent)
        {
            if (target == null)
                return null;

            PageOffsetElement item = new PageOffsetElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Left = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Left")))); }
            catch { }

            try { item.Right = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Right")))); }
            catch { }

            try { item.Top = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Top")))); }
            catch { }

            try { item.Bottom = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Bottom")))); }
            catch { }

            return item;
        }
    }

    public enum MasterPageType : byte
    {
        Both,
        Even,
        Odd
    }

    public sealed class MasterPageElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        public MasterPageElement()
            : base()
        {
            this.Type = MasterPageType.Both;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "MASTERPAGE"; } }
        public SecDefElement Parent { get; set; }

        public ParaListElement ParaList;

        public MasterPageType Type;
        public int TextWidth;
        public int TextHeight;
        public bool HasTextRef;
        public bool HasNumRef;

        public static MasterPageElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            MasterPageElement item = new MasterPageElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("ParaList"), item);

            try { item.Type = (MasterPageType)Enum.Parse(typeof(MasterPageType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }
            
            try { item.TextWidth = Int32.Parse(String.Concat(nav.SelectSingleNode("@TextWidth"))); }
            catch { }

            try { item.TextHeight = Int32.Parse(String.Concat(nav.SelectSingleNode("@TextHeight"))); }
            catch { }

            try { item.HasTextRef = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HasTextRef"))); }
            catch { }

            try { item.HasNumRef = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HasNumRef"))); }
            catch { }

            return item;
        }
    }

    public enum ParaListVertAlign : byte
    {
        Top,
        Center,
        Bottom
    }

    public sealed class ParaListElement : List<PElement>, IHwpmlElement, IHwpmlElement<MasterPageElement>, IHwpmlElement<ExtMasterPageElement>, IHwpmlElement<CellElement>, IHwpmlElement<DrawTextElement>, IHwpmlElement<CaptionElement>, IHwpmlElement<HeaderElement>, IHwpmlElement<FooterElement>, IHwpmlElement<FootnoteElement>, IHwpmlElement<EndnoteElement>, IHwpmlElement<HiddenCommentElement>
    {
        public ParaListElement()
            : base()
        {
            this.LineWrap = LineWrapType.Break;
            this.VertAlign = ParaListVertAlign.Top;
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as MasterPageElement) != null)
                    return;
                if ((this.parent = value as ExtMasterPageElement) != null)
                    return;
                if ((this.parent = value as CellElement) != null)
                    return;
                if ((this.parent = value as DrawTextElement) != null)
                    return;
                if ((this.parent = value as CaptionElement) != null)
                    return;
                if ((this.parent = value as HeaderElement) != null)
                    return;
                if ((this.parent = value as FooterElement) != null)
                    return;
                if ((this.parent = value as FootnoteElement) != null)
                    return;
                if ((this.parent = value as EndnoteElement) != null)
                    return;
                if ((this.parent = value as HiddenCommentElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        MasterPageElement IHwpmlElement<MasterPageElement>.Parent
        {
            get { return this.parent as MasterPageElement; }
            set { this.parent = value; }
        }

        ExtMasterPageElement IHwpmlElement<ExtMasterPageElement>.Parent
        {
            get { return this.parent as ExtMasterPageElement; }
            set { this.parent = value; }
        }

        CellElement IHwpmlElement<CellElement>.Parent
        {
            get { return this.parent as CellElement; }
            set { this.parent = value; }
        }

        DrawTextElement IHwpmlElement<DrawTextElement>.Parent
        {
            get { return this.parent as DrawTextElement; }
            set { this.parent = value; }
        }

        CaptionElement IHwpmlElement<CaptionElement>.Parent
        {
            get { return this.parent as CaptionElement; }
            set { this.parent = value; }
        }

        HeaderElement IHwpmlElement<HeaderElement>.Parent
        {
            get { return this.parent as HeaderElement; }
            set { this.parent = value; }
        }

        FooterElement IHwpmlElement<FooterElement>.Parent
        {
            get { return this.parent as FooterElement; }
            set { this.parent = value; }
        }

        FootnoteElement IHwpmlElement<FootnoteElement>.Parent
        {
            get { return this.parent as FootnoteElement; }
            set { this.parent = value; }
        }

        EndnoteElement IHwpmlElement<EndnoteElement>.Parent
        {
            get { return this.parent as EndnoteElement; }
            set { this.parent = value; }
        }

        HiddenCommentElement IHwpmlElement<HiddenCommentElement>.Parent
        {
            get { return this.parent as HiddenCommentElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "PARALIST"; } }

        public int TextDirection;
        public LineWrapType LineWrap;
        public ParaListVertAlign VertAlign;
        public string LinkListID;
        public string LinkListIDNext;

        public static ParaListElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ParaListElement item = new ParaListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("P"))
                item.Add(PElement.Create(each, item));

            try { item.TextDirection = Int32.Parse(String.Concat(nav.SelectSingleNode("@TextDirection"))); }
            catch { }

            try { item.LineWrap = (LineWrapType)Enum.Parse(typeof(LineWrapType), String.Concat(nav.SelectSingleNode("@LineWrap")), true); }
            catch { }

            try { item.VertAlign = (ParaListVertAlign)Enum.Parse(typeof(ParaListVertAlign), String.Concat(nav.SelectSingleNode("@VertAlign")), true); }
            catch { }

            item.LinkListID = String.Concat(nav.SelectSingleNode("@LinkListID"));
            item.LinkListIDNext = String.Concat(nav.SelectSingleNode("@LinkListIDNext"));

            return item;
        }
    }

    public enum ExtMasterPageType : byte
    {
        LastPage,
        OptionalPage
    }

    public sealed class ExtMasterPageElement : IHwpmlElement, IHwpmlElement<SecDefElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (SecDefElement)value; }
        }

        public string ElementName { get { return "EXT_MASTERPAGE"; } }
        public SecDefElement Parent { get; set; }

        public ParaListElement ParaList;

        public ExtMasterPageType Type;
        public int PageNumber;
        public bool PageDuplicate;
        public bool PageFront;

        public static ExtMasterPageElement Create(IXPathNavigable target, SecDefElement parent)
        {
            if (target == null)
                return null;

            ExtMasterPageElement item = new ExtMasterPageElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            try { item.Type = (ExtMasterPageType)Enum.Parse(typeof(ExtMasterPageType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }
            
            try { item.PageNumber = Int32.Parse(String.Concat(nav.SelectSingleNode("@PageNumber"))); }
            catch { }

            try { item.PageDuplicate = Boolean.Parse(String.Concat(nav.SelectSingleNode("@PageDuplicate"))); }
            catch { }

            try { item.PageFront = Boolean.Parse(String.Concat(nav.SelectSingleNode("@PageFront"))); }
            catch { }

            return item;
        }
    }

    public enum ColDefType : byte
    {
        Newspaper,
        BalancedNewspaper,
        Parallel
    }

    public enum ColDefLayout : byte
    {
        Left,
        Right,
        Mirror
    }

    public sealed class ColDefElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public ColDefElement()
            : base()
        {
            this.ColumnLine = new List<ColumnLineElement>();
            this.ColumnTable = new List<ColumnTableElement>();

            this.Type = ColDefType.Newspaper;
            this.Count = 1;
            this.Layout = ColDefLayout.Left;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "COLDEF"; } }
        public TextElement Parent { get; set; }

        public ParameterSetElement ParameterSet;
        public IList<ColumnLineElement> ColumnLine;
        public IList<ColumnTableElement> ColumnTable;

        public ColDefType Type;
        public int Count;
        public ColDefLayout Layout;
        public bool SameSize;
        public HwpUnit SameGap;

        public static ColDefElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            ColDefElement item = new ColDefElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParameterSet = ParameterSetElement.Create(target, item);

            foreach (IXPathNavigable each in nav.Select("COLUMNLINE"))
                item.ColumnLine.Add(ColumnLineElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("COLUMNTABLE"))
                item.ColumnTable.Add(ColumnTableElement.Create(each, item));

            try { item.Type = (ColDefType)Enum.Parse(typeof(ColDefType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }
            
            try { item.Count = Int32.Parse(String.Concat(nav.SelectSingleNode("@Count"))); }
            catch { }

            try { item.Layout = (ColDefLayout)Enum.Parse(typeof(ColDefLayout), String.Concat(nav.SelectSingleNode("@Layout")), true); }
            catch { }

            try { item.SameSize = Boolean.Parse(String.Concat(nav.SelectSingleNode("@SameSize"))); }
            catch { }

            try { item.SameGap = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@SameGap")))); }
            catch { }

            return item;
        }
    }

    public sealed class ColumnLineElement : IHwpmlElement, IHwpmlElement<ColDefElement>
    {
        public ColumnLineElement()
            : base()
        {
            this.Type = LineType1.Solid;
            this.Width = new LineWidth("0.12mm");
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ColDefElement)value; }
        }

        public string ElementName { get { return "COLUMNLINE"; } }
        public ColDefElement Parent { get; set; }

        public LineType1 Type;
        public LineWidth Width;
        public RGBColor Color;

        public static ColumnLineElement Create(IXPathNavigable target, ColDefElement parent)
        {
            if (target == null)
                return null;

            ColumnLineElement item = new ColumnLineElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Width = new LineWidth(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            return item;
        }
    }

    public sealed class ColumnTableElement : List<ColumnElement>, IHwpmlElement, IHwpmlElement<ColDefElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ColDefElement)value; }
        }

        public string ElementName { get { return "COLUMNTABLE"; } }
        public ColDefElement Parent { get; set; }

        public static ColumnTableElement Create(IXPathNavigable target, ColDefElement parent)
        {
            if (target == null)
                return null;

            ColumnTableElement item = new ColumnTableElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("COLUMN"))
                item.Add(ColumnElement.Create(each, item));

            return item;
        }
    }

    public sealed class ColumnElement : IHwpmlElement, IHwpmlElement<ColumnTableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ColumnTableElement)value; }
        }

        public string ElementName { get { return "COLUMN"; } }
        public ColumnTableElement Parent { get; set; }

        public HwpUnit Width;
        public HwpUnit Gap;

        public static ColumnElement Create(IXPathNavigable target, ColumnTableElement parent)
        {
            if (target == null)
                return null;

            ColumnElement item = new ColumnElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Width = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Width")))); }
            catch { }

            try { item.Gap = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Gap")))); }
            catch { }

            return item;
        }
    }

    public enum TablePageBreak : byte
    {
        Table,
        Cell,
        None
    }

    public sealed class TableElement : List<RowElement>, IHwpmlElement, IHwpmlElement<TextElement>
    {
        public TableElement()
            : base()
        {
            this.PageBreak = TablePageBreak.Cell;
            this.RepeatHeader = true;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "TABLE"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public InsideMarginElement InsideMargin;
        public CellZoneListElement CellZoneList;

        public TablePageBreak PageBreak;
        public bool RepeatHeader;
        public int RowCount;
        public int ColCount;
        public HwpUnit CellSpacing;
        public int BorderFill;

        public static TableElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            TableElement item = new TableElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.InsideMargin = InsideMarginElement.Create(nav.SelectSingleNode("INSIDEMARGIN"), item);
            item.CellZoneList = CellZoneListElement.Create(nav.SelectSingleNode("CELLZONELIST"), item);

            foreach (IXPathNavigable each in nav.Select("ROW"))
                item.Add(RowElement.Create(each, item));

            try { item.PageBreak = (TablePageBreak)Enum.Parse(typeof(TablePageBreak), String.Concat(nav.SelectSingleNode("@PageBreak")), true); }
            catch { }
            
            try { item.RepeatHeader = Boolean.Parse(String.Concat(nav.SelectSingleNode("@RepeatHeader"))); }
            catch { }

            try { item.RowCount = Int32.Parse(String.Concat(nav.SelectSingleNode("@RowCount"))); }
            catch { }

            try { item.ColCount = Int32.Parse(String.Concat(nav.SelectSingleNode("@ColCount"))); }
            catch { }

            try { item.CellSpacing = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@CellSpacing")))); }
            catch { }

            try { item.BorderFill = Int32.Parse(String.Concat(nav.SelectSingleNode("@BorderFill"))); }
            catch { }

            return item;
        }
    }

    public enum ShapeObjectNumberingType : byte
    {
        None,
        Figure,
        Table,
        Equation
    }

    public enum ShapeObjectTextFlow : byte
    {
        BothSides,
        LeftOnly,
        RightOnly,
        LargestOnly
    }

    public sealed class ShapeObjectElement : IHwpmlElement, IHwpmlElement<TableElement>, IHwpmlElement<PictureElement>, IHwpmlElement<LineElement>, IHwpmlElement<RectangleElement>, IHwpmlElement<EllipseElement>, IHwpmlElement<ArcElement>, IHwpmlElement<PolygonElement>, IHwpmlElement<OleElement>, IHwpmlElement<EquationElement>, IHwpmlElement<UnknownObjectElement>, IHwpmlElement<ButtonElement>, IHwpmlElement<RadioButtonElement>, IHwpmlElement<CheckButtonElement>, IHwpmlElement<EditElement>, IHwpmlElement<ListBoxElement>, IHwpmlElement<ScrollBarElement>, IHwpmlElement<ContainerElement>, IHwpmlElement<ConnectLineElement>
    {
        public ShapeObjectElement()
            : base()
        {
            this.NumberingType = ShapeObjectNumberingType.None;
            this.TextFlow = ShapeObjectTextFlow.BothSides;
        }

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TableElement) != null)
                    return;
                if ((this.parent = value as PictureElement) != null)
                    return;
                if ((this.parent = value as LineElement) != null)
                    return;
                if ((this.parent = value as RectangleElement) != null)
                    return;
                if ((this.parent = value as EllipseElement) != null)
                    return;
                if ((this.parent = value as ArcElement) != null)
                    return;
                if ((this.parent = value as PolygonElement) != null)
                    return;
                if ((this.parent = value as OleElement) != null)
                    return;
                if ((this.parent = value as EquationElement) != null)
                    return;
                if ((this.parent = value as UnknownObjectElement) != null)
                    return;
                if ((this.parent = value as ButtonElement) != null)
                    return;
                if ((this.parent = value as RadioButtonElement) != null)
                    return;
                if ((this.parent = value as CheckButtonElement) != null)
                    return;
                if ((this.parent = value as EditElement) != null)
                    return;
                if ((this.parent = value as ListBoxElement) != null)
                    return;
                if ((this.parent = value as ScrollBarElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                if ((this.parent = value as ConnectLineElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TableElement IHwpmlElement<TableElement>.Parent
        {
            get { return this.parent as TableElement; }
            set { this.parent = value; }
        }

        PictureElement IHwpmlElement<PictureElement>.Parent
        {
            get { return this.parent as PictureElement; }
            set { this.parent = value; }
        }

        LineElement IHwpmlElement<LineElement>.Parent
        {
            get { return this.parent as LineElement; }
            set { this.parent = value; }
        }

        RectangleElement IHwpmlElement<RectangleElement>.Parent
        {
            get { return this.parent as RectangleElement; }
            set { this.parent = value; }
        }

        EllipseElement IHwpmlElement<EllipseElement>.Parent
        {
            get { return this.parent as EllipseElement; }
            set { this.parent = value; }
        }

        ArcElement IHwpmlElement<ArcElement>.Parent
        {
            get { return this.parent as ArcElement; }
            set { this.parent = value; }
        }

        PolygonElement IHwpmlElement<PolygonElement>.Parent
        {
            get { return this.parent as PolygonElement; }
            set { this.parent = value; }
        }

        OleElement IHwpmlElement<OleElement>.Parent
        {
            get { return this.parent as OleElement; }
            set { this.parent = value; }
        }

        EquationElement IHwpmlElement<EquationElement>.Parent
        {
            get { return this.parent as EquationElement; }
            set { this.parent = value; }
        }

        UnknownObjectElement IHwpmlElement<UnknownObjectElement>.Parent
        {
            get { return this.parent as UnknownObjectElement; }
            set { this.parent = value; }
        }

        ButtonElement IHwpmlElement<ButtonElement>.Parent
        {
            get { return this.parent as ButtonElement; }
            set { this.parent = value; }
        }

        RadioButtonElement IHwpmlElement<RadioButtonElement>.Parent
        {
            get { return this.parent as RadioButtonElement; }
            set { this.parent = value; }
        }

        CheckButtonElement IHwpmlElement<CheckButtonElement>.Parent
        {
            get { return this.parent as CheckButtonElement; }
            set { this.parent = value; }
        }

        EditElement IHwpmlElement<EditElement>.Parent
        {
            get { return this.parent as EditElement; }
            set { this.parent = value; }
        }

        ListBoxElement IHwpmlElement<ListBoxElement>.Parent
        {
            get { return this.parent as ListBoxElement; }
            set { this.parent = value; }
        }

        ScrollBarElement IHwpmlElement<ScrollBarElement>.Parent
        {
            get { return this.parent as ScrollBarElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        ConnectLineElement IHwpmlElement<ConnectLineElement>.Parent
        {
            get { return this.parent as ConnectLineElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "SHAPEOBJECT"; } }

        private IHwpmlElement parent;

        public SizeElement Size;
        public PositionElement Position;
        public OutsideMarginElement OutsideMargin;
        public CaptionElement Caption;
        public string ShapeComment;

        public string InstId;
        public int ZOrder;
        public ShapeObjectNumberingType NumberingType;
        public TextWrapType TextWrap;
        public ShapeObjectTextFlow TextFlow;
        public bool Lock;

        public static ShapeObjectElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ShapeObjectElement item = new ShapeObjectElement();

            XPathNavigator nav = target.CreateNavigator();
            item.parent = parent;
            item.Size = SizeElement.Create(nav.SelectSingleNode("SIZE"), item);
            item.Position = PositionElement.Create(nav.SelectSingleNode("POSITION"), item);
            item.OutsideMargin = OutsideMarginElement.Create(nav.SelectSingleNode("OUTSIDEMARGIN"), item);
            item.Caption = CaptionElement.Create(nav.SelectSingleNode("CAPTION"), item);
            item.ShapeComment = String.Concat(nav.SelectSingleNode("SHAPECOMMENT"));

            item.InstId = String.Concat(nav.SelectSingleNode("@InstId"));
            
            try { item.ZOrder = Int32.Parse(String.Concat(nav.SelectSingleNode("@ZOrder"))); }
            catch { }

            try { item.NumberingType = (ShapeObjectNumberingType)Enum.Parse(typeof(ShapeObjectNumberingType), String.Concat(nav.SelectSingleNode("@NumberingType")), true); }
            catch { }

            try { item.TextWrap = (TextWrapType)Enum.Parse(typeof(TextWrapType), String.Concat(nav.SelectSingleNode("@TextWrap")), true); }
            catch { }

            try { item.TextFlow = (ShapeObjectTextFlow)Enum.Parse(typeof(ShapeObjectTextFlow), String.Concat(nav.SelectSingleNode("@TextFlow")), true); }
            catch { }
            
            try { item.Lock = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Lock"))); }
            catch { }

            return item;
        }
    }

    public enum SizeWidthRel : byte
    {
        Paper,
        Page,
        Column,
        Para,
        Absolute
    }

    public enum SizeHeightRel : byte
    {
        Paper,
        Page,
        Absolute
    }

    public sealed class SizeElement : IHwpmlElement, IHwpmlElement<ShapeObjectElement>
    {
        public SizeElement()
            : base()
        {
            this.WidthRelTo = SizeWidthRel.Absolute;
            this.HeightRelTo = SizeHeightRel.Absolute;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ShapeObjectElement)value; }
        }

        public string ElementName { get { return "SIZE"; } }
        public ShapeObjectElement Parent { get; set; }

        public HwpUnit Width;
        public HwpUnit Height;
        public SizeWidthRel WidthRelTo;
        public SizeHeightRel HeightRelTo;
        public bool Protect;

        public static SizeElement Create(IXPathNavigable target, ShapeObjectElement parent)
        {
            if (target == null)
                return null;

            SizeElement item = new SizeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Width = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Width")))); }
            catch { }

            try { item.Height = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Height")))); }
            catch { }

            try { item.WidthRelTo = (SizeWidthRel)Enum.Parse(typeof(SizeWidthRel), String.Concat(nav.SelectSingleNode("@WidthRelTo")), true); }
            catch { }

            try { item.HeightRelTo = (SizeHeightRel)Enum.Parse(typeof(SizeHeightRel), String.Concat(nav.SelectSingleNode("@HeightRelTo")), true); }
            catch { }

            try { item.Protect = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Protect"))); }
            catch { }

            return item;
        }
    }

    public enum PositionVertRelTo : byte
    {
        Paper,
        Page,
        Para
    }

    public enum PositionVertAlign : byte
    {
        Top,
        Center,
        Bottom,
        Inside,
        Outside
    }

    public enum PositionHorzRelTo : byte
    {
        Paper,
        Page,
        Column,
        Para
    }

    public enum PositionHorzAlign : byte
    {
        Left,
        Center,
        Right,
        Inside,
        Outside
    }

    public sealed class PositionElement : IHwpmlElement, IHwpmlElement<ShapeObjectElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ShapeObjectElement)value; }
        }
        
        public string ElementName { get { return "POSITION"; } }
        public ShapeObjectElement Parent { get; set; }

        public bool TreatAsChar;
        public bool AffectLSpacing;
        public PositionVertRelTo VertRelTo;
        public PositionVertAlign VertAlign;
        public PositionHorzRelTo HorzRelTo;
        public PositionHorzAlign HorzAlign;
        public HwpUnit VertOffset;
        public HwpUnit HorzOffset;
        public bool FlowWithText;
        public bool AllowOverlap;
        public bool HoldAnchorAndSO;

        public static PositionElement Create(IXPathNavigable target, ShapeObjectElement parent)
        {
            if (target == null)
                return null;

            PositionElement item = new PositionElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.TreatAsChar = Boolean.Parse(String.Concat(nav.SelectSingleNode("@TreatAsChar"))); }
            catch { }

            try { item.AffectLSpacing = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AffectLSpacing"))); }
            catch { }

            try { item.VertRelTo = (PositionVertRelTo)Enum.Parse(typeof(PositionVertRelTo), String.Concat(nav.SelectSingleNode("@VertRelTo")), true); }
            catch { }

            try { item.VertAlign = (PositionVertAlign)Enum.Parse(typeof(PositionVertAlign), String.Concat(nav.SelectSingleNode("@VertAlign")), true); }
            catch { }

            try { item.HorzRelTo = (PositionHorzRelTo)Enum.Parse(typeof(PositionHorzRelTo), String.Concat(nav.SelectSingleNode("@HorzRelTo")), true); }
            catch { }

            try { item.HorzAlign = (PositionHorzAlign)Enum.Parse(typeof(PositionHorzAlign), String.Concat(nav.SelectSingleNode("@HorzAlign")), true); }
            catch { }

            try { item.VertOffset = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@VertOffset")))); }
            catch { }

            try { item.HorzOffset = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@HorzOffset")))); }
            catch { }

            try { item.FlowWithText = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FlowWithText"))); }
            catch { }

            try { item.AllowOverlap = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AllowOverlap"))); }
            catch { }

            try { item.HoldAnchorAndSO = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HoldAnchorAndSO"))); }
            catch { }

            return item;
        }
    }

    public sealed class OutsideMarginElement : IHwpmlElement, IHwpmlElement<ShapeObjectElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ShapeObjectElement)value; }
        }

        public string ElementName { get { return "OUTSIDEMARGIN"; } }
        public ShapeObjectElement Parent { get; set; }

        public HwpUnit Left;
        public HwpUnit Right;
        public HwpUnit Top;
        public HwpUnit Bottom;

        public static OutsideMarginElement Create(IXPathNavigable target, ShapeObjectElement parent)
        {
            if (target == null)
                return null;

            OutsideMarginElement item = new OutsideMarginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            int value = 0;
            if (parent.Parent is TableElement)
                value = 283;
            if (parent.Parent is EquationElement)
                value = 56;

            item.Left = new HwpUnit(value);
            item.Right = new HwpUnit(value);
            item.Top = new HwpUnit(value);
            item.Bottom = new HwpUnit(value);

            try { item.Left = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Left")))); }
            catch { }

            try { item.Right = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Right")))); }
            catch { }

            try { item.Top = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Top")))); }
            catch { }

            try { item.Bottom = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Bottom")))); }
            catch { }

            return item;
        }
    }

    public enum CaptionSide : byte
    {
        Left,
        Right,
        Top,
        Bottom
    }

    public sealed class CaptionElement : IHwpmlElement, IHwpmlElement<ShapeObjectElement>
    {
        public CaptionElement()
            : base()
        {
            this.Side = CaptionSide.Left;
            this.Gap = new HwpUnit(850);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ShapeObjectElement)value; }
        }

        public string ElementName { get { return "CAPTION"; } }
        public ShapeObjectElement Parent { get; set; }

        public ParaListElement ParaList;

        public CaptionSide Side;
        public bool FullSize;
        public int Width;
        public HwpUnit Gap;
        public int LastWidth;

        public static CaptionElement Create(IXPathNavigable target, ShapeObjectElement parent)
        {
            if (target == null)
                return null;

            CaptionElement item = new CaptionElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            try { item.Side = (CaptionSide)Enum.Parse(typeof(CaptionSide), String.Concat(nav.SelectSingleNode("@Side")), true); }
            catch { }
            
            try { item.FullSize = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FullSize"))); }
            catch { }

            try { item.Width = Int32.Parse(String.Concat(nav.SelectSingleNode("@Width"))); }
            catch { }

            try { item.Gap = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Gap")))); }
            catch { }

            try { item.LastWidth = Int32.Parse(String.Concat(nav.SelectSingleNode("@LastWidth"))); }
            catch { }

            return item;
        }
    }

    public sealed class InsideMarginElement : IHwpmlElement, IHwpmlElement<TableElement>, IHwpmlElement<PictureElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = null; return; }
                if ((this.parent = value as TableElement) != null)
                    return;
                if ((this.parent = value as PictureElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TableElement IHwpmlElement<TableElement>.Parent
        {
            get { return this.parent as TableElement; }
            set { this.parent = value; }
        }

        PictureElement IHwpmlElement<PictureElement>.Parent
        {
            get { return this.parent as PictureElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "INSIDEMARGIN"; } }

        public HwpUnit Left;
        public HwpUnit Right;
        public HwpUnit Top;
        public HwpUnit Bottom;

        public static InsideMarginElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            InsideMarginElement item = new InsideMarginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            int value = 0;
            if (parent is TableElement)
                value = 141;

            item.Left = new HwpUnit(value);
            item.Right = new HwpUnit(value);
            item.Top = new HwpUnit(value);
            item.Bottom = new HwpUnit(value);

            try { item.Left = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Left")))); }
            catch { }

            try { item.Right = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Right")))); }
            catch { }

            try { item.Top = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Top")))); }
            catch { }

            try { item.Bottom = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Bottom")))); }
            catch { }

            return item;
        }
    }

    public sealed class CellZoneListElement : List<CellZoneElement>, IHwpmlElement, IHwpmlElement<TableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TableElement)value; }
        }

        public string ElementName { get { return "CELLZONELINE"; } }
        public TableElement Parent { get; set; }

        public static CellZoneListElement Create(IXPathNavigable target, TableElement parent)
        {
            if (target == null)
                return null;

            CellZoneListElement item = new CellZoneListElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("CELLZONE"))
                item.Add(CellZoneElement.Create(each, item));

            return item;
        }
    }

    public sealed class CellZoneElement : IHwpmlElement, IHwpmlElement<CellZoneListElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CellZoneListElement)value; }
        }

        public string ElementName { get { return "CELLZONE"; } }
        public CellZoneListElement Parent { get; set; }

        public string StartRowAddr;
        public string StartColAddr;
        public string EndRowAddr;
        public string EndColAddr;
        public int BorderFill;

        public static CellZoneElement Create(IXPathNavigable target, CellZoneListElement parent)
        {
            if (target == null)
                return null;

            CellZoneElement item = new CellZoneElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.StartRowAddr = String.Concat(nav.SelectSingleNode("@StartRowAddr"));
            item.StartColAddr = String.Concat(nav.SelectSingleNode("@StartColAddr"));
            item.EndRowAddr = String.Concat(nav.SelectSingleNode("@EndRowAddr"));
            item.EndColAddr = String.Concat(nav.SelectSingleNode("@EndColAddr"));
            
            try { item.BorderFill = Int32.Parse(String.Concat(nav.SelectSingleNode("@BorderFill"))); }
            catch { }

            return item;
        }
    }

    public sealed class RowElement : List<CellElement>, IHwpmlElement, IHwpmlElement<TableElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TableElement)value; }
        }

        public string ElementName { get { return "ROW"; } }
        public TableElement Parent { get; set; }

        public static RowElement Create(IXPathNavigable target, TableElement parent)
        {
            if (target == null)
                return null;

            RowElement item = new RowElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("CELL"))
                item.Add(CellElement.Create(each, item));

            return item;
        }
    }

    public sealed class CellElement : IHwpmlElement, IHwpmlElement<RowElement>
    {
        public CellElement()
            : base()
        {
            this.ColSpan = 1;
            this.RowSpan = 1;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (RowElement)value; }
        }

        public string ElementName { get { return "CELL"; } }
        public RowElement Parent { get; set; }

        public CellMarginElement CellMargin;
        public ParaListElement ParaList;

        public string Name;
        public string ColAddr;
        public string RowAddr;
        public int ColSpan;
        public int RowSpan;
        public HwpUnit Width;
        public HwpUnit Height;
        public bool Header;
        public bool HasMargin;
        public bool Protect;
        public bool Editable;
        public bool Dirty;
        public int BorderFill;

        public static CellElement Create(IXPathNavigable target, RowElement parent)
        {
            if (target == null)
                return null;

            CellElement item = new CellElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.CellMargin = CellMarginElement.Create(nav.SelectSingleNode("CELLMARGIN"), item);
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));
            item.ColAddr = String.Concat(nav.SelectSingleNode("@ColAddr"));
            item.RowAddr = String.Concat(nav.SelectSingleNode("@RowAddr"));

            try { item.ColSpan = Int32.Parse(String.Concat(nav.SelectSingleNode("@ColSpan"))); }
            catch { }

            try { item.RowSpan = Int32.Parse(String.Concat(nav.SelectSingleNode("@RowSpan"))); }
            catch { }

            try { item.Width = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Width")))); }
            catch { }

            try { item.Height = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Height")))); }
            catch { }

            try { item.Header = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Header"))); }
            catch { }

            try { item.HasMargin = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HasMargin"))); }
            catch { }

            try { item.Protect = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Protect"))); }
            catch { }

            try { item.Editable = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Editable"))); }
            catch { }

            try { item.Dirty = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Dirty"))); }
            catch { }

            try { item.BorderFill = Int32.Parse(String.Concat(nav.SelectSingleNode("@BorderFill"))); }
            catch { }

            return item;
        }
    }

    public sealed class CellMarginElement : IHwpmlElement, IHwpmlElement<CellElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CellElement)value; }
        }

        public string ElementName { get { return "CELLMARGIN"; } }
        public CellElement Parent { get; set; }

        public HwpUnit Left;
        public HwpUnit Right;
        public HwpUnit Top;
        public HwpUnit Bottom;

        public static CellMarginElement Create(IXPathNavigable target, CellElement parent)
        {
            if (target == null)
                return null;

            CellMarginElement item = new CellMarginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Left = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Left")))); }
            catch { }

            try { item.Right = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Right")))); }
            catch { }

            try { item.Top = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Top")))); }
            catch { }

            try { item.Bottom = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Bottom")))); }
            catch { }

            return item;
        }
    }

    public sealed class PictureElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "PICTURE"; } }

        public ShapeObjectElement ShapeObject;
        public ShapeComponentElement ShapeComponent;
        public LineShapeElement LineShape;
        public ImageRectElement ImageRect;
        public ImageClipElement ImageClip;
        public EffectsElement Effects;
        public InsideMarginElement InsideMargin;
        public ImageElement Image;

        public bool Reverse;

        public static PictureElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            PictureElement item = new PictureElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.ShapeComponent = ShapeComponentElement.Create(nav.SelectSingleNode("SHAPECOMPONENT"), item);
            item.LineShape = LineShapeElement.Create(nav.SelectSingleNode("LINESHAPE"), item);
            item.ImageRect = ImageRectElement.Create(nav.SelectSingleNode("IMAGERECT"), item);
            item.ImageClip = ImageClipElement.Create(nav.SelectSingleNode("IMAGECLIP"), item);
            item.Effects = EffectsElement.Create(nav.SelectSingleNode("EFFECTS"), item);
            item.InsideMargin = InsideMarginElement.Create(nav.SelectSingleNode("INSIDEMARGIN"), item);
            item.Image = ImageElement.Create(nav.SelectSingleNode("IMAGE"), item);

            try { item.Reverse = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Reverse"))); }
            catch { }

            return item;
        }
    }

    public sealed class ShapeComponentElement : IHwpmlElement, IHwpmlElement<PictureElement>, IHwpmlElement<DrawingObjectElement>, IHwpmlElement<ContainerElement>, IHwpmlElement<OleElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = null; return; }
                if ((this.parent = value as PictureElement) != null)
                    return;
                if ((this.parent = value as DrawingObjectElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                if ((this.parent = value as OleElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        PictureElement IHwpmlElement<PictureElement>.Parent
        {
            get { return this.parent as PictureElement; }
            set { this.parent = value; }
        }

        DrawingObjectElement IHwpmlElement<DrawingObjectElement>.Parent
        {
            get { return this.parent as DrawingObjectElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        OleElement IHwpmlElement<OleElement>.Parent
        {
            get { return this.parent as OleElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "SHAPECOMPONENT"; } }

        public ParameterSetElement ParameterSet;
        public RotationInfoElement RotationInfo;
        public RenderingInfoElement RenderingInfo;

        public string HRef;
        public HwpUnit XPos;
        public HwpUnit YPos;
        public int GroupLevel;
        public HwpUnit OriWidth;
        public HwpUnit OriHeight;
        public HwpUnit CurWidth;
        public HwpUnit CurHeight;
        public bool HorzFlip;
        public bool VertFlip;
        public string InstID;

        public static ShapeComponentElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ShapeComponentElement item = new ShapeComponentElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParameterSet = ParameterSetElement.Create(nav.SelectSingleNode("PARAMETERSET"), item);
            item.RotationInfo = RotationInfoElement.Create(nav.SelectSingleNode("ROTATIONINFO"), item);
            item.RenderingInfo = RenderingInfoElement.Create(nav.SelectSingleNode("RENDERINGINFO"), item);

            item.HRef = String.Concat(nav.SelectSingleNode("@HRef"));

            try { item.XPos = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@XPos")))); }
            catch { }

            try { item.YPos = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@YPos")))); }
            catch { }

            try { item.GroupLevel = Int32.Parse(String.Concat(nav.SelectSingleNode("@GroupLevel"))); }
            catch { }

            try { item.OriWidth = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@OriWidth")))); }
            catch { }

            try { item.OriHeight = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@OriHeight")))); }
            catch { }

            try { item.CurWidth = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@CurWidth")))); }
            catch { }

            try { item.CurHeight = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@CurHeight")))); }
            catch { }

            try { item.HorzFlip = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HorzFlip"))); }
            catch { }

            try { item.VertFlip = Boolean.Parse(String.Concat(nav.SelectSingleNode("@VertFlip"))); }
            catch { }

            item.InstID = String.Concat(nav.SelectSingleNode("@InstID"));

            return item;
        }
    }

    public sealed class RotationInfoElement : IHwpmlElement, IHwpmlElement<ShapeComponentElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ShapeComponentElement)value; }
        }

        public string ElementName { get { return "ROTATIONINFO"; } }
        public ShapeComponentElement Parent { get; set; }

        public double Angle;
        public int CenterX;
        public int CenterY;

        public static RotationInfoElement Create(IXPathNavigable target, ShapeComponentElement parent)
        {
            if (target == null)
                return null;

            RotationInfoElement item = new RotationInfoElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Angle = Double.Parse(String.Concat(nav.SelectSingleNode("@Angle"))); }
            catch { }

            try { item.CenterX = Int32.Parse(String.Concat(nav.SelectSingleNode("@CenterX"))); }
            catch { }

            try { item.CenterY = Int32.Parse(String.Concat(nav.SelectSingleNode("@CenterY"))); }
            catch { }

            return item;
        }
    }

    public sealed class RenderingInfoElement : IHwpmlElement, IHwpmlElement<ShapeComponentElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ShapeComponentElement)value; }
        }

        public string ElementName { get { return "RENDERINGINFO"; } }
        public ShapeComponentElement Parent { get; set; }

        public TransMatrixElement TransMatrix;
        public ScaMatrixElement ScaMatrix;
        public RotMatrixElement RotMatrix;

        public static RenderingInfoElement Create(IXPathNavigable target, ShapeComponentElement parent)
        {
            if (target == null)
                return null;

            RenderingInfoElement item = new RenderingInfoElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.TransMatrix = TransMatrixElement.Create(nav.SelectSingleNode("TRANSMATRIX"), item);
            item.ScaMatrix = ScaMatrixElement.Create(nav.SelectSingleNode("SCAMATRIX"), item);
            item.RotMatrix = RotMatrixElement.Create(nav.SelectSingleNode("ROTMATRIX"), item);

            return item;
        }
    }

    public sealed class TransMatrixElement : IHwpmlElement, IHwpmlElement<RenderingInfoElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (RenderingInfoElement)value; }
        }

        public string ElementName { get { return "TRANSMATRIX"; } }
        public RenderingInfoElement Parent { get; set; }

        public string E1;
        public string E2;
        public string E3;
        public string E4;
        public string E5;
        public string E6;
        public string E7;
        public string E8;
        public string E9;

        public static TransMatrixElement Create(IXPathNavigable target, RenderingInfoElement parent)
        {
            if (target == null)
                return null;

            TransMatrixElement item = new TransMatrixElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.E1 = String.Concat(nav.SelectSingleNode("@E1"));
            item.E2 = String.Concat(nav.SelectSingleNode("@E2"));
            item.E3 = String.Concat(nav.SelectSingleNode("@E3"));
            item.E4 = String.Concat(nav.SelectSingleNode("@E4"));
            item.E5 = String.Concat(nav.SelectSingleNode("@E5"));
            item.E6 = String.Concat(nav.SelectSingleNode("@E6"));
            item.E7 = String.Concat(nav.SelectSingleNode("@E7"));
            item.E8 = String.Concat(nav.SelectSingleNode("@E8"));
            item.E9 = String.Concat(nav.SelectSingleNode("@E9"));

            return item;
        }
    }

    public sealed class ScaMatrixElement : IHwpmlElement, IHwpmlElement<RenderingInfoElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (RenderingInfoElement)value; }
        }

        public string ElementName { get { return "SCAMATRIX"; } }
        public RenderingInfoElement Parent { get; set; }

        public string E1;
        public string E2;
        public string E3;
        public string E4;
        public string E5;
        public string E6;
        public string E7;
        public string E8;
        public string E9;

        public static ScaMatrixElement Create(IXPathNavigable target, RenderingInfoElement parent)
        {
            if (target == null)
                return null;

            ScaMatrixElement item = new ScaMatrixElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.E1 = String.Concat(nav.SelectSingleNode("@E1"));
            item.E2 = String.Concat(nav.SelectSingleNode("@E2"));
            item.E3 = String.Concat(nav.SelectSingleNode("@E3"));
            item.E4 = String.Concat(nav.SelectSingleNode("@E4"));
            item.E5 = String.Concat(nav.SelectSingleNode("@E5"));
            item.E6 = String.Concat(nav.SelectSingleNode("@E6"));
            item.E7 = String.Concat(nav.SelectSingleNode("@E7"));
            item.E8 = String.Concat(nav.SelectSingleNode("@E8"));
            item.E9 = String.Concat(nav.SelectSingleNode("@E9"));

            return item;
        }
    }

    public sealed class RotMatrixElement : IHwpmlElement, IHwpmlElement<RenderingInfoElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (RenderingInfoElement)value; }
        }

        public string ElementName { get { return "ROTMATRIX"; } }
        public RenderingInfoElement Parent { get; set; }

        public string E1;
        public string E2;
        public string E3;
        public string E4;
        public string E5;
        public string E6;
        public string E7;
        public string E8;
        public string E9;

        public static RotMatrixElement Create(IXPathNavigable target, RenderingInfoElement parent)
        {
            if (target == null)
                return null;

            RotMatrixElement item = new RotMatrixElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.E1 = String.Concat(nav.SelectSingleNode("@E1"));
            item.E2 = String.Concat(nav.SelectSingleNode("@E2"));
            item.E3 = String.Concat(nav.SelectSingleNode("@E3"));
            item.E4 = String.Concat(nav.SelectSingleNode("@E4"));
            item.E5 = String.Concat(nav.SelectSingleNode("@E5"));
            item.E6 = String.Concat(nav.SelectSingleNode("@E6"));
            item.E7 = String.Concat(nav.SelectSingleNode("@E7"));
            item.E8 = String.Concat(nav.SelectSingleNode("@E8"));
            item.E9 = String.Concat(nav.SelectSingleNode("@E9"));

            return item;
        }
    }

    public enum LineShapeEndCap : byte
    {
        Round,
        Flat
    }

    public enum LineShapeOutlineStyle : byte
    {
        Normal,
        Outer,
        Inner
    }

    public sealed class LineShapeElement : IHwpmlElement, IHwpmlElement<PictureElement>, IHwpmlElement<DrawingObjectElement>, IHwpmlElement<OleElement>
    {
        public LineShapeElement()
            : base()
        {
            this.Style = LineType1.Solid;
            this.EndCap = LineShapeEndCap.Flat;
            this.HeadStyle = ArrowType.Normal;
            this.TailStyle = ArrowType.Normal;
            this.HeadSize = ArrowSize.SmallSmall;
            this.TailSize = ArrowSize.SmallSmall;
            this.OutlineStyle = LineShapeOutlineStyle.Normal;
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as PictureElement) != null)
                    return;
                if ((this.parent = value as DrawingObjectElement) != null)
                    return;
                if ((this.parent = value as OleElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        PictureElement IHwpmlElement<PictureElement>.Parent
        {
            get { return this.parent as PictureElement; }
            set { this.parent = value; }
        }

        DrawingObjectElement IHwpmlElement<DrawingObjectElement>.Parent
        {
            get { return this.parent as DrawingObjectElement; }
            set { this.parent = value; }
        }

        OleElement IHwpmlElement<OleElement>.Parent
        {
            get { return this.parent as OleElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "LINESHAPE"; } }

        public RGBColor Color;
        public HwpUnit Width;
        public LineType1 Style;
        public LineShapeEndCap EndCap;
        public ArrowType HeadStyle;
        public ArrowType TailStyle;
        public ArrowSize HeadSize;
        public ArrowSize TailSize;
        public LineShapeOutlineStyle OutlineStyle;
        public string Alpha;

        public static LineShapeElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            LineShapeElement item = new LineShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Color = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@Color")))); }
            catch { }

            try { item.Width = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Width")))); }
            catch { }

            try { item.Style = (LineType1)Enum.Parse(typeof(LineType1), String.Concat(nav.SelectSingleNode("@Style")), true); }
            catch { }

            try { item.EndCap = (LineShapeEndCap)Enum.Parse(typeof(LineShapeEndCap), String.Concat(nav.SelectSingleNode("@EndCap")), true); }
            catch { }

            try { item.HeadStyle = (ArrowType)Enum.Parse(typeof(ArrowType), String.Concat(nav.SelectSingleNode("@HeadStyle")), true); }
            catch { }

            try { item.TailStyle = (ArrowType)Enum.Parse(typeof(ArrowType), String.Concat(nav.SelectSingleNode("@TailStyle")), true); }
            catch { }

            try { item.HeadSize = (ArrowSize)Enum.Parse(typeof(ArrowSize), String.Concat(nav.SelectSingleNode("@HeadSize")), true); }
            catch { }

            try { item.TailSize = (ArrowSize)Enum.Parse(typeof(ArrowSize), String.Concat(nav.SelectSingleNode("@TailSize")), true); }
            catch { }

            try { item.OutlineStyle = (LineShapeOutlineStyle)Enum.Parse(typeof(LineShapeOutlineStyle), String.Concat(nav.SelectSingleNode("@OutlineStyle")), true); }
            catch { }

            item.Alpha = String.Concat(nav.SelectSingleNode("@Alpha"));

            return item;
        }
    }

    public sealed class ImageRectElement : IHwpmlElement, IHwpmlElement<PictureElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (PictureElement)value; }
        }

        public string ElementName { get { return "IMAGERECT"; } }
        public PictureElement Parent { get; set; }

        public int X0;
        public int Y0;
        public int X1;
        public int Y1;
        public int X2;
        public int Y2;

        public static ImageRectElement Create(IXPathNavigable target, PictureElement parent)
        {
            if (target == null)
                return null;

            ImageRectElement item = new ImageRectElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.X0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X0"))); }
            catch { }

            try { item.Y0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y0"))); }
            catch { }

            try { item.X1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X1"))); }
            catch { }

            try { item.Y1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y1"))); }
            catch { }

            try { item.X2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X2"))); }
            catch { }

            try { item.Y2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y2"))); }
            catch { }

            return item;
        }
    }

    public sealed class ImageClipElement : IHwpmlElement, IHwpmlElement<PictureElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (PictureElement)value; }
        }

        public string ElementName { get { return "IMAGECLIP"; } }
        public PictureElement Parent { get; set; }

        public int Left;
        public int Right;
        public int Top;
        public int Bottom;

        public static ImageClipElement Create(IXPathNavigable target, PictureElement parent)
        {
            if (target == null)
                return null;

            ImageClipElement item = new ImageClipElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Left = Int32.Parse(String.Concat(nav.SelectSingleNode("@Left"))); }
            catch { }

            try { item.Right = Int32.Parse(String.Concat(nav.SelectSingleNode("@Right"))); }
            catch { }

            try { item.Top = Int32.Parse(String.Concat(nav.SelectSingleNode("@Top"))); }
            catch { }

            try { item.Bottom = Int32.Parse(String.Concat(nav.SelectSingleNode("@Bottom"))); }
            catch { }

            return item;
        }
    }

    public sealed class EffectsElement : IHwpmlElement, IHwpmlElement<PictureElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (PictureElement)value; }
        }

        public string ElementName { get { return "EFFECTS"; } }
        public PictureElement Parent { get; set; }

        public ShadowEffectElement ShadowEffect;
        public GlowElement Glow;
        public SoftEdgeElement SoftEdge;
        public ReflectionElement Reflection;

        public static EffectsElement Create(IXPathNavigable target, PictureElement parent)
        {
            if (target == null)
                return null;

            EffectsElement item = new EffectsElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShadowEffect = ShadowEffectElement.Create(nav.SelectSingleNode("SHADOWEFFECT"), item);
            item.Glow = GlowElement.Create(nav.SelectSingleNode("GLOW"), item);
            item.SoftEdge = SoftEdgeElement.Create(nav.SelectSingleNode("SOFTEDGE"), item);
            item.Reflection = ReflectionElement.Create(nav.SelectSingleNode("REFLECTION"), item);

            return item;
        }
    }

    public sealed class ShadowEffectElement : IHwpmlElement, IHwpmlElement<EffectsElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (EffectsElement)value; }
        }

        public string ElementName { get { return "SHADOWEFFECT"; } }
        public EffectsElement Parent { get; set; }

        public EffectsColorElement EffectsColor;

        public string Style;
        public string Alpha;
        public double Radius;
        public double Direction;
        public double Distance;
        public string AlignStyle;
        public int SkewX;
        public int SkewY;
        public int ScaleX;
        public int ScaleY;
        public string RotationStyle;

        public static ShadowEffectElement Create(IXPathNavigable target, EffectsElement parent)
        {
            if (target == null)
                return null;

            ShadowEffectElement item = new ShadowEffectElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.EffectsColor = EffectsColorElement.Create(nav.SelectSingleNode("EFFECTSCOLOR"), item);

            item.Style = String.Concat(nav.SelectSingleNode("@Style"));
            item.Alpha = String.Concat(nav.SelectSingleNode("@Alpha"));

            try { item.Radius = Double.Parse(String.Concat(nav.SelectSingleNode("@Radius"))); }
            catch { }

            try { item.Direction = Double.Parse(String.Concat(nav.SelectSingleNode("@Direction"))); }
            catch { }

            try { item.Distance = Double.Parse(String.Concat(nav.SelectSingleNode("@Distance"))); }
            catch { }

            item.AlignStyle = String.Concat(nav.SelectSingleNode("@AlignStyle"));

            try { item.SkewX = Int32.Parse(String.Concat(nav.SelectSingleNode("@SkewX"))); }
            catch { }

            try { item.SkewY = Int32.Parse(String.Concat(nav.SelectSingleNode("@SkewY"))); }
            catch { }

            try { item.ScaleX = Int32.Parse(String.Concat(nav.SelectSingleNode("@ScaleX"))); }
            catch { }

            try { item.ScaleY = Int32.Parse(String.Concat(nav.SelectSingleNode("@ScaleY"))); }
            catch { }

            item.RotationStyle = String.Concat(nav.SelectSingleNode("@RotationStyle"));

            return item;
        }
    }

    public sealed class GlowElement : IHwpmlElement, IHwpmlElement<EffectsElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (EffectsElement)value; }
        }

        public string ElementName { get { return "GLOW"; } }
        public EffectsElement Parent { get; set; }

        public EffectsColorElement EffectsColor;

        public string Alpha;
        public string Radius;

        public static GlowElement Create(IXPathNavigable target, EffectsElement parent)
        {
            if (target == null)
                return null;

            GlowElement item = new GlowElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.EffectsColor = EffectsColorElement.Create(nav.SelectSingleNode("EFFECTSCOLOR"), item);

            item.Alpha = String.Concat(nav.SelectSingleNode("@Alpha"));
            item.Radius = String.Concat(nav.SelectSingleNode("@Radius"));

            return item;
        }
    }

    public sealed class SoftEdgeElement : IHwpmlElement, IHwpmlElement<EffectsElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (EffectsElement)value; }
        }

        public string ElementName { get { return "SOFTEDGE"; } }
        public EffectsElement Parent { get; set; }

        public string Radius;

        public static SoftEdgeElement Create(IXPathNavigable target, EffectsElement parent)
        {
            if (target == null)
                return null;

            SoftEdgeElement item = new SoftEdgeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Radius = String.Concat(nav.SelectSingleNode("@Radius"));

            return item;
        }
    }

    public sealed class ReflectionElement : IHwpmlElement, IHwpmlElement<EffectsElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (EffectsElement)value; }
        }

        public string ElementName { get { return "REFLECTION"; } }
        public EffectsElement Parent { get; set; }

        public string AlignStyle;
        public string Radius;
        public string Direction;
        public string Distance;
        public string SkewX;
        public string SkewY;
        public string ScaleX;
        public string ScaleY;
        public string RotationStyle;
        public string StartAlpha;
        public string StartPos;
        public string EndAlpha;
        public string EndPos;
        public string FadeDirection;

        public static ReflectionElement Create(IXPathNavigable target, EffectsElement parent)
        {
            if (target == null)
                return null;

            ReflectionElement item = new ReflectionElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.AlignStyle = String.Concat(nav.SelectSingleNode("@AlignStyle"));
            item.Radius = String.Concat(nav.SelectSingleNode("@Radius"));
            item.Direction = String.Concat(nav.SelectSingleNode("@Direction"));
            item.Distance = String.Concat(nav.SelectSingleNode("@Distance"));
            item.SkewX = String.Concat(nav.SelectSingleNode("@SkewX"));
            item.SkewY = String.Concat(nav.SelectSingleNode("@SkewY"));
            item.ScaleX = String.Concat(nav.SelectSingleNode("@ScaleX"));
            item.ScaleY = String.Concat(nav.SelectSingleNode("@ScaleY"));
            item.RotationStyle = String.Concat(nav.SelectSingleNode("@RotationStyle"));
            item.StartAlpha = String.Concat(nav.SelectSingleNode("@StartAlpha"));
            item.StartPos = String.Concat(nav.SelectSingleNode("@StartPos"));
            item.EndAlpha = String.Concat(nav.SelectSingleNode("@EndAlpha"));
            item.EndPos = String.Concat(nav.SelectSingleNode("@EndPos"));
            item.FadeDirection = String.Concat(nav.SelectSingleNode("@FadeDirection"));

            return item;
        }
    }

    public sealed class EffectsColorElement : IHwpmlElement, IHwpmlElement<ShadowEffectElement>, IHwpmlElement<GlowElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = null; return; }
                if ((this.parent = value as ShadowEffectElement) != null)
                    return;
                if ((this.parent = value as GlowElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        ShadowEffectElement IHwpmlElement<ShadowEffectElement>.Parent
        {
            get { return this.parent as ShadowEffectElement; }
            set { this.parent = value; }
        }

        GlowElement IHwpmlElement<GlowElement>.Parent
        {
            get { return this.parent as GlowElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "EFFECTSCOLOR"; } }

        public ColorEffectElement ColorEffect;

        public string Type;
        public string SchemeIndex;
        public string SystemIndex;
        public string PresetIndex;
        public string ColorR;
        public string ColorG;
        public string ColorB;
        public string ColorC;
        public string ColorM;
        public string ColorY;
        public string ColorK;
        public string ColorSCR;
        public string ColorSCG;
        public string ColorSCB;
        public string ColorH;
        public string ColorS;
        public string ColorL;

        public static EffectsColorElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            EffectsColorElement item = new EffectsColorElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ColorEffect = ColorEffectElement.Create(nav.SelectSingleNode("COLOREFFECT"), item);

            item.Type = String.Concat(nav.SelectSingleNode("@Type"));
            item.SchemeIndex = String.Concat(nav.SelectSingleNode("@SchemeIndex"));
            item.SystemIndex = String.Concat(nav.SelectSingleNode("@SystemIndex"));
            item.PresetIndex = String.Concat(nav.SelectSingleNode("@PresetIndex"));
            item.ColorR = String.Concat(nav.SelectSingleNode("@ColorR"));
            item.ColorG = String.Concat(nav.SelectSingleNode("@ColorG"));
            item.ColorB = String.Concat(nav.SelectSingleNode("@ColorB"));
            item.ColorC = String.Concat(nav.SelectSingleNode("@ColorC"));
            item.ColorM = String.Concat(nav.SelectSingleNode("@ColorM"));
            item.ColorY = String.Concat(nav.SelectSingleNode("@ColorY"));
            item.ColorK = String.Concat(nav.SelectSingleNode("@ColorK"));
            item.ColorSCR = String.Concat(nav.SelectSingleNode("@ColorSCR"));
            item.ColorSCG = String.Concat(nav.SelectSingleNode("@ColorSCG"));
            item.ColorSCB = String.Concat(nav.SelectSingleNode("@ColorSCB"));
            item.ColorH = String.Concat(nav.SelectSingleNode("@ColorH"));
            item.ColorS = String.Concat(nav.SelectSingleNode("@ColorS"));
            item.ColorL = String.Concat(nav.SelectSingleNode("@ColorL"));

            return item;
        }
    }

    public sealed class ColorEffectElement : IHwpmlElement, IHwpmlElement<EffectsColorElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (EffectsColorElement)value; }
        }

        public string ElementName { get { return "COLOREFFECT"; } }
        public EffectsColorElement Parent { get; set; }

        public string Type;
        public string Value;

        public static ColorEffectElement Create(IXPathNavigable target, EffectsColorElement parent)
        {
            if (target == null)
                return null;

            ColorEffectElement item = new ColorEffectElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Type = String.Concat(nav.SelectSingleNode("@Type"));
            item.Value = String.Concat(nav.SelectSingleNode("@Value"));

            return item;
        }
    }

    public sealed class DrawingObjectElement : IHwpmlElement, IHwpmlElement<LineElement>, IHwpmlElement<RectangleElement>, IHwpmlElement<EllipseElement>, IHwpmlElement<ArcElement>, IHwpmlElement<PolygonElement>, IHwpmlElement<CurveElement>, IHwpmlElement<ConnectLineElement>, IHwpmlElement<UnknownObjectElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as LineElement) != null)
                    return;
                if ((this.parent = value as RectangleElement) != null)
                    return;
                if ((this.parent = value as EllipseElement) != null)
                    return;
                if ((this.parent = value as ArcElement) != null)
                    return;
                if ((this.parent = value as PolygonElement) != null)
                    return;
                if ((this.parent = value as CurveElement) != null)
                    return;
                if ((this.parent = value as ConnectLineElement) != null)
                    return;
                if ((this.parent = value as UnknownObjectElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        LineElement IHwpmlElement<LineElement>.Parent
        {
            get { return this.parent as LineElement; }
            set { this.parent = value; }
        }

        RectangleElement IHwpmlElement<RectangleElement>.Parent
        {
            get { return this.parent as RectangleElement; }
            set { this.parent = value; }
        }

        EllipseElement IHwpmlElement<EllipseElement>.Parent
        {
            get { return this.parent as EllipseElement; }
            set { this.parent = value; }
        }

        ArcElement IHwpmlElement<ArcElement>.Parent
        {
            get { return this.parent as ArcElement; }
            set { this.parent = value; }
        }

        PolygonElement IHwpmlElement<PolygonElement>.Parent
        {
            get { return this.parent as PolygonElement; }
            set { this.parent = value; }
        }

        CurveElement IHwpmlElement<CurveElement>.Parent
        {
            get { return this.parent as CurveElement; }
            set { this.parent = value; }
        }

        ConnectLineElement IHwpmlElement<ConnectLineElement>.Parent
        {
            get { return this.parent as ConnectLineElement; }
            set { this.parent = value; }
        }

        UnknownObjectElement IHwpmlElement<UnknownObjectElement>.Parent
        {
            get { return this.parent as UnknownObjectElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "DRAWINGOBJECT"; } }

        public ShapeComponentElement ShapeComponent;
        public LineShapeElement LineShape;
        public FillBrushElement FillBrush;
        public DrawTextElement DrawText;
        public ShadowElement Shadow;

        private static DrawingObjectElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            DrawingObjectElement item = new DrawingObjectElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeComponent = ShapeComponentElement.Create(nav.SelectSingleNode("SHAPECOMPONENT"), item);
            item.LineShape = LineShapeElement.Create(nav.SelectSingleNode("LINESHAPE"), item);
            item.FillBrush = FillBrushElement.Create(nav.SelectSingleNode("FILLBRUSH"), item);
            item.DrawText = DrawTextElement.Create(nav.SelectSingleNode("DRAWTEXT"), item);
            item.Shadow = ShadowElement.Create(nav.SelectSingleNode("SHADOW"), item);

            return item;
        }

        public static DrawingObjectElement Create(IXPathNavigable target, LineElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static DrawingObjectElement Create(IXPathNavigable target, RectangleElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static DrawingObjectElement Create(IXPathNavigable target, EllipseElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static DrawingObjectElement Create(IXPathNavigable target, ArcElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static DrawingObjectElement Create(IXPathNavigable target, PolygonElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static DrawingObjectElement Create(IXPathNavigable target, CurveElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static DrawingObjectElement Create(IXPathNavigable target, UnknownObjectElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static DrawingObjectElement Create(IXPathNavigable target, ConnectLineElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public sealed class DrawTextElement : IHwpmlElement, IHwpmlElement<DrawingObjectElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (DrawingObjectElement)value; }
        }

        public string ElementName { get { return "DRAWTEXT"; } }
        public DrawingObjectElement Parent { get; set; }

        public TextMarginElement TextMargin;
        public ParaListElement ParaList;

        public int LastWidth;
        public string Name;
        public bool Editable;

        public static DrawTextElement Create(IXPathNavigable target, DrawingObjectElement parent)
        {
            if (target == null)
                return null;

            DrawTextElement item = new DrawTextElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.TextMargin = TextMarginElement.Create(nav.SelectSingleNode("TEXTMARGIN"), item);
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            try { item.LastWidth = Int32.Parse(String.Concat(nav.SelectSingleNode("@LastWidth"))); }
            catch { }

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));

            try { item.Editable = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Editable"))); }
            catch { }

            return item;
        }
    }
    
    public sealed class TextMarginElement : IHwpmlElement, IHwpmlElement<DrawTextElement>
    {
        public TextMarginElement()
            : base()
        {
            this.Left = new HwpUnit(238);
            this.Right = new HwpUnit(238);
            this.Top = new HwpUnit(238);
            this.Bottom = new HwpUnit(238);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (DrawTextElement)value; }
        }

        public string ElementName { get { return "TEXTMARGIN"; } }
        public DrawTextElement Parent { get; set; }

        public HwpUnit Left;
        public HwpUnit Right;
        public HwpUnit Top;
        public HwpUnit Bottom;

        public static TextMarginElement Create(IXPathNavigable target, DrawTextElement parent)
        {
            if (target == null)
                return null;

            TextMarginElement item = new TextMarginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Left = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Left")))); }
            catch { }

            try { item.Right = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Right")))); }
            catch { }

            try { item.Top = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Top")))); }
            catch { }

            try { item.Bottom = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@Bottom")))); }
            catch { }

            return item;
        }
    }

    public sealed class LineElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "LINE"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        public HwpUnit StartX;
        public HwpUnit StartY;
        public HwpUnit EndX;
        public HwpUnit EndY;
        public bool IsReverseHV;

        private static LineElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            LineElement item = new LineElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            try { item.StartX = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@StartX")))); }
            catch { }

            try { item.StartY = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@StartY")))); }
            catch { }

            try { item.EndX = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@EndX")))); }
            catch { }

            try { item.EndY = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@EndY")))); }
            catch { }

            try { item.IsReverseHV = Boolean.Parse(String.Concat(nav.SelectSingleNode("@IsReverseHV"))); }
            catch { }

            return item;
        }

        public static LineElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static LineElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public sealed class RectangleElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "RECTANGLE"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        public double Ratio;
        public int X0;
        public int Y0;
        public int X1;
        public int Y1;
        public int X2;
        public int Y2;
        public int X3;
        public int Y3;

        private static RectangleElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            RectangleElement item = new RectangleElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            try { item.Ratio = Double.Parse(String.Concat(nav.SelectSingleNode("@Ratio"))); }
            catch { }

            try { item.X0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X0"))); }
            catch { }

            try { item.Y0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y0"))); }
            catch { }

            try { item.X1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X1"))); }
            catch { }

            try { item.Y1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y1"))); }
            catch { }

            try { item.X2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X2"))); }
            catch { }

            try { item.Y2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y2"))); }
            catch { }

            try { item.X3 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X3"))); }
            catch { }

            try { item.Y3 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y3"))); }
            catch { }

            return item;
        }

        public static RectangleElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static RectangleElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public enum EllipseArcType : byte
    {
        Normal,
        Pie,
        Chord
    }

    public sealed class EllipseElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        public EllipseElement()
            : base()
        {
            this.ArcType = EllipseArcType.Normal;
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "ELLIPSE"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        public bool IntervalDirty;
        public bool HasArcProperty;
        public EllipseArcType ArcType;
        public int CenterX;
        public int CenterY;
        public int Axis1X;
        public int Axis1Y;
        public int Axis2X;
        public int Axis2Y;
        public int Start1X;
        public int Start1Y;
        public int End1X;
        public int End1Y;
        public int Start2X;
        public int Start2Y;
        public int End2X;
        public int End2Y;

        private static EllipseElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            EllipseElement item = new EllipseElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            try { item.IntervalDirty = Boolean.Parse(String.Concat(nav.SelectSingleNode("@IntervalDirty"))); }
            catch { }

            try { item.ArcType = (EllipseArcType)Enum.Parse(typeof(EllipseArcType), String.Concat(nav.SelectSingleNode("@ArcType")), true); }
            catch { }

            try { item.CenterX = Int32.Parse(String.Concat(nav.SelectSingleNode("@CenterX"))); }
            catch { }

            try { item.CenterY = Int32.Parse(String.Concat(nav.SelectSingleNode("@CenterY"))); }
            catch { }

            try { item.Axis1X = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis1X"))); }
            catch { }

            try { item.Axis1Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis1Y"))); }
            catch { }

            try { item.Axis2X = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis2X"))); }
            catch { }

            try { item.Axis2Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis2Y"))); }
            catch { }

            try { item.Start1X = Int32.Parse(String.Concat(nav.SelectSingleNode("@Start1X"))); }
            catch { }

            try { item.Start1Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@Start1Y"))); }
            catch { }

            try { item.End1X = Int32.Parse(String.Concat(nav.SelectSingleNode("@End1X"))); }
            catch { }

            try { item.End1Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@End1Y"))); }
            catch { }

            try { item.Start2X = Int32.Parse(String.Concat(nav.SelectSingleNode("@Start2X"))); }
            catch { }

            try { item.Start2Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@CenterX"))); }
            catch { }

            try { item.End2X = Int32.Parse(String.Concat(nav.SelectSingleNode("@End2X"))); }
            catch { }

            try { item.End2Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@End2Y"))); }
            catch { }

            return item;
        }

        public static EllipseElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static EllipseElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public enum ArcType : byte
    {
        Normal,
        Pie,
        Chord
    }

    public sealed class ArcElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        public ArcElement()
            : base()
        {
            this.Type = ArcType.Normal;
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "ARC"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        public ArcType Type;
        public int CenterX;
        public int CenterY;
        public int Axis1X;
        public int Axis1Y;
        public int Axis2X;
        public int Axis2Y;

        private static ArcElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ArcElement item = new ArcElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            try { item.Type = (ArcType)Enum.Parse(typeof(ArcType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.CenterX = Int32.Parse(String.Concat(nav.SelectSingleNode("@CenterX"))); }
            catch { }

            try { item.CenterY = Int32.Parse(String.Concat(nav.SelectSingleNode("@CenterY"))); }
            catch { }

            try { item.Axis1X = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis1X"))); }
            catch { }

            try { item.Axis1Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis1Y"))); }
            catch { }

            try { item.Axis2X = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis2X"))); }
            catch { }

            try { item.Axis2Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@Axis2Y"))); }
            catch { }

            return item;
        }

        public static ArcElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static ArcElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public sealed class PolygonElement : List<PointElement>, IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "POLYGON"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        private static PolygonElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            PolygonElement item = new PolygonElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            foreach (IXPathNavigable each in nav.Select("POINT"))
                item.Add(PointElement.Create(each, item));

            return item;
        }

        public static PolygonElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static PolygonElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public sealed class PointElement : IHwpmlElement, IHwpmlElement<PolygonElement>, IHwpmlElement<OutlineDataElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as PolygonElement) != null)
                    return;
                if ((this.parent = value as OutlineDataElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        PolygonElement IHwpmlElement<PolygonElement>.Parent
        {
            get { return this.parent as PolygonElement; }
            set { this.parent = value; }
        }

        OutlineDataElement IHwpmlElement<OutlineDataElement>.Parent
        {
            get { return this.parent as OutlineDataElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "POINT"; } }

        public int X;
        public int Y;

        private static PointElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            PointElement item = new PointElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.X = Int32.Parse(String.Concat(nav.SelectSingleNode("@X"))); }
            catch { }

            try { item.Y = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y"))); }
            catch { }

            return item;
        }

        public static PointElement Create(IXPathNavigable target, PolygonElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static PointElement Create(IXPathNavigable target, OutlineDataElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public sealed class CurveElement : List<SegmentElement>, IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "CURVE"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        private static CurveElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            CurveElement item = new CurveElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            foreach (IXPathNavigable each in nav.Select("SEGMENT"))
                item.Add(SegmentElement.Create(each, item));

            return item;
        }

        public static CurveElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static CurveElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public enum SegmentType : byte
    {
        Line,
        Curve
    }
    
    public sealed class SegmentElement : IHwpmlElement, IHwpmlElement<CurveElement>
    {
        public SegmentElement()
            : base()
        {
            this.Type = SegmentType.Curve;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CurveElement)value; }
        }

        public string ElementName { get { return "SEGMENT"; } }
        public CurveElement Parent { get; set; }

        public SegmentType Type;
        public int X1;
        public int Y1;
        public int X2;
        public int Y2;

        public static SegmentElement Create(IXPathNavigable target, CurveElement parent)
        {
            if (target == null)
                return null;

            SegmentElement item = new SegmentElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (SegmentType)Enum.Parse(typeof(SegmentType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.X1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X1"))); }
            catch { }

            try { item.Y1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y1"))); }
            catch { }

            try { item.X2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X2"))); }
            catch { }

            try { item.Y2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y2"))); }
            catch { }

            return item;
        }
    }

    public sealed class ConnectLineElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "CONNECTLINE"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        public string Type;
        public int StartX;
        public int StartY;
        public int EndX;
        public int EndY;
        public string StartSubjectID;
        public string StartSubjectIndex;
        public string EndSubjectID;
        public string EndSubjectIndex;

        private static ConnectLineElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ConnectLineElement item = new ConnectLineElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            item.Type = String.Concat(nav.SelectSingleNode("@Type"));

            try { item.StartX = Int32.Parse(String.Concat(nav.SelectSingleNode("@StartX"))); }
            catch { }

            try { item.StartY = Int32.Parse(String.Concat(nav.SelectSingleNode("@StartY"))); }
            catch { }

            try { item.EndX = Int32.Parse(String.Concat(nav.SelectSingleNode("@EndX"))); }
            catch { }

            try { item.EndY = Int32.Parse(String.Concat(nav.SelectSingleNode("@EndY"))); }
            catch { }

            item.StartSubjectID = String.Concat(nav.SelectSingleNode("@StartSubjectID"));
            item.StartSubjectIndex = String.Concat(nav.SelectSingleNode("@StartSubjectIndex"));
            item.EndSubjectID = String.Concat(nav.SelectSingleNode("@EndSubjectID"));
            item.EndSubjectIndex = String.Concat(nav.SelectSingleNode("@EndSubjectIndex"));

            return item;
        }

        public static ConnectLineElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static ConnectLineElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public sealed class UnknownObjectElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public TextElement Parent { get; set; }
        public string ElementName { get { return "UNKNOWNOBJECT"; } }

        public ShapeObjectElement ShapeObject;
        public DrawingObjectElement DrawingObject;

        public string Ctrlid;
        public int X0;
        public int Y0;
        public int X1;
        public int Y1;
        public int X2;
        public int Y2;
        public int X3;
        public int Y3;

        public static UnknownObjectElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            UnknownObjectElement item = new UnknownObjectElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.DrawingObject = DrawingObjectElement.Create(nav.SelectSingleNode("DRAWINGOBJECT"), item);

            item.Ctrlid = String.Concat(nav.SelectSingleNode("@Ctrlid"));

            try { item.X0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X0"))); }
            catch { }

            try { item.Y0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y0"))); }
            catch { }

            try { item.X1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X1"))); }
            catch { }

            try { item.Y1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y1"))); }
            catch { }

            try { item.X2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X2"))); }
            catch { }

            try { item.Y2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y2"))); }
            catch { }

            try { item.X3 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X3"))); }
            catch { }

            try { item.Y3 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y3"))); }
            catch { }

            return item;
        }
    }

    public sealed class FormObjectElement : IHwpmlElement, IHwpmlElement<ButtonElement>, IHwpmlElement<RadioButtonElement>, IHwpmlElement<CheckButtonElement>, IHwpmlElement<ComboBoxElement>, IHwpmlElement<EditElement>, IHwpmlElement<ListBoxElement>, IHwpmlElement<ScrollBarElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as ButtonElement) != null)
                    return;
                if ((this.parent = value as RadioButtonElement) != null)
                    return;
                if ((this.parent = value as CheckButtonElement) != null)
                    return;
                if ((this.parent = value as ComboBoxElement) != null)
                    return;
                if ((this.parent = value as EditElement) != null)
                    return;
                if ((this.parent = value as ListBoxElement) != null)
                    return;
                if ((this.parent = value as ScrollBarElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        ButtonElement IHwpmlElement<ButtonElement>.Parent
        {
            get { return this.parent as ButtonElement; }
            set { this.parent = value; }
        }

        RadioButtonElement IHwpmlElement<RadioButtonElement>.Parent
        {
            get { return this.parent as RadioButtonElement; }
            set { this.parent = value; }
        }

        CheckButtonElement IHwpmlElement<CheckButtonElement>.Parent
        {
            get { return this.parent as CheckButtonElement; }
            set { this.parent = value; }
        }

        ComboBoxElement IHwpmlElement<ComboBoxElement>.Parent
        {
            get { return this.parent as ComboBoxElement; }
            set { this.parent = value; }
        }

        EditElement IHwpmlElement<EditElement>.Parent
        {
            get { return this.parent as EditElement; }
            set { this.parent = value; }
        }

        ListBoxElement IHwpmlElement<ListBoxElement>.Parent
        {
            get { return this.parent as ListBoxElement; }
            set { this.parent = value; }
        }

        ScrollBarElement IHwpmlElement<ScrollBarElement>.Parent
        {
            get { return this.parent as ScrollBarElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "FORMOBJECT"; } }

        public ParameterSetElement ParameterSet;
        public FormCharShapeElement FormCharShape;
        public ButtonSetElement ButtonSet;

        public string Name;
        public string ForeColor;
        public string BackColor;
        public string GroupName;
        public bool TabStop;
        public string TabOrder;
        public bool Enabled;
        public int BorderType;
        public bool DrawFrame;
        public bool Printable;

        public static FormObjectElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            FormObjectElement item = new FormObjectElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParameterSet = ParameterSetElement.Create(nav.SelectSingleNode("PARAMETERSET"), item);
            item.FormCharShape = FormCharShapeElement.Create(nav.SelectSingleNode("FORMCHARSHAPE"), item);
            item.ButtonSet = ButtonSetElement.Create(nav.SelectSingleNode("BUTTONSET"), item);

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));
            item.ForeColor = String.Concat(nav.SelectSingleNode("@ForeColor"));
            item.BackColor = String.Concat(nav.SelectSingleNode("@BackColor"));
            item.GroupName = String.Concat(nav.SelectSingleNode("@GroupName"));

            try { item.TabStop = Boolean.Parse(String.Concat(nav.SelectSingleNode("@TabStop"))); }
            catch { }

            item.TabOrder = String.Concat(nav.SelectSingleNode("@TabOrder"));

            try { item.Enabled = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Enabled"))); }
            catch { }

            try { item.BorderType = Int32.Parse(String.Concat(nav.SelectSingleNode("@BorderType"))); }
            catch { }

            try { item.DrawFrame = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DrawFrame"))); }
            catch { }

            try { item.Printable = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Printable"))); }
            catch { }

            return item;
        }
    }

    public sealed class FormCharShapeElement : IHwpmlElement, IHwpmlElement<FormObjectElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FormObjectElement)value; }
        }

        public string ElementName { get { return "FORMCHARSHAPE"; } }
        public FormObjectElement Parent { get; set; }

        public int CharShape;
        public bool FollowContext;
        public bool AutoSize;
        public bool WordWrap;

        public static FormCharShapeElement Create(IXPathNavigable target, FormObjectElement parent)
        {
            if (target == null)
                return null;

            FormCharShapeElement item = new FormCharShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.CharShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@CharShape"))); }
            catch { }

            try { item.FollowContext = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FollowContext"))); }
            catch { }

            try { item.AutoSize = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AutoSize"))); }
            catch { }

            try { item.WordWrap = Boolean.Parse(String.Concat(nav.SelectSingleNode("@WordWrap"))); }
            catch { }

            return item;
        }
    }

    public sealed class ButtonSetElement : IHwpmlElement, IHwpmlElement<FormObjectElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (FormObjectElement)value; }
        }

        public string ElementName { get { return "BUTTONSET"; } }
        public FormObjectElement Parent { get; set; }

        public string Caption;
        public string Value;
        public string RadioGroupName;
        public bool TriState;
        public string BackStyle;

        public static ButtonSetElement Create(IXPathNavigable target, FormObjectElement parent)
        {
            if (target == null)
                return null;

            ButtonSetElement item = new ButtonSetElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Caption = String.Concat(nav.SelectSingleNode("@Caption"));
            item.Value = String.Concat(nav.SelectSingleNode("@Value"));
            item.RadioGroupName = String.Concat(nav.SelectSingleNode("@RadioGroupName"));
            
            try { item.TriState = Boolean.Parse(String.Concat(nav.SelectSingleNode("@TriState"))); }
            catch { }

            item.BackStyle = String.Concat(nav.SelectSingleNode("@BackStyle"));

            return item;
        }
    }

    public sealed class ButtonElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "BUTTON"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public FormObjectElement FormObject;

        public static ButtonElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            ButtonElement item = new ButtonElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.FormObject = FormObjectElement.Create(nav.SelectSingleNode("FORMOBJECT"), item);

            return item;
        }
    }

    public sealed class RadioButtonElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "RADIOBUTTON"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public FormObjectElement FormObject;

        public static RadioButtonElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            RadioButtonElement item = new RadioButtonElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.FormObject = FormObjectElement.Create(nav.SelectSingleNode("FORMOBJECT"), item);

            return item;
        }
    }

    public sealed class CheckButtonElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "CHECKBUTTON"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public FormObjectElement FormObject;

        public static CheckButtonElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            CheckButtonElement item = new CheckButtonElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.FormObject = FormObjectElement.Create(nav.SelectSingleNode("FORMOBJECT"), item);

            return item;
        }
    }

    public sealed class ComboBoxElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "COMBOBOX"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public FormObjectElement FormObject;

        public int ListBoxRows;
        public int ListBoxWidth;
        public string Text;
        public string EditEnable;

        public static ComboBoxElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            ComboBoxElement item = new ComboBoxElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.FormObject = FormObjectElement.Create(nav.SelectSingleNode("FORMOBJECT"), item);

            try { item.ListBoxRows = Int32.Parse(String.Concat(nav.SelectSingleNode("@ListBoxRows"))); }
            catch { }

            try { item.ListBoxWidth = Int32.Parse(String.Concat(nav.SelectSingleNode("@ListBoxWidth"))); }
            catch { }

            item.Text = String.Concat(nav.SelectSingleNode("@Text"));
            item.EditEnable = String.Concat(nav.SelectSingleNode("@EditEnable"));

            return item;
        }
    }

    public sealed class EditElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "EDIT"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public FormObjectElement FormObject;
        public string EditText;

        public string MultiLine;
        public string PasswordChar;
        public int MaxLength;
        public string ScrollBars;
        public string TabKeyBehavior;
        public bool Number;
        public bool ReadOnly;
        public string AlignText;

        public static EditElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            EditElement item = new EditElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.FormObject = FormObjectElement.Create(nav.SelectSingleNode("FORMOBJECT"), item);
            item.EditText = String.Concat(nav.SelectSingleNode("EDITTEXT"));

            item.MultiLine = String.Concat(nav.SelectSingleNode("@MultiLine"));
            item.PasswordChar = String.Concat(nav.SelectSingleNode("@PasswordChar"));

            try { item.MaxLength = Int32.Parse(String.Concat(nav.SelectSingleNode("@MaxLength"))); }
            catch { }

            item.ScrollBars = String.Concat(nav.SelectSingleNode("@ScrollBars"));
            item.TabKeyBehavior = String.Concat(nav.SelectSingleNode("@TabKeyBehavior"));

            try { item.Number = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Number"))); }
            catch { }

            try { item.ReadOnly = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ReadOnly"))); }
            catch { }

            item.AlignText = String.Concat(nav.SelectSingleNode("@AlignText"));

            return item;
        }
    }

    public sealed class ListBoxElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "LISTBOX"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public FormObjectElement FormObject;

        public string Text;
        public int ItemHeight;
        public int TopIndex;

        public static ListBoxElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            ListBoxElement item = new ListBoxElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.FormObject = FormObjectElement.Create(nav.SelectSingleNode("FORMOBJECT"), item);

            item.Text = String.Concat(nav.SelectSingleNode("@Text"));

            try { item.ItemHeight = Int32.Parse(String.Concat(nav.SelectSingleNode("@ItemHeight"))); }
            catch { }

            try { item.TopIndex = Int32.Parse(String.Concat(nav.SelectSingleNode("@TopIndex"))); }
            catch { }

            return item;
        }
    }

    public sealed class ScrollBarElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "SCROLLBAR"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public FormObjectElement FormObject;

        public int Delay;
        public int LargeChange;
        public int SmallChange;
        public int Min;
        public int Max;
        public int Page;
        public int Value;
        public string Type;

        public static ScrollBarElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            ScrollBarElement item = new ScrollBarElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.FormObject = FormObjectElement.Create(nav.SelectSingleNode("FORMOBJECT"), item);

            try { item.Delay = Int32.Parse(String.Concat(nav.SelectSingleNode("@Delay"))); }
            catch { }

            try { item.LargeChange = Int32.Parse(String.Concat(nav.SelectSingleNode("@LargeChange"))); }
            catch { }

            try { item.SmallChange = Int32.Parse(String.Concat(nav.SelectSingleNode("@SmallChange"))); }
            catch { }

            try { item.Min = Int32.Parse(String.Concat(nav.SelectSingleNode("@Min"))); }
            catch { }

            try { item.Max = Int32.Parse(String.Concat(nav.SelectSingleNode("@Max"))); }
            catch { }

            try { item.Page = Int32.Parse(String.Concat(nav.SelectSingleNode("@Page"))); }
            catch { }

            try { item.Value = Int32.Parse(String.Concat(nav.SelectSingleNode("@Value"))); }
            catch { }

            item.Type = String.Concat(nav.SelectSingleNode("@Type"));

            return item;
        }
    }

    public sealed class ContainerElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        public ContainerElement()
            : base()
        {
            this.Containers = new List<ContainerElement>();
            this.Lines = new List<LineElement>();
            this.Rectangles = new List<RectangleElement>();
            this.Ellipses = new List<EllipseElement>();
            this.Arcs = new List<ArcElement>();
            this.Polygons = new List<PolygonElement>();
            this.Curves = new List<CurveElement>();
            this.ConnectLines = new List<ConnectLineElement>();
            this.Pictures = new List<PictureElement>();
            this.Oles = new List<OleElement>();
        }

        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "CONTAINER"; } }

        public ShapeObjectElement ShapeObject;
        public ShapeComponentElement ShapeComponent;
        public IList<ContainerElement> Containers;
        public IList<LineElement> Lines;
        public IList<RectangleElement> Rectangles;
        public IList<EllipseElement> Ellipses;
        public IList<ArcElement> Arcs;
        public IList<PolygonElement> Polygons;
        public IList<CurveElement> Curves;
        public IList<ConnectLineElement> ConnectLines;
        public IList<PictureElement> Pictures;
        public IList<OleElement> Oles;

        private static ContainerElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            ContainerElement item = new ContainerElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.ShapeComponent = ShapeComponentElement.Create(nav.SelectSingleNode("SHAPECOMPONENT"), item);

            foreach (IXPathNavigable each in nav.Select("CONTAINER"))
                item.Containers.Add(ContainerElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("LINE"))
                item.Lines.Add(LineElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("RECTANGLE"))
                item.Rectangles.Add(RectangleElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("ELLIPSE"))
                item.Ellipses.Add(EllipseElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("ARC"))
                item.Arcs.Add(ArcElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("POLYGON"))
                item.Polygons.Add(PolygonElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("CURVE"))
                item.Curves.Add(CurveElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("CONNECTLINE"))
                item.ConnectLines.Add(ConnectLineElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("PICTURE"))
                item.Pictures.Add(PictureElement.Create(each, item));

            foreach (IXPathNavigable each in nav.Select("OLE"))
                item.Oles.Add(OleElement.Create(each, item));

            return item;
        }

        public static ContainerElement Create(IXPathNavigable target, TextElement parent)
        { return Create(target, (IHwpmlElement)parent); }

        public static ContainerElement Create(IXPathNavigable target, ContainerElement parent)
        { return Create(target, (IHwpmlElement)parent); }
    }

    public enum OleObjetType : byte
    {
        Unknown,
        Embedded,
        Link,
        Static,
        Equation
    }

    public enum OleDrawAspect : byte
    {
        Content,
        ThumbNail,
        Icon,
        DocPrint
    }

    public sealed class OleElement : IHwpmlElement, IHwpmlElement<TextElement>, IHwpmlElement<ContainerElement>
    {
        private IHwpmlElement parent;

        public IHwpmlElement Parent
        {
            get { return this.parent; }
            set
            {
                if (value == null) { this.parent = value; return; }
                if ((this.parent = value as TextElement) != null)
                    return;
                if ((this.parent = value as ContainerElement) != null)
                    return;
                throw new InvalidCastException();
            }
        }

        TextElement IHwpmlElement<TextElement>.Parent
        {
            get { return this.parent as TextElement; }
            set { this.parent = value; }
        }

        ContainerElement IHwpmlElement<ContainerElement>.Parent
        {
            get { return this.parent as ContainerElement; }
            set { this.parent = value; }
        }

        public string ElementName { get { return "OLE"; } }

        public ShapeObjectElement ShapeObject;
        public ShapeComponentElement ShapeComponent;
        public LineShapeElement LineShape;

        public OleObjetType ObjetType;
        public int ExtentX;
        public int ExtentY;
        public string BinItem;
        public OleDrawAspect DrawAspect;
        public bool HasMoniker;
        public string EqBaseLine;

        public static OleElement Create(IXPathNavigable target, IHwpmlElement parent)
        {
            if (target == null)
                return null;

            OleElement item = new OleElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.ShapeComponent = ShapeComponentElement.Create(nav.SelectSingleNode("SHAPECOMPONENT"), item);
            item.LineShape = LineShapeElement.Create(nav.SelectSingleNode("LINESHAPE"), item);

            try { item.ObjetType = (OleObjetType)Enum.Parse(typeof(OleObjetType), String.Concat(nav.SelectSingleNode("@ObjetType")), true); }
            catch { }
            
            try { item.ExtentX = Int32.Parse(String.Concat(nav.SelectSingleNode("@ExtentX"))); }
            catch { }

            try { item.ExtentY = Int32.Parse(String.Concat(nav.SelectSingleNode("@ExtentY"))); }
            catch { }

            item.BinItem = String.Concat(nav.SelectSingleNode("@BinItem"));

            try { item.DrawAspect = (OleDrawAspect)Enum.Parse(typeof(OleDrawAspect), String.Concat(nav.SelectSingleNode("@DrawAspect")), true); }
            catch { }

            try { item.HasMoniker = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HasMoniker"))); }
            catch { }

            item.EqBaseLine = String.Concat(nav.SelectSingleNode("@EqBaseLine"));

            return item;
        }
    }

    public sealed class EquationElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public EquationElement()
            : base()
        {
            this.BaseUnit = new HwpUnit(1000);
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "EQUATION"; } }
        public TextElement Parent { get; set; }

        public ShapeObjectElement ShapeObject;
        public string Script;

        public bool LineMode;
        public HwpUnit BaseUnit;
        public RGBColor TextColor;
        public string BaseLine;
        public string Version;

        public static EquationElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            EquationElement item = new EquationElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ShapeObject = ShapeObjectElement.Create(nav.SelectSingleNode("SHAPEOBJECT"), item);
            item.Script = String.Concat(nav.SelectSingleNode("SCRIPT"));

            try { item.LineMode = Boolean.Parse(String.Concat(nav.SelectSingleNode("@LineMode"))); }
            catch { }

            try { item.BaseUnit = new HwpUnit(Int32.Parse(String.Concat(nav.SelectSingleNode("@BaseUnit")))); }
            catch { }

            try { item.TextColor = new RGBColor(Int64.Parse(String.Concat(nav.SelectSingleNode("@TextColor")))); }
            catch { }

            item.BaseLine = String.Concat(nav.SelectSingleNode("@BaseLine"));
            item.Version = String.Concat(nav.SelectSingleNode("@Version"));

            return item;
        }
    }

    public sealed class TextArtElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "TEXTART"; } }
        public TextElement Parent { get; set; }

        public TextArtShapeElement TextArtShape;
        public OutlineDataElement OutlineData;

        public string Text;
        public int X0;
        public int Y0;
        public int X1;
        public int Y1;
        public int X2;
        public int Y2;
        public int X3;
        public int Y3;

        public static TextArtElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            TextArtElement item = new TextArtElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.TextArtShape = TextArtShapeElement.Create(nav.SelectSingleNode("TEXTARTSHAPE"), item);
            item.OutlineData = OutlineDataElement.Create(nav.SelectSingleNode("OUTLINEDATA"), item);

            item.Text = String.Concat(nav.SelectSingleNode("@Text"));
            
            try { item.X0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X0"))); }
            catch { }

            try { item.Y0 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y0"))); }
            catch { }

            try { item.X1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X1"))); }
            catch { }

            try { item.Y1 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y1"))); }
            catch { }

            try { item.X2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X2"))); }
            catch { }

            try { item.Y2 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y2"))); }
            catch { }

            try { item.X3 = Int32.Parse(String.Concat(nav.SelectSingleNode("@X3"))); }
            catch { }

            try { item.Y3 = Int32.Parse(String.Concat(nav.SelectSingleNode("@Y3"))); }
            catch { }

            return item;
        }
    }

    public enum TextArtShapeFontStyle : byte
    {
        Regular
    }

    public enum TextArtShapeAlign : byte
    {
        Left,
        Right,
        Center,
        Full,
        Table
    }

    public sealed class TextArtShapeElement : IHwpmlElement, IHwpmlElement<TextArtElement>
    {
        public TextArtShapeElement()
            : base()
        {
            this.FontStyle = TextArtShapeFontStyle.Regular;
            this.FontType = TextArtShapeFontType.ttf;
            this.LineSpacing = 120;
            this.CharSpacing = 100;
            this.Align = TextArtShapeAlign.Left;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextArtElement)value; }
        }

        public string ElementName { get { return "TEXTARTSHAPE"; } }
        public TextArtElement Parent { get; set; }

        public ShadowElement Shadow;

        public string FontName;
        public TextArtShapeFontStyle FontStyle;
        public TextArtShapeFontType FontType;
        public int TextShape;
        public int LineSpacing;
        public int CharSpacing;
        public TextArtShapeAlign Align;

        public static TextArtShapeElement Create(IXPathNavigable target, TextArtElement parent)
        {
            if (target == null)
                return null;

            TextArtShapeElement item = new TextArtShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.Shadow = ShadowElement.Create(nav.SelectSingleNode("SHADOW"), item);

            item.FontName = String.Concat(nav.SelectSingleNode("@FontName"));

            try { item.FontStyle = (TextArtShapeFontStyle)Enum.Parse(typeof(TextArtShapeFontStyle), String.Concat(nav.SelectSingleNode("@FontStyle")), true); }
            catch { }

            try { item.FontType = (TextArtShapeFontType)Enum.Parse(typeof(TextArtShapeFontType), String.Concat(nav.SelectSingleNode("@FontType")), true); }
            catch { }

            try { item.TextShape = Int32.Parse(String.Concat(nav.SelectSingleNode("@TextShape"))); }
            catch { }

            try { item.LineSpacing = Int32.Parse(String.Concat(nav.SelectSingleNode("@LineSpacing"))); }
            catch { }

            try { item.CharSpacing = Int32.Parse(String.Concat(nav.SelectSingleNode("@CharSpacing"))); }
            catch { }

            try { item.Align = (TextArtShapeAlign)Enum.Parse(typeof(TextArtShapeAlign), String.Concat(nav.SelectSingleNode("@Align")), true); }
            catch { }

            return item;
        }
    }

    public sealed class OutlineDataElement : List<PointElement>, IHwpmlElement, IHwpmlElement<TextArtElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextArtElement)value; }
        }

        public string ElementName { get { return "OUTLINEDATA"; } }
        public TextArtElement Parent { get; set; }

        public static OutlineDataElement Create(IXPathNavigable target, TextArtElement parent)
        {
            if (target == null)
                return null;

            OutlineDataElement item = new OutlineDataElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("POINT"))
                item.Add(PointElement.Create(each, item));

            return item;
        }
    }

    public sealed class FieldBeginElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public FieldBeginElement()
            : base()
        {
            this.Editable = true;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "FIELDBEGIN"; } }
        public TextElement Parent { get; set; }

        public FieldType Type;
        public string Name;
        public string InstId;
        public bool Editable;
        public bool Dirty;
        public string Property;
        public string Command;

        public static FieldBeginElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            FieldBeginElement item = new FieldBeginElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (FieldType)Enum.Parse(typeof(FieldType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));
            item.InstId = String.Concat(nav.SelectSingleNode("@InstId"));

            try { item.Editable = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Editable"))); }
            catch { }

            try { item.Dirty = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Dirty"))); }
            catch { }

            item.Property = String.Concat(nav.SelectSingleNode("@Property"));
            item.Command = String.Concat(nav.SelectSingleNode("@Command"));

            return item;
        }
    }

    public sealed class FieldEndElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public FieldEndElement()
            : base()
        {
            this.Editable = true;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "FIELDEND"; } }
        public TextElement Parent { get; set; }

        public FieldType Type;
        public bool Editable;
        public string Property;

        public static FieldEndElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            FieldEndElement item = new FieldEndElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Type = (FieldType)Enum.Parse(typeof(FieldType), String.Concat(nav.SelectSingleNode("@Type")), true); }
            catch { }

            try { item.Editable = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Editable"))); }
            catch { }

            item.Property = String.Concat(nav.SelectSingleNode("@Property"));

            return item;
        }
    }

    public sealed class BookmarkElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "BOOKMARK"; } }
        public TextElement Parent { get; set; }

        public string Name;

        public static BookmarkElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            BookmarkElement item = new BookmarkElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.Name = String.Concat(nav.SelectSingleNode("@Name"));

            return item;
        }
    }

    public enum ApplyPageType : byte
    {
        Both,
        Even,
        Odd
    }

    public sealed class HeaderElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public HeaderElement()
            : base()
        {
            this.ApplyPageType = ApplyPageType.Both;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "HEADER"; } }
        public TextElement Parent { get; set; }

        public ParaListElement ParaList;

        public ApplyPageType ApplyPageType;
        public string SeriesNum;

        public static HeaderElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            HeaderElement item = new HeaderElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            try { item.ApplyPageType = (ApplyPageType)Enum.Parse(typeof(ApplyPageType), String.Concat(nav.SelectSingleNode("@ApplyPageType")), true); }
            catch { }

            item.SeriesNum = String.Concat(nav.SelectSingleNode("@SeriesNum"));

            return item;
        }
    }

    public sealed class FooterElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public FooterElement()
            : base()
        {
            this.ApplyPageType = ApplyPageType.Both;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "FOOTER"; } }
        public TextElement Parent { get; set; }

        public ParaListElement ParaList;

        public ApplyPageType ApplyPageType;
        public string SeriesNum;

        public static FooterElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            FooterElement item = new FooterElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            try { item.ApplyPageType = (ApplyPageType)Enum.Parse(typeof(ApplyPageType), String.Concat(nav.SelectSingleNode("@ApplyPageType")), true); }
            catch { }

            item.SeriesNum = String.Concat(nav.SelectSingleNode("@SeriesNum"));

            return item;
        }
    }

    public sealed class FootnoteElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "FOOTNOTE"; } }
        public TextElement Parent { get; set; }

        public ParaListElement ParaList;

        public static FootnoteElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            FootnoteElement item = new FootnoteElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            return item;
        }
    }

    public sealed class EndnoteElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "ENDNOTE"; } }
        public TextElement Parent { get; set; }

        public ParaListElement ParaList;

        public static EndnoteElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            EndnoteElement item = new EndnoteElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            return item;
        }
    }

    public enum NumberType : byte
    {
        Page,
        Footnote,
        Endnote,
        Figure,
        Table,
        Equation,
        TotalPage
    }

    public sealed class AutoNumElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public AutoNumElement()
            : base()
        {
            this.Number = 1;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "AUTONUM"; } }
        public TextElement Parent { get; set; }

        public AutoNumFormatElement AutoNumFormat;

        public int Number;
        public NumberType NumberType;

        public static AutoNumElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            AutoNumElement item = new AutoNumElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.AutoNumFormat = AutoNumFormatElement.Create(nav.SelectSingleNode("AUTONUMFORMAT"), item);

            try { item.Number = Int32.Parse(String.Concat(nav.SelectSingleNode("@Number"))); }
            catch { }

            try { item.NumberType = (NumberType)Enum.Parse(typeof(NumberType), String.Concat(nav.SelectSingleNode("@NumberType")), true); }
            catch { }

            return item;
        }
    }

    public sealed class NewNumElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public NewNumElement()
            : base()
        {
            this.Number = 1;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "NEWNUM"; } }
        public TextElement Parent { get; set; }

        public AutoNumFormatElement AutoNumFormat;

        public int Number;
        public NumberType NumberType;

        public static NewNumElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            NewNumElement item = new NewNumElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.AutoNumFormat = AutoNumFormatElement.Create(nav.SelectSingleNode("AUTONUMFORMAT"), item);

            try { item.Number = Int32.Parse(String.Concat(nav.SelectSingleNode("@Number"))); }
            catch { }

            try { item.NumberType = (NumberType)Enum.Parse(typeof(NumberType), String.Concat(nav.SelectSingleNode("@NumberType")), true); }
            catch { }

            return item;
        }
    }

    public enum PageNumCtrlPageStartsOn : byte
    {
        Both,
        Even,
        Odd
    }

    public sealed class PageNumCtrlElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public PageNumCtrlElement()
            : base()
        {
            this.PageStartsOn = PageNumCtrlPageStartsOn.Both;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "PAGENUMCTRL"; } }
        public TextElement Parent { get; set; }

        public PageNumCtrlPageStartsOn PageStartsOn;

        public static PageNumCtrlElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            PageNumCtrlElement item = new PageNumCtrlElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.PageStartsOn = (PageNumCtrlPageStartsOn)Enum.Parse(typeof(PageNumCtrlPageStartsOn), String.Concat(nav.SelectSingleNode("@PageStartsOn")), true); }
            catch { }

            return item;
        }
    }

    public sealed class PageHidingElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "PAGEHIDING"; } }
        public TextElement Parent { get; set; }

        public bool HideHeader;
        public bool HideFooter;
        public bool HideMasterPage;
        public bool HideBorder;
        public bool HideFill;
        public bool HidePageNum;

        public static PageHidingElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            PageHidingElement item = new PageHidingElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.HideHeader = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HideHeader"))); }
            catch { }

            try { item.HideFooter = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HideFooter"))); }
            catch { }

            try { item.HideMasterPage = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HideMasterPage"))); }
            catch { }

            try { item.HideBorder = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HideBorder"))); }
            catch { }

            try { item.HideFill = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HideFill"))); }
            catch { }

            try { item.HidePageNum = Boolean.Parse(String.Concat(nav.SelectSingleNode("@HidePageNum"))); }
            catch { }

            return item;
        }
    }

    public enum PageNumPos : byte
    {
        None,
        TopLeft,
        TopCenter,
        TopRight,
        BottomLeft,
        BottomCenter,
        BottomRight,
        OutsideTop,
        OutsideBottom,
        InsideTop,
        InsideBottom
    }

    public sealed class PageNumElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public PageNumElement()
            : base()
        {
            this.Pos = PageNumPos.TopLeft;
            this.FormatType = NumberType1.Digit;
            this.SideChar = "-";
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "PAGENUM"; } }
        public TextElement Parent { get; set; }

        public PageNumPos Pos;
        public NumberType1 FormatType;
        public string SideChar;

        public static PageNumElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            PageNumElement item = new PageNumElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.Pos = (PageNumPos)Enum.Parse(typeof(PageNumPos), String.Concat(nav.SelectSingleNode("@Pos")), true); }
            catch { }

            try { item.FormatType = (NumberType1)Enum.Parse(typeof(NumberType1), String.Concat(nav.SelectSingleNode("@FormatType")), true); }
            catch { }

            item.SideChar = String.Concat(nav.SelectSingleNode("@SideChar"));

            return item;
        }
    }

    public sealed class IndexMarkElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "INDEXMARK"; } }
        public TextElement Parent { get; set; }

        public string KeyFirst;
        public string KeySecond;

        public static IndexMarkElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            IndexMarkElement item = new IndexMarkElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.KeyFirst = String.Concat(nav.SelectSingleNode("KEYFIRST"));
            item.KeySecond = String.Concat(nav.SelectSingleNode("KEYSECOND"));

            return item;
        }
    }

    public sealed class ComposeElement : List<CompCharShapeElement>, IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "COMPOSE"; } }
        public TextElement Parent { get; set; }

        public string CircleType;
        public int CharSize;
        public string ComposeType;
        public int CharShapeSize;

        public static ComposeElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            ComposeElement item = new ComposeElement();
            
            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("COMPCHARSHAPE"))
                item.Add(CompCharShapeElement.Create(each, item));

            item.CircleType = String.Concat(nav.SelectSingleNode("@CircleType"));

            try { item.CharSize = Int32.Parse(String.Concat(nav.SelectSingleNode("@CharSize"))); }
            catch { }

            item.ComposeType = String.Concat(nav.SelectSingleNode("@ComposeType"));
            
            try { item.CharShapeSize = Int32.Parse(String.Concat(nav.SelectSingleNode("@CharShapeSize"))); }
            catch { }

            return item;
        }
    }

    public sealed class CompCharShapeElement : IHwpmlElement, IHwpmlElement<ComposeElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ComposeElement)value; }
        }

        public string ElementName { get { return "COMPCHARSHAPE"; } }
        public ComposeElement Parent { get; set; }

        public string ShapeID;

        public static CompCharShapeElement Create(IXPathNavigable target, ComposeElement parent)
        {
            if (target == null)
                return null;

            CompCharShapeElement item = new CompCharShapeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            item.ShapeID = String.Concat(nav.SelectSingleNode("@ShapeID"));

            return item;
        }
    }

    public enum DutmalPosType : byte
    {
        Top,
        Bottom
    }

    public sealed class DutmalElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        public DutmalElement()
            : base()
        {
            this.PosType = DutmalPosType.Top;
            this.Align = AlignmentType1.Center;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "DUTMAL"; } }
        public TextElement Parent { get; set; }

        public string MainText;
        public string SubText;

        public DutmalPosType PosType;
        public string SizeRatio;
        public string Option;
        public string StyleNo;
        public AlignmentType1 Align;

        public static DutmalElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            DutmalElement item = new DutmalElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.MainText = String.Concat(nav.SelectSingleNode("MAINTEXT"));
            item.SubText = String.Concat(nav.SelectSingleNode("SUBTEXT"));

            try { item.PosType = (DutmalPosType)Enum.Parse(typeof(DutmalPosType), String.Concat(nav.SelectSingleNode("@PosType")), true); }
            catch { }

            item.SizeRatio = String.Concat(nav.SelectSingleNode("@SizeRatio"));
            item.Option = String.Concat(nav.SelectSingleNode("@Option"));
            item.StyleNo = String.Concat(nav.SelectSingleNode("@StyleNo"));

            try { item.Align = (AlignmentType1)Enum.Parse(typeof(AlignmentType1), String.Concat(nav.SelectSingleNode("@Align")), true); }
            catch { }

            return item;
        }
    }

    public sealed class HiddenCommentElement : IHwpmlElement, IHwpmlElement<TextElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TextElement)value; }
        }

        public string ElementName { get { return "HIDDENCOMMENT"; } }
        public TextElement Parent { get; set; }

        public ParaListElement ParaList;

        public static HiddenCommentElement Create(IXPathNavigable target, TextElement parent)
        {
            if (target == null)
                return null;

            HiddenCommentElement item = new HiddenCommentElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ParaList = ParaListElement.Create(nav.SelectSingleNode("PARALIST"), item);

            return item;
        }
    }

    public sealed class TailElement : IHwpmlElement, IHwpmlElement<HwpmlElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HwpmlElement)value; }
        }

        public string ElementName { get { return "TAIL"; } }
        public HwpmlElement Parent { get; set; }

        public BinDataStorageElement BinDataStorage;
        public ScriptCodeElement ScriptCode;
        public XmlTemplateElement XmlTemplate;

        public static TailElement Create(IXPathNavigable target, HwpmlElement parent)
        {
            if (target == null)
                return null;

            TailElement item = new TailElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.BinDataStorage = BinDataStorageElement.Create(nav.SelectSingleNode("BINDATASTORAGE"), item);
            item.ScriptCode = ScriptCodeElement.Create(nav.SelectSingleNode("SCRIPTCODE"), item);
            item.XmlTemplate = XmlTemplateElement.Create(nav.SelectSingleNode("XMLTEMPLATE"), item);

            return item;
        }
    }

    public sealed class BinDataStorageElement : List<BinDataElement>, IHwpmlElement, IHwpmlElement<TailElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TailElement)value; }
        }

        public string ElementName { get { return "BINDATASTORAGE"; } }
        public TailElement Parent { get; set; }

        public static BinDataStorageElement Create(IXPathNavigable target, TailElement parent)
        {
            if (target == null)
                return null;

            BinDataStorageElement item = new BinDataStorageElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("BINDATA"))
                item.Add(BinDataElement.Create(each, item));

            return item;
        }
    }

    public sealed class BinDataElement : IHwpmlElement, IHwpmlElement<BinDataStorageElement>, ITextElement
    {
        public BinDataElement()
            : base()
        {
            this.Encoding = "Base64";
            this.Compress = true;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (BinDataStorageElement)value; }
        }

        public string ElementName { get { return "BINDATA"; } }
        public BinDataStorageElement Parent { get; set; }

        public string Content { get; set; }

        public string Id;
        public long Size;
        public string Encoding;
        public bool Compress;

        public Stream GetStream(bool tryDecompress)
        {
            Stream memStream = new MemoryStream(Convert.FromBase64String(this.Content));
            return tryDecompress ?
                new GZipStream(memStream, CompressionMode.Decompress) :
                memStream;
        }

        public Stream GetStream() { return this.GetStream(this.Compress); }

        public static BinDataElement Create(IXPathNavigable target, BinDataStorageElement parent)
        {
            if (target == null)
                return null;

            BinDataElement item = new BinDataElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.Content = nav.ToString();

            item.Id = String.Concat(nav.SelectSingleNode("@Id"));
            
            try { item.Size = Int64.Parse(String.Concat(nav.SelectSingleNode("@Size"))); }
            catch { }

            item.Encoding = String.Concat(nav.SelectSingleNode("@Encoding"));
            
            try { item.Compress = Boolean.Parse(String.Concat(nav.SelectSingleNode("@Compress"))); }
            catch { }

            return item;
        }
    }

    public sealed class ScriptCodeElement : IHwpmlElement, IHwpmlElement<TailElement>
    {
        public ScriptCodeElement()
            : base()
        {
            this.Type = "JScript";
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TailElement)value; }
        }

        public string ElementName { get { return "SCRIPTCODE"; } }
        public TailElement Parent { get; set; }

        public string ScriptHeader;
        public string ScriptSource;
        public ScriptElement PreScript;
        public ScriptElement PostScript;

        public string Type;
        public string Version;

        public static ScriptCodeElement Create(IXPathNavigable target, TailElement parent)
        {
            if (target == null)
                return null;

            ScriptCodeElement item = new ScriptCodeElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.ScriptHeader = String.Concat(nav.SelectSingleNode("SCRIPTHEADER"));
            item.ScriptSource = String.Concat(nav.SelectSingleNode("SCRIPTSOURCE"));
            item.PreScript = ScriptElement.Create(nav.SelectSingleNode("PRESCRIPT"), item);
            item.PostScript = ScriptElement.Create(nav.SelectSingleNode("POSTSCRIPT"), item);

            item.Type = String.Concat(nav.SelectSingleNode("@Type"));
            item.Version = String.Concat(nav.SelectSingleNode("@Version"));

            return item;
        }
    }

    public sealed class ScriptElement : List<string>, IHwpmlElement, IHwpmlElement<ScriptCodeElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (ScriptCodeElement)value; }
        }

        public string ElementName { get { return "SCRIPT"; } }
        public ScriptCodeElement Parent { get; set; }

        public static ScriptElement Create(IXPathNavigable target, ScriptCodeElement parent)
        {
            if (target == null)
                return null;

            ScriptElement item = new ScriptElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            foreach (IXPathNavigable each in nav.Select("SCRIPTFUNCTION"))
                item.Add(String.Concat(each.CreateNavigator().ToString()));

            return item;
        }
    }

    public sealed class XmlTemplateElement : IHwpmlElement, IHwpmlElement<TailElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (TailElement)value; }
        }

        public string ElementName { get { return "XMLTEMPLATE"; } }
        public TailElement Parent { get; set; }

        public string Schema;
        public string Instance;

        public static XmlTemplateElement Create(IXPathNavigable target, TailElement parent)
        {
            if (target == null)
                return null;

            XmlTemplateElement item = new XmlTemplateElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.Schema = String.Concat(nav.SelectSingleNode("SCHEMA"));
            item.Instance = String.Concat(nav.SelectSingleNode("INSTANCE"));

            return item;
        }
    }

    public enum CompatibleDocumentTargetProgram : byte
    {
        None,
        Hwp70,
        Word
    }

    public sealed class CompatibleDocumentElement : IHwpmlElement, IHwpmlElement<HeadElement>
    {
        public CompatibleDocumentElement()
            : base()
        {
            this.TargetProgram = CompatibleDocumentTargetProgram.None;
        }

        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (HeadElement)value; }
        }

        public string ElementName { get { return "COMPATIBLEDOCUMENT"; } }
        public HeadElement Parent { get; set; }

        public LayoutCompatibilityElement LayoutCompatibility;

        public CompatibleDocumentTargetProgram TargetProgram;

        public static CompatibleDocumentElement Create(IXPathNavigable target, HeadElement parent)
        {
            if (target == null)
                return null;

            CompatibleDocumentElement item = new CompatibleDocumentElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;
            item.LayoutCompatibility = LayoutCompatibilityElement.Create(nav.SelectSingleNode("LAYOUTCOMPATIBILITY"), item);

            try { item.TargetProgram = (CompatibleDocumentTargetProgram)Enum.Parse(typeof(CompatibleDocumentTargetProgram), String.Concat(nav.SelectSingleNode("@TargetProgram")), true); }
            catch { }

            return item;
        }
    }

    public sealed class LayoutCompatibilityElement : IHwpmlElement, IHwpmlElement<CompatibleDocumentElement>
    {
        IHwpmlElement IHwpmlElement.Parent
        {
            get { return this.Parent; }
            set { this.Parent = (CompatibleDocumentElement)value; }
        }

        public string ElementName { get { return "LAYOUTCOMPATIBILITY"; } }
        public CompatibleDocumentElement Parent { get; set; }

        public bool ApplyFontWeightToBold;
        public bool UseInnerUnderline;
        public bool FixedUnderlineWidth;
        public bool DoNotApplyStrikeout;
        public bool UseLowercaseStrikeout;
        public bool ExtendLineheightToOffset;
        public bool TreatQuotationAsLatin;
        public bool DoNotAlignWhitespaceOnRight;
        public bool DoNotAdjustWordInJustify;
        public bool BaseCharUnitOnEAsian;
        public bool BaseCharUnitOfIndentOnFirstChar;
        public bool AdjustLineheightToFont;
        public bool AdjustBaselineInFixedLinespacing;
        public bool ExcludeOverlappingParaSpacing;
        public bool ApplyNextspacingOfLastPara;
        public bool ApplyAtLeastToPercent100Pct;
        public bool DoNotApplyAutoSpaceEAsianEng;
        public bool DoNotApplyAutoSpaceEAsianNum;
        public bool AdjustParaBorderfillToSpacing;
        public bool ConnectParaBorderfillOfEqualBorder;
        public bool AdjustParaBorderOffsetWithBorder;
        public bool ExtendLineheightToParaBorderOffset;
        public bool ApplyParaBorderToOutside;
        public bool BaseLinespacingOnLinegrid;
        public bool ApplyCharSpacingToCharGrid;
        public bool DoNotApplyGridInHeaderfooter;
        public bool ExtendHeaderfooterToBody;
        public bool AdjustEndnotePositionToFootnote;
        public bool DoNotApplyImageEffect;
        public bool DoNotApplyShapeComment;
        public bool DoNotAdjustEmptyAnchorLine;
        public bool OverlapBothAllowOverlap;
        public bool DoNotApplyVertOffsetOfForward;
        public bool ExtendVertLimitToPageMargins;
        public bool DoNotHoldAnchorOfTable;
        public bool DoNotFormattingAtBeneathAnchor;
        public bool DoNotApplyExtensionCharCompose;

        public static LayoutCompatibilityElement Create(IXPathNavigable target, CompatibleDocumentElement parent)
        {
            if (target == null)
                return null;

            LayoutCompatibilityElement item = new LayoutCompatibilityElement();

            XPathNavigator nav = target.CreateNavigator();
            item.Parent = parent;

            try { item.ApplyFontWeightToBold = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ApplyFontWeightToBold"))); }
            catch { }

            try { item.UseInnerUnderline = Boolean.Parse(String.Concat(nav.SelectSingleNode("@UseInnerUnderline"))); }
            catch { }

            try { item.FixedUnderlineWidth = Boolean.Parse(String.Concat(nav.SelectSingleNode("@FixedUnderlineWidth"))); }
            catch { }

            try { item.DoNotApplyStrikeout = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyStrikeout"))); }
            catch { }

            try { item.UseLowercaseStrikeout = Boolean.Parse(String.Concat(nav.SelectSingleNode("@UseLowercaseStrikeout"))); }
            catch { }

            try { item.ExtendLineheightToOffset = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ExtendLineheightToOffset"))); }
            catch { }

            try { item.TreatQuotationAsLatin = Boolean.Parse(String.Concat(nav.SelectSingleNode("@TreatQuotationAsLatin"))); }
            catch { }

            try { item.DoNotAlignWhitespaceOnRight = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotAlignWhitespaceOnRight"))); }
            catch { }

            try { item.DoNotAdjustWordInJustify = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotAdjustWordInJustify"))); }
            catch { }

            try { item.BaseCharUnitOnEAsian = Boolean.Parse(String.Concat(nav.SelectSingleNode("@BaseCharUnitOnEAsian"))); }
            catch { }

            try { item.BaseCharUnitOfIndentOnFirstChar = Boolean.Parse(String.Concat(nav.SelectSingleNode("@BaseCharUnitOfIndentOnFirstChar"))); }
            catch { }

            try { item.AdjustLineheightToFont = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AdjustLineheightToFont"))); }
            catch { }

            try { item.AdjustBaselineInFixedLinespacing = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AdjustBaselineInFixedLinespacing"))); }
            catch { }

            try { item.ExcludeOverlappingParaSpacing = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ExcludeOverlappingParaSpacing"))); }
            catch { }

            try { item.ApplyNextspacingOfLastPara = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ApplyNextspacingOfLastPara"))); }
            catch { }

            try { item.ApplyAtLeastToPercent100Pct = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ApplyAtLeastToPercent100Pct"))); }
            catch { }

            try { item.DoNotApplyAutoSpaceEAsianEng = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyAutoSpaceEAsianEng"))); }
            catch { }

            try { item.DoNotApplyAutoSpaceEAsianNum = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyAutoSpaceEAsianNum"))); }
            catch { }

            try { item.AdjustParaBorderfillToSpacing = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AdjustParaBorderfillToSpacing"))); }
            catch { }

            try { item.ConnectParaBorderfillOfEqualBorder = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ConnectParaBorderfillOfEqualBorder"))); }
            catch { }

            try { item.AdjustParaBorderOffsetWithBorder = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AdjustParaBorderOffsetWithBorder"))); }
            catch { }

            try { item.ExtendLineheightToParaBorderOffset = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ExtendLineheightToParaBorderOffset"))); }
            catch { }

            try { item.ApplyParaBorderToOutside = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ApplyParaBorderToOutside"))); }
            catch { }

            try { item.BaseLinespacingOnLinegrid = Boolean.Parse(String.Concat(nav.SelectSingleNode("@BaseLinespacingOnLinegrid"))); }
            catch { }

            try { item.ApplyCharSpacingToCharGrid = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ApplyCharSpacingToCharGrid"))); }
            catch { }

            try { item.DoNotApplyGridInHeaderfooter = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyGridInHeaderfooter"))); }
            catch { }

            try { item.ExtendHeaderfooterToBody = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ExtendHeaderfooterToBody"))); }
            catch { }

            try { item.AdjustEndnotePositionToFootnote = Boolean.Parse(String.Concat(nav.SelectSingleNode("@AdjustEndnotePositionToFootnote"))); }
            catch { }

            try { item.DoNotApplyImageEffect = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyImageEffect"))); }
            catch { }

            try { item.DoNotApplyShapeComment = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyShapeComment"))); }
            catch { }

            try { item.DoNotAdjustEmptyAnchorLine = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotAdjustEmptyAnchorLine"))); }
            catch { }

            try { item.OverlapBothAllowOverlap = Boolean.Parse(String.Concat(nav.SelectSingleNode("@OverlapBothAllowOverlap"))); }
            catch { }

            try { item.DoNotApplyVertOffsetOfForward = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyVertOffsetOfForward"))); }
            catch { }

            try { item.ExtendVertLimitToPageMargins = Boolean.Parse(String.Concat(nav.SelectSingleNode("@ExtendVertLimitToPageMargins"))); }
            catch { }

            try { item.DoNotHoldAnchorOfTable = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotHoldAnchorOfTable"))); }
            catch { }

            try { item.DoNotFormattingAtBeneathAnchor = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotFormattingAtBeneathAnchor"))); }
            catch { }

            try { item.DoNotApplyExtensionCharCompose = Boolean.Parse(String.Concat(nav.SelectSingleNode("@DoNotApplyExtensionCharCompose"))); }
            catch { }

            return item;
        }
    }
}

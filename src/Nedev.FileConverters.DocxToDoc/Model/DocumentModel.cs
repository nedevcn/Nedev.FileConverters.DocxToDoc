using System.Collections.Generic;

namespace Nedev.FileConverters.DocxToDoc.Model
{
    /// <summary>
    /// Represents the parsed content of a document, ready for serialization to binary MS-DOC format.
    /// </summary>
    public class DocumentModel
    {
        public List<object> Content { get; } = new List<object>(); // Can be ParagraphModel or TableModel
        public List<ParagraphModel> Paragraphs { get; } = new List<ParagraphModel>(); // Keeping for compatibility? No, let's refactor.
        public List<StyleModel> Styles { get; } = new List<StyleModel>();
        public List<FontModel> Fonts { get; } = new List<FontModel>();
        public List<SectionModel> Sections { get; } = new List<SectionModel>();
        public List<AbstractNumberingModel> AbstractNumbering { get; } = new List<AbstractNumberingModel>();
        public List<NumberingInstanceModel> NumberingInstances { get; } = new List<NumberingInstanceModel>();
        public List<BookmarkModel> Bookmarks { get; } = new List<BookmarkModel>();
        public List<CommentModel> Comments { get; } = new List<CommentModel>();
        public DocumentProperties Properties { get; } = new DocumentProperties();

        // As the Document is parsed we will accumulate the plain text here
        // The length of this text will determine the Piece Table and CCP Text
        public string TextBuffer { get; set; } = string.Empty;
    }

    public class DocumentProperties
    {
        public string? Title { get; set; }
        public string? Subject { get; set; }
        public string? Author { get; set; }
        public string? Manager { get; set; }
        public string? Company { get; set; }
        public string? Category { get; set; }
        public string? Keywords { get; set; }
        public string? Comments { get; set; }
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
        public DateTime? LastPrinted { get; set; }
        public int? Revision { get; set; }
        public int? TotalEditingTime { get; set; }
        public int? Pages { get; set; }
        public int? Words { get; set; }
        public int? Characters { get; set; }
    }

    public class SectionModel
    {
        public int PageWidth { get; set; } = 11906; // Default Letter/A4
        public int PageHeight { get; set; } = 16838;
        public int MarginLeft { get; set; } = 1440;
        public int MarginRight { get; set; } = 1440;
        public int MarginTop { get; set; } = 1440;
        public int MarginBottom { get; set; } = 1440;

        // Header/Footer references
        public string? HeaderReference { get; set; }
        public string? FooterReference { get; set; }
        public string? FirstPageHeaderReference { get; set; }
        public string? FirstPageFooterReference { get; set; }
        public string? EvenPagesHeaderReference { get; set; }
        public string? EvenPagesFooterReference { get; set; }

        // Page numbering
        public int StartPageNumber { get; set; } = 1;
        public bool RestartPageNumbering { get; set; } = false;
    }

    public class AbstractNumberingModel
    {
        public int Id { get; set; }
        public List<NumberingLevelModel> Levels { get; } = new List<NumberingLevelModel>();
    }

    public class NumberingLevelModel
    {
        public int Level { get; set; }
        public string NumberFormat { get; set; } = "decimal";
        public string LevelText { get; set; } = string.Empty;
        public int Start { get; set; } = 1;
    }

    public class NumberingInstanceModel
    {
        public int Id { get; set; }
        public int AbstractNumberId { get; set; }
    }

    public class StyleModel
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public bool IsParagraphStyle { get; set; }
        public int StyleId { get; set; } // ISTD index
        public int? BasedOn { get; set; } // Parent style ID
        public int? NextStyle { get; set; } // Next paragraph style
        public ParagraphModel.ParagraphProperties? ParagraphProps { get; set; }
        public RunModel.CharacterProperties? CharacterProps { get; set; }
    }

    public class FontModel
    {
        public string Name { get; set; } = string.Empty;
        public FontFamily Family { get; set; } = FontFamily.Auto;
        public FontPitch Pitch { get; set; } = FontPitch.Default;
        public short Weight { get; set; } = 400; // Normal weight
        public byte Charset { get; set; } = 0; // ANSI_CHARSET
    }

    public enum FontFamily : byte
    {
        Auto = 0,
        Roman = 1,
        Swiss = 2,
        Modern = 3,
        Script = 4,
        Decorative = 5
    }

    public enum FontPitch : byte
    {
        Default = 0,
        Fixed = 1,
        Variable = 2
    }

    public class TableModel
    {
        public List<TableRowModel> Rows { get; } = new List<TableRowModel>();
        public List<int> GridColumnWidths { get; } = new List<int>();
        public int PreferredWidthValue { get; set; }
        public TableWidthUnit PreferredWidthUnit { get; set; } = TableWidthUnit.Auto;
        public int CellSpacingTwips { get; set; }
        public int DefaultInsideHorizontalBorderTwips { get; set; }
        public int DefaultInsideVerticalBorderTwips { get; set; }
        public int DefaultBorderLeftTwips { get; set; }
        public int DefaultBorderRightTwips { get; set; }
        public int DefaultBorderTopTwips { get; set; }
        public int DefaultBorderBottomTwips { get; set; }
        public int DefaultCellPaddingLeftTwips { get; set; }
        public int DefaultCellPaddingRightTwips { get; set; }
        public int DefaultCellPaddingTopTwips { get; set; }
        public int DefaultCellPaddingBottomTwips { get; set; }
    }

    public enum TableWidthUnit
    {
        Auto = 0,
        Dxa = 1,
        Pct = 2
    }

    public class TableRowModel
    {
        public List<TableCellModel> Cells { get; } = new List<TableCellModel>();
        public int HeightTwips { get; set; }
        public TableRowHeightRule HeightRule { get; set; } = TableRowHeightRule.Auto;
    }

    public enum TableRowHeightRule
    {
        Auto = 0,
        AtLeast = 1,
        Exact = 2
    }

    public class TableCellModel
    {
        public List<ParagraphModel> Paragraphs { get; } = new List<ParagraphModel>();
        public int Width { get; set; }
        public TableWidthUnit WidthUnit { get; set; } = TableWidthUnit.Dxa;
        public int GridSpan { get; set; } = 1;
        public TableCellVerticalAlignment VerticalAlignment { get; set; } = TableCellVerticalAlignment.Top;
        public bool HasLeftPaddingOverride { get; set; }
        public bool HasRightPaddingOverride { get; set; }
        public bool HasTopPaddingOverride { get; set; }
        public bool HasBottomPaddingOverride { get; set; }
        public bool HasLeftBorderOverride { get; set; }
        public bool HasRightBorderOverride { get; set; }
        public bool HasTopBorderOverride { get; set; }
        public bool HasBottomBorderOverride { get; set; }
        public int BorderLeftTwips { get; set; }
        public int BorderRightTwips { get; set; }
        public int BorderTopTwips { get; set; }
        public int BorderBottomTwips { get; set; }
        public int PaddingLeftTwips { get; set; }
        public int PaddingRightTwips { get; set; }
        public int PaddingTopTwips { get; set; }
        public int PaddingBottomTwips { get; set; }
    }

    public enum TableCellVerticalAlignment
    {
        Top = 0,
        Center = 1,
        Bottom = 2
    }

    public class ParagraphModel
    {
        public List<RunModel> Runs { get; } = new List<RunModel>();
        public ParagraphProperties Properties { get; } = new ParagraphProperties();

        public class ParagraphProperties
        {
            public Justification Alignment { get; set; } = Justification.Left;
            public int? NumberingId { get; set; }
            public int? NumberingLevel { get; set; }
            public int LeftIndentTwips { get; set; }
            public int RightIndentTwips { get; set; }
            public int FirstLineIndentTwips { get; set; }
            public int SpacingBeforeTwips { get; set; }
            public int SpacingAfterTwips { get; set; }
            public int? LineSpacing { get; set; }
            public string? LineSpacingRule { get; set; }
        }

        public enum Justification
        {
            Left, Center, Right, Both
        }
    }

    public class RunModel
    {
        public string Text { get; set; } = string.Empty;
        public CharacterProperties Properties { get; } = new CharacterProperties();
        public ImageModel? Image { get; set; }
        public HyperlinkModel? Hyperlink { get; set; }
        public FieldModel? Field { get; set; }
        public bool IsFieldBegin { get; set; }
        public bool IsFieldSeparate { get; set; }
        public bool IsFieldEnd { get; set; }

        public class CharacterProperties
        {
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
            public bool IsStrike { get; set; }
            public int? FontSize { get; set; } // In half-points (e.g., 24 = 12pt)
            public string? FontName { get; set; }
            public UnderlineType Underline { get; set; } = UnderlineType.None;
            public string? Color { get; set; } // Hex color like "FF0000"
        }
    }

    public enum UnderlineType
    {
        None = 0,
        Single = 1,
        Double = 2,
        Thick = 3,
        Dotted = 4,
        Dashed = 5,
        Wave = 6
    }

    public class HyperlinkModel
    {
        public string? RelationshipId { get; set; }
        public string? Anchor { get; set; } // Internal bookmark
        public string? Tooltip { get; set; }
        public string DisplayText { get; set; } = string.Empty;
        public string TargetUrl { get; set; } = string.Empty;
    }

    public class ImageModel
    {
        public string Id { get; set; } = string.Empty;
        public string RelationshipId { get; set; } = string.Empty;
        public byte[]? Data { get; set; }
        public string ContentType { get; set; } = string.Empty;
        public int Width { get; set; }
        public int Height { get; set; }
        public string? FileName { get; set; }
        public ImageLayoutType LayoutType { get; set; } = ImageLayoutType.Inline;
        public ImageWrapType WrapType { get; set; } = ImageWrapType.Inline;
        public string? HorizontalRelativeTo { get; set; }
        public string? VerticalRelativeTo { get; set; }
        public string? HorizontalAlignment { get; set; }
        public string? VerticalAlignment { get; set; }
        public int PositionXTwips { get; set; }
        public int PositionYTwips { get; set; }
        public int DistanceLeftTwips { get; set; }
        public int DistanceRightTwips { get; set; }
        public int DistanceTopTwips { get; set; }
        public int DistanceBottomTwips { get; set; }
        public bool BehindText { get; set; }
        public bool AllowOverlap { get; set; } = true;
    }

    public enum ImageLayoutType
    {
        Inline = 0,
        Floating = 1
    }

    public enum ImageWrapType
    {
        Inline = 0,
        None = 1,
        Square = 2,
        Tight = 3,
        Through = 4,
        TopAndBottom = 5
    }

    public class FieldModel
    {
        public FieldType Type { get; set; }
        public string Instruction { get; set; } = string.Empty;
        public string Result { get; set; } = string.Empty;
        public bool IsLocked { get; set; }
        public bool IsDirty { get; set; }
    }

    public enum FieldType
    {
        Unknown = 0,
        Page = 1,           // PAGE - Current page number
        NumPages = 2,       // NUMPAGES - Total pages
        Date = 3,           // DATE - Current date
        Time = 4,           // TIME - Current time
        Author = 5,         // AUTHOR - Document author
        Title = 6,          // TITLE - Document title
        Subject = 7,        // SUBJECT - Document subject
        FileName = 8,       // FILENAME - Document filename
        Hyperlink = 9,      // HYPERLINK - Hyperlink field
        Bookmark = 10,      // BOOKMARK - Bookmark reference
        Index = 11,         // INDEX - Table of contents/index
        Seq = 12,           // SEQ - Sequence field
        Ref = 13,           // REF - Cross-reference
        MergeField = 14     // MERGEFIELD - Mail merge field
    }

    public class BookmarkModel
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public int StartCp { get; set; }
        public int EndCp { get; set; }
        public bool IsCollapsed { get; set; }
    }

    /// <summary>
    /// Represents a comment/annotation in the document.
    /// </summary>
    public class CommentModel
    {
        /// <summary>
        /// Gets or sets the unique identifier of the comment.
        /// </summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the author of the comment.
        /// </summary>
        public string Author { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the initials of the author.
        /// </summary>
        public string Initials { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the date and time when the comment was created.
        /// </summary>
        public DateTime? Date { get; set; }

        /// <summary>
        /// Gets or sets the text content of the comment.
        /// </summary>
        public string Text { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the starting character position of the comment reference.
        /// </summary>
        public int StartCp { get; set; }

        /// <summary>
        /// Gets or sets the ending character position of the comment reference.
        /// </summary>
        public int EndCp { get; set; }

        /// <summary>
        /// Gets or sets whether this comment is a reply to another comment.
        /// </summary>
        public bool IsReply { get; set; }

        /// <summary>
        /// Gets or sets the ID of the parent comment if this is a reply.
        /// </summary>
        public string? ParentId { get; set; }

        /// <summary>
        /// Gets or sets whether the comment is marked as done/resolved.
        /// </summary>
        public bool IsDone { get; set; }
    }
}

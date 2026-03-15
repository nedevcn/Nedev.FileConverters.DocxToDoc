using System;

namespace Nedev.FileConverters.DocxToDoc
{
    /// <summary>
    /// Represents errors that occur during document conversion.
    /// </summary>
    public class ConversionException : Exception
    {
        /// <summary>
        /// Gets the path to the source file being converted, if applicable.
        /// </summary>
        public string? SourcePath { get; }

        /// <summary>
        /// Gets the path to the destination file, if applicable.
        /// </summary>
        public string? DestinationPath { get; }

        /// <summary>
        /// Gets the stage of conversion where the error occurred.
        /// </summary>
        public ConversionStage Stage { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionException"/> class.
        /// </summary>
        public ConversionException(string message)
            : base(message)
        {
            Stage = ConversionStage.Unknown;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionException"/> class.
        /// </summary>
        public ConversionException(string message, Exception innerException)
            : base(message, innerException)
        {
            Stage = ConversionStage.Unknown;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionException"/> class.
        /// </summary>
        public ConversionException(string message, ConversionStage stage)
            : base(message)
        {
            Stage = stage;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionException"/> class.
        /// </summary>
        public ConversionException(string message, ConversionStage stage, Exception innerException)
            : base(message, innerException)
        {
            Stage = stage;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionException"/> class.
        /// </summary>
        public ConversionException(string message, string? sourcePath, string? destinationPath, ConversionStage stage)
            : base(message)
        {
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
            Stage = stage;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionException"/> class.
        /// </summary>
        public ConversionException(string message, string? sourcePath, string? destinationPath, ConversionStage stage, Exception innerException)
            : base(message, innerException)
        {
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
            Stage = stage;
        }
    }

    /// <summary>
    /// Represents the stage of conversion where an error occurred.
    /// </summary>
    public enum ConversionStage
    {
        /// <summary>
        /// Unknown stage.
        /// </summary>
        Unknown,

        /// <summary>
        /// Input validation stage.
        /// </summary>
        Validation,

        /// <summary>
        /// Reading the source DOCX file.
        /// </summary>
        Reading,

        /// <summary>
        /// Parsing document content.
        /// </summary>
        Parsing,

        /// <summary>
        /// Writing the destination DOC file.
        /// </summary>
        Writing,

        /// <summary>
        /// Finalizing the output.
        /// </summary>
        Finalizing
    }
}

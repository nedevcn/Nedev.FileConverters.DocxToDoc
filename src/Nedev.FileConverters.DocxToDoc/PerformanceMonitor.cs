using System;
using System.Diagnostics;

namespace Nedev.FileConverters.DocxToDoc
{
    /// <summary>
    /// Provides performance monitoring and metrics collection for document conversion operations.
    /// </summary>
    public class PerformanceMonitor : IDisposable
    {
        private readonly Stopwatch _stopwatch;
        private readonly ILogger? _logger;
        private long _bytesRead;
        private long _bytesWritten;
        private int _paragraphsProcessed;
        private int _imagesProcessed;
        private int _tablesProcessed;

        /// <summary>
        /// Gets the elapsed time since monitoring started.
        /// </summary>
        public TimeSpan Elapsed => _stopwatch.Elapsed;

        /// <summary>
        /// Gets the total bytes read during the operation.
        /// </summary>
        public long BytesRead => _bytesRead;

        /// <summary>
        /// Gets the total bytes written during the operation.
        /// </summary>
        public long BytesWritten => _bytesWritten;

        /// <summary>
        /// Gets the number of paragraphs processed.
        /// </summary>
        public int ParagraphsProcessed => _paragraphsProcessed;

        /// <summary>
        /// Gets the number of images processed.
        /// </summary>
        public int ImagesProcessed => _imagesProcessed;

        /// <summary>
        /// Gets the number of tables processed.
        /// </summary>
        public int TablesProcessed => _tablesProcessed;

        /// <summary>
        /// Gets the processing rate in bytes per second.
        /// </summary>
        public double BytesPerSecond
        {
            get
            {
                var elapsedSeconds = _stopwatch.Elapsed.TotalSeconds;
                return elapsedSeconds > 0 ? _bytesRead / elapsedSeconds : 0;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PerformanceMonitor"/> class.
        /// </summary>
        public PerformanceMonitor() : this(null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PerformanceMonitor"/> class with a logger.
        /// </summary>
        /// <param name="logger">The logger to use for performance logging.</param>
        public PerformanceMonitor(ILogger? logger)
        {
            _logger = logger;
            _stopwatch = new Stopwatch();
        }

        /// <summary>
        /// Starts performance monitoring.
        /// </summary>
        public void Start()
        {
            _stopwatch.Start();
            _logger?.LogDebug("Performance monitoring started");
        }

        /// <summary>
        /// Stops performance monitoring.
        /// </summary>
        public void Stop()
        {
            _stopwatch.Stop();
            _logger?.LogDebug($"Performance monitoring stopped. Elapsed: {Elapsed.TotalMilliseconds:F2}ms");
        }

        /// <summary>
        /// Records the number of bytes read.
        /// </summary>
        public void RecordBytesRead(long bytes)
        {
            _bytesRead += bytes;
        }

        /// <summary>
        /// Records the number of bytes written.
        /// </summary>
        public void RecordBytesWritten(long bytes)
        {
            _bytesWritten += bytes;
        }

        /// <summary>
        /// Records a processed paragraph.
        /// </summary>
        public void RecordParagraph()
        {
            _paragraphsProcessed++;
        }

        /// <summary>
        /// Records multiple processed paragraphs.
        /// </summary>
        public void RecordParagraphs(int count)
        {
            _paragraphsProcessed += count;
        }

        /// <summary>
        /// Records a processed image.
        /// </summary>
        public void RecordImage()
        {
            _imagesProcessed++;
        }

        /// <summary>
        /// Records multiple processed images.
        /// </summary>
        public void RecordImages(int count)
        {
            _imagesProcessed += count;
        }

        /// <summary>
        /// Records a processed table.
        /// </summary>
        public void RecordTable()
        {
            _tablesProcessed++;
        }

        /// <summary>
        /// Records multiple processed tables.
        /// </summary>
        public void RecordTables(int count)
        {
            _tablesProcessed += count;
        }

        /// <summary>
        /// Logs a performance summary.
        /// </summary>
        public void LogSummary()
        {
            if (_logger == null) return;

            _logger.LogInfo("=== Performance Summary ===");
            _logger.LogInfo($"Elapsed Time: {Elapsed.TotalMilliseconds:F2}ms");
            _logger.LogInfo($"Bytes Read: {_bytesRead:N0}");
            _logger.LogInfo($"Bytes Written: {_bytesWritten:N0}");
            _logger.LogInfo($"Processing Rate: {BytesPerSecond:N0} bytes/second");
            _logger.LogInfo($"Paragraphs: {_paragraphsProcessed}");
            _logger.LogInfo($"Tables: {_tablesProcessed}");
            _logger.LogInfo($"Images: {_imagesProcessed}");
            _logger.LogInfo("===========================");
        }

        /// <summary>
        /// Gets a performance summary string.
        /// </summary>
        public string GetSummary()
        {
            return $"Time: {Elapsed.TotalMilliseconds:F2}ms, " +
                   $"Read: {_bytesRead:N0} bytes, " +
                   $"Written: {_bytesWritten:N0} bytes, " +
                   $"Rate: {BytesPerSecond:N0} bytes/s";
        }

        /// <summary>
        /// Releases all resources used by the performance monitor.
        /// </summary>
        public void Dispose()
        {
            if (_stopwatch.IsRunning)
            {
                Stop();
            }
        }
    }

    /// <summary>
    /// Provides memory usage monitoring utilities.
    /// </summary>
    public static class MemoryMonitor
    {
        /// <summary>
        /// Gets the current working set size in bytes.
        /// </summary>
        public static long GetWorkingSet()
        {
            return Process.GetCurrentProcess().WorkingSet64;
        }

        /// <summary>
        /// Gets the current managed memory usage in bytes.
        /// </summary>
        public static long GetManagedMemory()
        {
            return GC.GetTotalMemory(false);
        }

        /// <summary>
        /// Forces garbage collection and returns the managed memory usage.
        /// </summary>
        public static long CollectAndGetManagedMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            return GC.GetTotalMemory(true);
        }

        /// <summary>
        /// Logs current memory usage.
        /// </summary>
        public static void LogMemoryUsage(ILogger logger, string context = "")
        {
            var workingSet = GetWorkingSet();
            var managedMemory = GetManagedMemory();
            var prefix = string.IsNullOrEmpty(context) ? "" : $"[{context}] ";

            logger.LogDebug($"{prefix}Working Set: {workingSet:N0} bytes ({workingSet / 1024 / 1024} MB)");
            logger.LogDebug($"{prefix}Managed Memory: {managedMemory:N0} bytes ({managedMemory / 1024 / 1024} MB)");
        }
    }
}

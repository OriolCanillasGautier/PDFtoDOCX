using Microsoft.Extensions.DependencyInjection;
using PDFtoDOCX.Extraction;

namespace PDFtoDOCX.Extensions
{
    /// <summary>
    /// Dependency injection registration helpers for the PDFtoDOCX library.
    /// Enables integration with ASP.NET Core, Worker Services, and other
    /// Microsoft.Extensions.DependencyInjection hosts.
    /// </summary>
    public static class ServiceCollectionExtensions
    {
        /// <summary>
        /// Registers the PDF-to-DOCX converter and its dependencies with the
        /// default <see cref="ConversionOptions"/>.
        /// </summary>
        /// <example>
        /// <code>
        /// services.AddPdfToDocxConverter();
        /// </code>
        /// </example>
        public static IServiceCollection AddPdfToDocxConverter(
            this IServiceCollection services)
        {
            return services.AddPdfToDocxConverter(new ConversionOptions());
        }

        /// <summary>
        /// Registers the PDF-to-DOCX converter and its dependencies with
        /// caller-supplied options.
        /// </summary>
        /// <param name="services">The service collection to populate.</param>
        /// <param name="options">Conversion options shared by all pipeline components.</param>
        /// <example>
        /// <code>
        /// services.AddPdfToDocxConverter(new ConversionOptions
        /// {
        ///     EnableDiagnostics = true,
        ///     ParagraphSpacingAfter = 8.0
        /// });
        /// </code>
        /// </example>
        public static IServiceCollection AddPdfToDocxConverter(
            this IServiceCollection services,
            ConversionOptions options)
        {
            // Register options as singleton so all components share the same settings
            services.AddSingleton(options ?? new ConversionOptions());

            // Register the OCR stub (replace with a real implementation if required)
            services.AddSingleton<ITextExtractor, OcrTextExtractor>();

            // Register the main converter
            services.AddSingleton<IPdfToDocxConverter>(provider =>
            {
                var opts = provider.GetRequiredService<ConversionOptions>();
                return new PdfToDocxConverter(opts);
            });

            return services;
        }

        /// <summary>
        /// Registers conversion options configured via a delegate.
        /// </summary>
        /// <param name="services">The service collection to populate.</param>
        /// <param name="configure">Delegate that mutates a default <see cref="ConversionOptions"/> instance.</param>
        public static IServiceCollection AddPdfToDocxConverter(
            this IServiceCollection services,
            System.Action<ConversionOptions> configure)
        {
            var options = new ConversionOptions();
            configure?.Invoke(options);
            return services.AddPdfToDocxConverter(options);
        }
    }
}

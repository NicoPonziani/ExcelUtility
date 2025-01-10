package interfaces;

/**
 * Marker interface for classes that are used in the Excel generation process.
 * <p>
 * Any class implementing this interface must also implement {@link ExportBaseInterface},
 * ensuring that it provides the base functionality required for export operations.
 * </p>
 * <p>
 * Classes implementing this interface are expected to represent data structures
 * that can be used as rows or entities in an Excel export sheet.
 * </p>
 *
 * <h3>Usage:</h3>
 * <pre>{@code
 * public class MyExportClass implements interfaces.ExportSimple {
 *     // Implementation details
 * }
 * }</pre>
 */
public interface ExportSimple extends ExportBaseInterface{

}

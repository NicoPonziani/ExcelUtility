package classes;

import interfaces.ExportBaseInterface;
import interfaces.ExportSimple;

import java.util.Arrays;
import java.util.List;

/**
 * Represents an exportable report containing generalities and complex data sections.
 * <p>
 * This class serves as a container for general information and detailed data,
 * providing a structure for exporting reports in various formats.
 * </p>
 *
 * @param <E> The type of the generalities object, representing high-level data 
 *            included in the report.
 *
 * <h3>Features:</h3>
 * <ul>
 *   <li>Encapsulates general information of type {@code E}.</li>
 *   <li>Contains a list of complex data sections, each represented by 
 *       {@link ComplexExport} objects.</li>
 * </ul>
 *
 * <h3>Usage:</h3>
 * <pre>{@code
 * classes.ReportExport<MyGeneralInfo> report = new classes.ReportExport<>();
 * report.setGeneralities(new MyGeneralInfo(...));
 * report.setDati(new classes.ComplexExport<>(...));
 * }</pre>
 *
 * @see ComplexExport
 * @see ExportBaseInterface
 */
public class ReportExport<E> implements ExportBaseInterface {

    /**
     * General information about the report, represented as an object of type {@code E}.
     */
    private E generalities;

    /**
     * List of complex data sections to be included in the report.
     * Each section is represented by a {@link ComplexExport} instance.
     */
    private List<ComplexExport<? extends ExportSimple>> data;

    /**
     * Default constructor.
     * Initializes an empty report export.
     */
    public ReportExport() {
        super();
    }

    /**
     * Retrieves the generalities of the report.
     *
     * @return the generalities object of type {@code E}.
     */
    public E getGeneralities() {
        return generalities;
    }

    /**
     * Sets the generalities of the report.
     *
     * @param generalities the generalities object of type {@code E}.
     */
    public void setGeneralities(E generalities) {
        this.generalities = generalities;
    }

    /**
     * Retrieves the list of complex data sections.
     *
     * @return a list of {@link ComplexExport} instances.
     */
    public List<ComplexExport<?>> getData() {
        return data;
    }

    /**
     * Sets the data sections of the report using a variable number of 
     * {@link ComplexExport} instances.
     *
     * @param data an array of {@link ComplexExport} instances representing 
     *             the data sections.
     */
    @SafeVarargs
    public final void setDati(ComplexExport<? extends ExportSimple>... data) {
        this.data = Arrays.asList(data);
    }

    /**
     * Sets the data sections of the report using a list of {@link ComplexExport} instances.
     *
     * @param data a list of {@link ComplexExport} instances representing 
     *             the data sections.
     */
    public void setDati(List<ComplexExport<?>> data) {
        this.data = data;
    }
}


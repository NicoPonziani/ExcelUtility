package classes;

import interfaces.ExportSimple;

import java.util.List;

/**
 * A class representing a complex export operation for data of type {@code T}, 
 * which extends {@link ExportSimple}. This class is designed to facilitate 
 * exporting data in tabular format using a specified sheet name and a list of data objects.
 *
 * @param <T> the type of data objects to be exported, extending {@link ExportSimple}.
 */
public class ComplexExport<T extends ExportSimple> extends ExportBaseTable {

    /**
     * A list containing the data to be exported.
     */
    private List<T> dati;

    /**
     * Default constructor for {@code classes.ComplexExport}.
     * Initializes an empty export operation.
     */
    public ComplexExport() {
        super();
    }

    /**
     * Constructor for {@code classes.ComplexExport} that initializes the export with a list of data
     * and a specified sheet name.
     *
     * @param dati  the list of data objects to be exported.
     * @param sheet the name of the sheet where the data will be exported.
     */
    public ComplexExport(List<T> dati, String sheet) {
        super(sheet);
        this.dati = dati;
    }

    /**
     * Retrieves the list of data objects to be exported.
     *
     * @return the list of data objects of type {@code T}.
     */
    public List<T> getDati() {
        return dati;
    }

    /**
     * Sets the list of data objects to be exported.
     *
     * @param dati the list of data objects of type {@code T}.
     */
    public void setDati(List<T> dati) {
        this.dati = dati;
    }
}


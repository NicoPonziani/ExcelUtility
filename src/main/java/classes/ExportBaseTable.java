package classes;

import interfaces.ExportSimple;
import services.ExcelUtility.*;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * An abstract base class for defining exportable table structures. 
 * Implements {@link ExportSimple} and provides common properties
 * and functionality for table exports.
 */
public abstract class ExportBaseTable implements ExportSimple{

    /**
     * The name of the sheet where data will be exported. 
     * Defaults to "export".
     */
	private String sheet = "export";
	/**
     * A list of {@link SpecialField} objects representing fields with special
     * behaviors or properties in the export.
     */
	private List<SpecialField> specialFields = new ArrayList<>();
	 /**
     * A list of {@link ReferenceLabel} objects used for adding additional 
     * labels or references to the export sheet.
     */
	private List<ReferenceLabel> referenceLabels = new ArrayList<>();
	/**
     * A {@link Pivot} object for defining pivot table configurations in the export.
     */
	private Pivot pivot;
	/**
     * A flag indicating whether a header row should be included in the export. 
     * Defaults to {@code true}.
     */
	private boolean header = true;
	
	protected ExportBaseTable() {super();}
	
	/**
     * Protected constructor for initializing the export with a custom sheet name.
     *
     * @param sheet the name of the sheet to use for the export
     */
    protected ExportBaseTable(String sheet) {
        super();
        this.sheet = sheet;
    }

    /**
     * Gets the name of the sheet for the export.
     *
     * @return the sheet name
     */
    public String getSheet() {
        return sheet;
    }

    /**
     * Sets the name of the sheet for the export.
     *
     * @param sheet the new sheet name
     */
    public void setSheet(String sheet) {
        this.sheet = sheet;
    }

    /**
     * Gets the list of {@link SpecialField} objects.
     *
     * @return the list of special fields
     */
    public List<SpecialField> getSpecialFields() {
        return specialFields;
    }

    /**
     * Sets the list of {@link SpecialField} objects.
     *
     * @param specialFields the list of special fields to set
     */
    public void setSpecialFields(List<SpecialField> specialFields) {
        this.specialFields = specialFields;
    }

    /**
     * Provides a {@link BuilderSpecialField} instance to add special fields.
     *
     * @return a builder for special fields
     */
    public BuilderSpecialField setSpecialFields() {
        return new BuilderSpecialField();
    }
	
    /**
     * A builder class for creating and configuring {@link SpecialField} objects.
     * This builder is used to define special fields for an exportable table, 
     * allowing customization of labels, columns, operations, order, formulas, 
     * and cell styles.
     */
    public class BuilderSpecialField {

        /**
         * The {@link SpecialField} instance being built and configured.
         */
        private SpecialField special = new SpecialField();

        /**
         * An optional {@link BuilderPivot} instance used for linking the special 
         * field to a pivot table configuration.
         */
        private BuilderPivot pivotBuilder = null;

        /**
         * Private constructor for internal use, initializes a new builder.
         */
        private BuilderSpecialField() {
            super();
        }

        /**
         * Private constructor for internal use, initializes a builder with an 
         * associated {@link BuilderPivot}.
         *
         * @param pivotBuilder the builder for the associated pivot configuration
         */
        private BuilderSpecialField(BuilderPivot pivotBuilder) {
            super();
            this.pivotBuilder = pivotBuilder;
        }

        /**
         * Sets the label for the {@link SpecialField}.
         *
         * @param label the label to set
         * @return this builder instance
         */
        public BuilderSpecialField setLabel(String label) {
            special.setLabel(label);
            return this;
        }

        /**
         * Sets the columns associated with the {@link SpecialField}.
         *
         * @param columns an array of column names to set
         * @return this builder instance
         */
        public BuilderSpecialField setColumns(String... columns) {
            special.setColumns(columns);
            return this;
        }

        /**
         * Sets the operation for the {@link SpecialField}.
         *
         * @param operation the {@link OperationEnum} to set
         * @return this builder instance
         */
        public BuilderSpecialField setOperation(OperationEnum operation) {
            special.setOperation(operation);
            return this;
        }

        /**
         * Sets the order for the {@link SpecialField}.
         *
         * @param order the order to set
         * @return this builder instance
         */
        public BuilderSpecialField setOrder(int order) {
            special.setOrder(order);
            return this;
        }

        /**
         * Sets the formula for the {@link SpecialField}.
         * The formula should be in ODF format (e.g., LibreOffice formulas) but 
         * can also accept direct values.
         *
         * @param formula the formula in ODF format
         * @return this builder instance
         * @see it.rve.postemergenze.services.ReportTrasversaliService#reportElencoFatturaDitta()
         */
        public BuilderSpecialField setFormula(String formula) {
            special.setFormula(formula);
            return this;
        }

        /**
         * Sets the cell style for the {@link SpecialField}.
         *
         * @param style the {@link CellStyleEnum} to set
         * @return this builder instance
         */
        public BuilderSpecialField setCellStyle(CellStyleEnum style) {
            special.setStyle(style);
            return this;
        }

        /**
         * Builds the {@link SpecialField} and links it to the {@link BuilderPivot}.
         * This is used to configure the special field for use in a pivot table.
         *
         * @return the {@link BuilderPivot} instance
         * @throws RuntimeException if the pivotBuilder is null
         */
        public BuilderPivot buildSpecialFieldPivot() {
            if (pivotBuilder == null) {
                throw new RuntimeException("Pivot is null");
            }

            pivotBuilder.pivotBuild.setSpecialField(special);
            return pivotBuilder;
        }

        /**
         * Builds the {@link SpecialField} and adds it to the global list of 
         * special fields. If the list does not exist, it is initialized.
         */
        public void build() {
            if (specialFields == null) {
                specialFields = new ArrayList<>();
            }
            specialFields.add(special);
        }
    }

	
    /**
     * Retrieves the list of {@link ReferenceLabel} objects associated with the export.
     *
     * @return the list of reference labels
     */
    public List<ReferenceLabel> getReferenceLabels() {
        return referenceLabels;
    }

    /**
     * Sets the list of {@link ReferenceLabel} objects for the export.
     *
     * @param referenceLabels the list of reference labels to set
     */
    public void setReferenceLabels(List<ReferenceLabel> referenceLabels) {
        this.referenceLabels = referenceLabels;
    }

    /**
     * Adds one or more {@link ReferenceLabel} objects to the list of reference labels.
     * If the list is not initialized, it will be created.
     *
     * @param labels the reference labels to add
     */
    public void setReferenceLabels(ReferenceLabel... labels) {
        if (referenceLabels == null) {
            referenceLabels = new ArrayList<>();
        }
        referenceLabels.addAll(Arrays.asList(labels));
    }

    /**
     * Retrieves the {@link Pivot} configuration associated with the export.
     *
     * @return the {@link Pivot} object
     */
    public Pivot getPivot() {
        return pivot;
    }

    /**
     * Sets the {@link Pivot} configuration for the export.
     *
     * @param pivot the {@link Pivot} object to set
     */
    public void setPivot(Pivot pivot) {
        this.pivot = pivot;
    }

    /**
     * Creates a new {@link BuilderPivot} instance to configure a pivot table.
     *
     * @return a new {@link BuilderPivot} instance
     */
    public BuilderPivot setPivot() {
        return new BuilderPivot();
    }

    /**
     * Checks if the header row is included in the export.
     *
     * @return {@code true} if the header is included, {@code false} otherwise
     */
    public boolean isHeader() {
        return header;
    }

    /**
     * Sets whether the header row should be included in the export.
     *
     * @param header {@code true} to include the header, {@code false} otherwise
     */
    public void setHeader(boolean header) {
        this.header = header;
    }


    /**
     * Builder class for configuring a {@link Pivot} object.
     * Provides a fluent API for setting up a pivot table configuration.
     */
    public class BuilderPivot {

        /**
         * The {@link Pivot} instance being built.
         */
        Pivot pivotBuild = new Pivot();

        /**
         * Private constructor to enforce controlled instantiation via parent class.
         */
        private BuilderPivot() {
            super();
        }

        /**
         * Sets the label for the pivot table.
         *
         * @param label the label to set for the pivot table
         * @return the current instance of {@link BuilderPivot} for chaining
         */
        public BuilderPivot setLabel(String label) {
            this.pivotBuild.setLabel(label);
            return this;
        }

        /**
         * Specifies the columns that define the conditions for the pivot table.
         *
         * @param columnCondition one or more column names to use as conditions
         * @return the current instance of {@link BuilderPivot} for chaining
         */
        public BuilderPivot setColumnCondition(String... columnCondition) {
            this.pivotBuild.setColumnCondition(columnCondition);
            return this;
        }

        /**
         * Specifies the columns to be calculated in the pivot table.
         *
         * @param columnToCalculate one or more column names to calculate
         * @return the current instance of {@link BuilderPivot} for chaining
         */
        public BuilderPivot setColumnToCalculate(String... columnToCalculate) {
            this.pivotBuild.setColumnToCalculate(columnToCalculate);
            return this;
        }

        /**
         * Sets the name of the sheet where the pivot table will be placed.
         *
         * @param sheet the name of the target sheet
         * @return the current instance of {@link BuilderPivot} for chaining
         */
        public BuilderPivot setSheet(String sheet) {
            this.pivotBuild.setSheet(sheet);
            return this;
        }

        /**
         * Creates a new {@link BuilderSpecialField} instance for configuring special fields.
         * The new builder is linked to the current pivot configuration.
         *
         * @return a new {@link BuilderSpecialField} instance
         */
        public BuilderSpecialField setSpecialField() {
            return new BuilderSpecialField(this);
        }

        /**
         * Finalizes the configuration and sets the built {@link Pivot} object
         * in the parent class or context.
         */
        public void build() {
            pivot = this.pivotBuild;
        }
    }

}

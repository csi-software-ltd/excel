package pl.touk.excel.export.abilities

import org.codehaus.groovy.runtime.NullObject
import pl.touk.excel.export.Formatters
import pl.touk.excel.export.XlsxExporter
import pl.touk.excel.export.getters.Getter
import org.apache.poi.ss.usermodel.Sheet

import java.sql.Timestamp

@Category(XlsxExporter)
class RowManipulationAbility {
    private static final handledPropertyTypes = [String, Getter, Date, Boolean, Timestamp, NullObject, Long, Integer, BigDecimal, BigInteger, Byte, Double, Float, Short]

    XlsxExporter fillHeader(List properties) {
        fillRow(Formatters.convertSafelyFromGetters(properties), 0)
    }

    XlsxExporter fillRow(List<Object> properties) {
        fillRow(properties, 1, true)
    }

    XlsxExporter fillRow(List<Object> properties, int rowNumber, Boolean isAddRow) {
        fillRowWithValues(properties, rowNumber, isAddRow)
    }

    XlsxExporter fillRowWithValues(List<Object> properties, int rowNumber, Boolean isAddRow) {
        properties.eachWithIndex { Object property, int index ->
            def propertyToBeInserted = (property == null ? "" : property)
            RowManipulationAbility.verifyPropertyTypeCanBeHandled(property)
            putCellValue(rowNumber, index, propertyToBeInserted)
        }
        if (isAddRow)
            shiftRows(rowNumber,sheet,this)
        this
    }

    XlsxExporter add(List<Object> objects, List<Object> selectedProperties) {
        add(objects, selectedProperties, 1)
    }

    XlsxExporter add(List<Object> objects, List<Object> selectedProperties, int rowNumber) {
        objects.eachWithIndex() { Object object, int index ->
            this.add(object, selectedProperties, rowNumber + index)
        }
        this
    }

    XlsxExporter add(Object object, List<Object> selectedProperties, int rowNumber) {
        List<Object> properties = getPropertiesFromObject(object, Formatters.convertSafelyToGetters(selectedProperties))
        fillRow(properties, rowNumber)
    }

    private static List<Object> getPropertiesFromObject(Object object, List<Getter> selectedProperties) {
        selectedProperties.collect { it.getFormattedValue(object) }
    }

    private static void verifyPropertyTypeCanBeHandled(Object property) {
        if(!(property.getClass() in handledPropertyTypes)) {
            throw new IllegalArgumentException("Properties should by of types: " + handledPropertyTypes + ". Found " + property.getClass())
        }
    }

    private static void shiftRows(int rowNumber, Sheet sheet, def reporter) {
        sheet.shiftRows(rowNumber+1, sheet.getLastRowNum(), 1)
        (0..<CellManipulationAbility.getOrCreateRow(rowNumber, sheet).getLastCellNum()).each{
            reporter.getCellAt(rowNumber+1, it).setCellStyle(reporter.getCellAt(rowNumber, it).getCellStyle())
        }
    }

}

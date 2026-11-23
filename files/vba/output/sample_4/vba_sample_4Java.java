import java.util.*;

public class vba_sample_4Java {

    /* Adapter interfaces ---------------------------------------------------- */

    public interface ApplicationAdapter {
        XlCalculation getCalculation();
        void setCalculation(XlCalculation calc);
        boolean getEnableEvents();
        void setEnableEvents(boolean enable);
        boolean getScreenUpdating();
        void setScreenUpdating(boolean updating);
        String getPathSeparator();
        void showMessage(String text, MessageType type);
    }

    public interface WorkbookAdapter {
        WorksheetAdapter getWorksheet(int index);
        WorksheetAdapter getWorksheet(String name);
        void close(boolean saveChanges);
        void saveAs(String filename, FileFormat format);
    }

    public interface WorksheetAdapter {
        CellRangeAdapter getCells();
        CellRangeAdapter getRange(String address);
        void setName(String name);
        void clearContents();
        void autoFitColumns(String columns);
        String getName();
    }

    public interface CellRangeAdapter {
        int getRowsCount();
        int getColumnsCount();
        Object getValue();
        void setValue(Object value);
        CellRangeAdapter offset(int rows, int cols);
        CellRangeAdapter resize(int rows, int cols);
        CellRangeAdapter end(Direction dir);
    }

    public interface FileSystemAdapter {
        boolean folderExists(String path);
        List<String> listFiles(String path, String pattern);
    }

    public interface WorkbookFactory {
        WorkbookAdapter open(String filename, boolean readOnly);
        WorkbookAdapter add(boolean visible);
    }

    /* Supporting types ----------------------------------------------------- */

    public enum XlCalculation { xlCalculationManual, xlCalculationAutomatic }

    public enum FileFormat { xlOpenXMLWorkbook }

    public enum Direction { xlUp }

    public enum MessageType { vbCritical, vbInformation }

    /* Main procedure -------------------------------------------------------- */

    public void ConsolidateDepartmentFiles(String inputFolder, String outputPath,
                                           ApplicationAdapter app,
                                           WorkbookAdapter thisWorkbook,
                                           WorkbookFactory workbookFactory,
                                           FileSystemAdapter fs) {
        String fileName = "";
        WorkbookAdapter sourceBook = null;
        WorksheetAdapter sourceSheet = null;
        WorksheetAdapter targetSheet = null;
        WorkbookAdapter outputBook = null;
        CellRangeAdapter sourceRange = null;
        long nextRow = 1;
        long lastRow = 0;
        long processedFiles = 0;
        XlCalculation savedCalculation = app.getCalculation();
        boolean savedEvents = app.getEnableEvents();
        boolean savedScreenUpdating = app.getScreenUpdating();

        try {
            /* Validation */
            if (!outputPath.toLowerCase().endsWith(".xlsx")) {
                throw new RuntimeException("Output path must use the .xlsx extension.");
            }
            if (outputPath.equalsIgnoreCase(thisWorkbook.getWorksheet(1).getName())) {
                throw new RuntimeException("Output cannot replace the macro workbook.");
            }
            if (!inputFolder.endsWith(app.getPathSeparator())) {
                inputFolder = inputFolder + app.getPathSeparator();
            }
            if (!fs.folderExists(inputFolder)) {
                throw new RuntimeException("Input folder is unavailable.");
            }

            targetSheet = thisWorkbook.getWorksheet("Consolidated");
            app.setScreenUpdating(false);
            app.setEnableEvents(false);
            app.setCalculation(XlCalculation.xlCalculationManual);
            targetSheet.clearContents();
            nextRow = 1;

            List<String> files = fs.listFiles(inputFolder, "*.xlsx");
            for (String fn : files) {
                fileName = fn;
                if (!fileName.equalsIgnoreCase(thisWorkbook.getWorksheet(1).getName())) {
                    sourceBook = workbookFactory.open(inputFolder + fileName, true);
                    sourceSheet = sourceBook.getWorksheet(1);
                    lastRow = sourceSheet.getCells().end(Direction.xlUp).getRowsCount();

                    if (lastRow >= 1) {
                        sourceRange = sourceSheet.getRange("A1:F" + lastRow);
                        if (nextRow == 1) {
                            targetSheet.getRange("A" + nextRow)
                                    .resize(sourceRange.getRowsCount(), sourceRange.getColumnsCount())
                                    .setValue(sourceRange.getValue());
                            nextRow += sourceRange.getRowsCount();
                        } else if (sourceRange.getRowsCount() > 1) {
                            targetSheet.getRange("A" + nextRow)
                                    .resize(sourceRange.getRowsCount() - 1, sourceRange.getColumnsCount())
                                    .setValue(
                                            sourceRange.offset(1, 0)
                                                    .resize(sourceRange.getRowsCount() - 1, sourceRange.getColumnsCount())
                                                    .getValue());
                            nextRow += sourceRange.getRowsCount() - 1;
                        }
                    }
                    sourceBook.close(false);
                    sourceBook = null;
                    processedFiles++;
                }
            }

            targetSheet.autoFitColumns("A:F");
            outputBook = workbookFactory.add(false);
            outputBook.getWorksheet(1).setName("Consolidated");
            if (nextRow > 1) {
                outputBook.getWorksheet(1)
                        .getRange("A1:F" + (nextRow - 1))
                        .setValue(targetSheet.getRange("A1:F" + (nextRow - 1)).getValue());
            }
            outputBook.saveAs(outputPath, FileFormat.xlOpenXMLWorkbook);
            outputBook.close(false);
            outputBook = null;
        } catch (Exception e) {
            if (sourceBook != null) sourceBook.close(false);
            if (outputBook != null) outputBook.close(false);
            app.showMessage("Consolidation failed for '" + fileName + "': " + e.getMessage(),
                    MessageType.vbCritical);
        } finally {
            app.setCalculation(savedCalculation);
            app.setEnableEvents(savedEvents);
            app.setScreenUpdating(savedScreenUpdating);
            if (/* no exception */ true) {
                app.showMessage("Consolidated " + processedFiles + " department files.",
                        MessageType.vbInformation);
            }
        }
    }
}
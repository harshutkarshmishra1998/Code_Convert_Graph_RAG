import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class vba_sample_4 {
    private vba_sample_4() { }

    public static int ConsolidateDepartmentFiles(Path inputFolder, Path outputPath) throws IOException {
        if (!outputPath.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
            throw new IllegalArgumentException("Output path must use the .xlsx extension.");
        }
        Path folder = inputFolder.toAbsolutePath().normalize();
        if (!Files.isDirectory(folder)) {
            throw new IllegalArgumentException("Input folder is unavailable: " + folder);
        }
        Path output = outputPath.toAbsolutePath().normalize();
        List<Path> inputFiles;
        try (Stream<Path> listed = Files.list(folder)) {
            inputFiles = listed.filter(Files::isRegularFile)
                .filter(path -> path.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".xlsx"))
                .map(path -> path.toAbsolutePath().normalize())
                .filter(path -> !path.equals(output))
                .sorted(Comparator.comparing(path -> path.getFileName().toString(),
                    String.CASE_INSENSITIVE_ORDER))
                .toList();
        }

        try (Workbook outputBook = new XSSFWorkbook()) {
            Sheet target = outputBook.createSheet("Consolidated");
            int nextRow = 0;
            int processedFiles = 0;

            for (Path inputFile : inputFiles) {
                try (Workbook sourceBook = VbaExcelSupport.open(inputFile)) {
                    if (sourceBook.getNumberOfSheets() == 0) {
                        throw new IllegalArgumentException("Workbook has no worksheets: " + inputFile);
                    }
                    Sheet source = sourceBook.getSheetAt(0);
                    int lastSourceRow = VbaExcelSupport.lastRowInColumn(source, 0);
                    int firstSourceRow = processedFiles == 0 ? 0 : 1;
                    if (lastSourceRow >= firstSourceRow) {
                        for (int sourceRow = firstSourceRow; sourceRow <= lastSourceRow; sourceRow++) {
                            for (int column = 0; column < 6; column++) {
                                Cell from = VbaExcelSupport.cell(source, sourceRow, column);
                                Cell to = VbaExcelSupport.writableCell(target, nextRow, column);
                                VbaExcelSupport.copyValue(from, to);
                            }
                            nextRow++;
                        }
                    }
                    processedFiles++;
                }
            }

            VbaExcelSupport.autoSize(target, 6);
            VbaExcelSupport.writeAtomically(outputBook, output);
            System.out.printf("Consolidated %d department files.%n", processedFiles);
            return processedFiles;
        }
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 1 || args.length > 2) {
            System.err.println("Usage: java vba_sample_4 <input-folder> [output.xlsx]");
            System.exit(2);
        }
        Path folder = Path.of(args[0]);
        Path output = args.length == 2 ? Path.of(args[1]) : folder.resolve("consolidated.xlsx");
        ConsolidateDepartmentFiles(folder, output);
    }
}

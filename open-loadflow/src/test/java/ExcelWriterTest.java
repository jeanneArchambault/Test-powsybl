import com.powsybl.commons.datasource.FileDataSource;
import com.powsybl.iidm.network.Network;
import com.powsybl.iidm.network.NetworkFactory;
import com.powsybl.loadflow.LoadFlow;
import com.powsybl.loadflow.LoadFlowParameters;
import com.powsybl.loadflow.LoadFlowResult;
import com.powsybl.math.matrix.SparseMatrixFactory;
import com.powsybl.matpower.converter.MatpowerImporter;
import com.powsybl.openloadflow.CompareKnitroToNewtonRaphson;
import com.powsybl.openloadflow.OpenLoadFlowParameters;
import com.powsybl.openloadflow.OpenLoadFlowProvider;
import com.powsybl.openloadflow.ac.solver.AcSolverType;
import net.jafama.DoubleWrapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.management.ManagementFactory;
import java.lang.management.MemoryUsage;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.List;
import java.util.Properties;

public class ExcelWriterTest {

    private List<String> cases;
    private Path basePath;
    private Path outputPath;

    public ExcelWriterTest(List<String> cases, Path basePath, Path outputPath) {
        this.cases = cases;
        this.basePath = basePath;
        this.outputPath = outputPath;
    }

    public void runCases() throws IOException {
        // Create Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Case Results");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Database Name");
        headerRow.createCell(1).setCellValue("Knitro Converged");
        headerRow.createCell(2).setCellValue("Newton-Raphson Converged");
        headerRow.createCell(3).setCellValue("Knitro Execution Time (s)");
        headerRow.createCell(4).setCellValue("Newton-Raphson Execution Time (s)");
        headerRow.createCell(5).setCellValue("Knitro Error");
        headerRow.createCell(6).setCellValue("Newton-Raphson Error");
        headerRow.createCell(7).setCellValue("Knitro Error In Newton Raphson");
        headerRow.createCell(8).setCellValue("Knitro Iterations");
        headerRow.createCell(9).setCellValue("Newton Raphson Iterations");
        headerRow.createCell(10).setCellValue("Knitro Memory Used (MB)");
        headerRow.createCell(11).setCellValue("Newton-Raphson Memory Used (MB)");


        int rowNum = 1;
        for (String caseName : cases) {
            CaseResults results = runCase(caseName); // Modified to return CaseResults object

            // Write results to the Excel file
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(caseName);
            row.createCell(1).setCellValue(results.isKnitroConverged() ? "Yes" : "No");
            row.createCell(2).setCellValue(results.isNewtonRaphsonConverged() ? "Yes" : "No");
            row.createCell(3).setCellValue(results.getKnitroDurationInSeconds());
            row.createCell(4).setCellValue(results.getNewtonRaphsonDurationInSeconds());
            row.createCell(5).setCellValue(results.getKnitroError());
            row.createCell(6).setCellValue(results.getNewtonRaphsonError());
            row.createCell(7).setCellValue(results.getErrorKnitroInNewtonRaphson());
            row.createCell(8).setCellValue(results.getKnitroIterations());
            row.createCell(9).setCellValue(results.getNewtonRaphsonIterations());
            row.createCell(10).setCellValue(results.getKnitroMemoryUsedInMB());
            row.createCell(11).setCellValue(results.getNewtonRaphsonMemoryUsedInMB());
        }

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream(outputPath.toString())) {
            workbook.write(fileOut);
        }
        workbook.close();
    }

    private CaseResults runCase(String caseName) {
        try {
            // Load network from .mat file
            Properties properties = new Properties();
            properties.put("matpower.import.ignore-base-voltage", false);

            Network network = new MatpowerImporter().importData(
                    new FileDataSource(basePath, caseName),
                    NetworkFactory.findDefault(),
                    properties
            );

            network.write("XIIDM", new Properties(), basePath.resolve(caseName + "_output"));

            // Run Knitro
            LoadFlowParameters parametersKnitro = new LoadFlowParameters();
            OpenLoadFlowParameters parametersExtKnitro = OpenLoadFlowParameters.create(parametersKnitro)
                    .setAcSolverType(AcSolverType.KNITRO);
            configureLoadFlowParameters(parametersKnitro, parametersExtKnitro);

            LoadFlow.Runner loadFlowRunner = new LoadFlow.Runner(new OpenLoadFlowProvider(new SparseMatrixFactory()));
            // execution time
            Instant startKnitro = Instant.now();
            // memory
            long startKnitroMemory = getUsedMemory();
            LoadFlowResult knitroResult = loadFlowRunner.run(network, parametersKnitro);
            long endKnitroMemory = getUsedMemory();
            Instant endKnitro = Instant.now();
            Duration durationKnitro = Duration.between(startKnitro, endKnitro);
            // convergence
            boolean knitroConverged = knitroResult.getComponentResults().get(0).getStatus() == LoadFlowResult.ComponentResult.Status.CONVERGED;
            double knitroDurationInSeconds = durationKnitro.toSeconds();
            double knitroMemoryUsedInMB = (endKnitroMemory - startKnitroMemory) / 1024.0 / 1024.0;
            // iterations
            int knitroIterations = knitroResult.getComponentResults().get(0).getIterationCount();
            // error
            String knitroError = knitroResult.getComponentResults().get(0).getMetrics().get("error");

            // Run Newton Raphson with Knitro result to get NR error
            parametersExtKnitro.setVoltageInitModeOverride(OpenLoadFlowParameters.VoltageInitModeOverride.NONE);
            LoadFlowResult compareKnitroToNewtonRaphson = CompareKnitroToNewtonRaphson.runComparison(loadFlowRunner, parametersKnitro, parametersExtKnitro, network);
            String errorKnitroInNewtonRaphson = compareKnitroToNewtonRaphson.getComponentResults().get(0).getMetrics().get("initialError");

            // Run Newton-Raphson
            LoadFlowParameters parametersNR = new LoadFlowParameters();
            OpenLoadFlowParameters parametersExtNR = OpenLoadFlowParameters.create(parametersNR)
                    .setAcSolverType(AcSolverType.NEWTON_RAPHSON);
            configureLoadFlowParameters(parametersNR, parametersExtNR);
            // execution time
            Instant startNR = Instant.now();
            // memory
            long startNRMemory = getUsedMemory();
            LoadFlowResult newtonRaphsonResult = loadFlowRunner.run(network, parametersNR);
            long endNRMemory = getUsedMemory();
            Instant endNR = Instant.now();
            Duration durationNR = Duration.between(startNR, endNR);
            // convergence
            boolean newtonRaphsonConverged = newtonRaphsonResult.getComponentResults().get(0).getStatus() == LoadFlowResult.ComponentResult.Status.CONVERGED;
            double newtonRaphsonDurationInSeconds = durationNR.toSeconds();
            double newtonRaphsonMemoryUsedInMB = (endNRMemory - startNRMemory) / 1024.0 / 1024.0;
            // iterations
            int newtonRaphsonIterations = newtonRaphsonResult.getComponentResults().get(0).getIterationCount();
            // error
            String newtonRaphsonError = newtonRaphsonResult.getComponentResults().get(0).getMetrics().get("error");

            return new CaseResults(knitroConverged, knitroDurationInSeconds, knitroMemoryUsedInMB, knitroIterations, knitroError,
                    newtonRaphsonConverged, newtonRaphsonDurationInSeconds, newtonRaphsonMemoryUsedInMB, newtonRaphsonIterations, newtonRaphsonError, errorKnitroInNewtonRaphson);


        } catch (Exception e) {
            System.err.println("Error running case " + caseName + ": " + e.getMessage());
            return new CaseResults(false, 0, 0, -1, "-1", false, 0, 0, -1, "-1", "-1");
        }
    }

    private long getUsedMemory() {
        MemoryUsage heapMemoryUsage = ManagementFactory.getMemoryMXBean().getHeapMemoryUsage();
        return heapMemoryUsage.getUsed();
    }

    private void configureLoadFlowParameters(LoadFlowParameters parameters, OpenLoadFlowParameters parametersExt) {
//        parameters.setBalanceType(LoadFlowParameters.BalanceType.PROPORTIONAL_TO_LOAD);
        parameters.setDistributedSlack(false).setUseReactiveLimits(false);
        parameters.getExtension(OpenLoadFlowParameters.class).setSvcVoltageMonitoring(false);
        parameters.setVoltageInitMode(LoadFlowParameters.VoltageInitMode.DC_VALUES);
        parametersExt.setVoltageInitModeOverride(OpenLoadFlowParameters.VoltageInitModeOverride.FULL_VOLTAGE);

        if (parametersExt.getAcSolverType() == AcSolverType.KNITRO) {
            parametersExt.setGradientComputationModeKnitro(1) // user routine
                    .setGradientUserRoutineKnitro(2);   // jac 2
//            parametersExt.setKnitroSolverConvEpsPerEq(Math.pow(10,10));
        }
    }

    private static class CaseResults {
        private final boolean knitroConverged;
        private final double knitroDurationInSeconds;
        private final double knitroMemoryUsedInMB;
        private final boolean newtonRaphsonConverged;
        private final double newtonRaphsonDurationInSeconds;
        private final double newtonRaphsonMemoryUsedInMB;
        private final int knitroIterations;
        private final int newtonRaphsonIterations;
        private String knitroError;
        private String newtonRaphsonError;
        private String errorKnitroInNewtonRaphson;


        public CaseResults(boolean knitroConverged, double knitroDurationInSeconds, double knitroMemoryUsedInMB, int knitroIterations, String knitroError,
                           boolean newtonRaphsonConverged, double newtonRaphsonDurationInSeconds, double newtonRaphsonMemoryUsedInMB, int newtonRaphsonIterations, String newtonRaphsonError, String errorKnitroInNewtonRaphson) {
            this.knitroConverged = knitroConverged;
            this.knitroDurationInSeconds = knitroDurationInSeconds;
            this.knitroMemoryUsedInMB = knitroMemoryUsedInMB;
            this.newtonRaphsonConverged = newtonRaphsonConverged;
            this.newtonRaphsonDurationInSeconds = newtonRaphsonDurationInSeconds;
            this.newtonRaphsonMemoryUsedInMB = newtonRaphsonMemoryUsedInMB;
            this.knitroIterations = knitroIterations;
            this.newtonRaphsonIterations = newtonRaphsonIterations;
            this.knitroError = knitroError;
            this.newtonRaphsonError = newtonRaphsonError;
            this.errorKnitroInNewtonRaphson = errorKnitroInNewtonRaphson;
        }

        public boolean isKnitroConverged() {
            return knitroConverged;
        }

        public double getKnitroDurationInSeconds() {
            return knitroDurationInSeconds;
        }

        public double getKnitroMemoryUsedInMB() {
            return knitroMemoryUsedInMB;
        }

        public int getKnitroIterations() {return knitroIterations; }

        public boolean isNewtonRaphsonConverged() {
            return newtonRaphsonConverged;
        }

        public double getNewtonRaphsonDurationInSeconds() {
            return newtonRaphsonDurationInSeconds;
        }

        public double getNewtonRaphsonMemoryUsedInMB() {
            return newtonRaphsonMemoryUsedInMB;
        }

        public int getNewtonRaphsonIterations() {
            return newtonRaphsonIterations;
        }

        public String getKnitroError() {
            return knitroError;
        }

        public String getNewtonRaphsonError() {
            return newtonRaphsonError;
        }

        public String getErrorKnitroInNewtonRaphson() {
            return errorKnitroInNewtonRaphson;
        }


    }

    public static void main(String[] args) throws IOException {
//        List<String> cases = List.of("case6495rte", "case6470rte", "case2868rte"); // Case names

        List<String> cases = List.of("case14","case57","case300","case1888rte","case1951rte","case2848rte","case2868rte","case6468rte","case6470rte","case6495rte","case6515rte"); // Case names "case14","case57","case300","case1888rte","case1951rte","case2848rte","case2868rte","case6468rte","case6470rte","case6495rte","case6515rte"
//        List<String> cases = List.of("case6470rte"); // Case names "case14","case57","case300","case1888rte","case1951rte","case2848rte","case2868rte","case6468rte","case6470rte","case6495rte","case6515rte"
        Path basePath = Path.of("C:", "Users", "jarchambault", "Downloads");
        Path outputPath = Path.of("C:", "Users", "jarchambault", "Downloads", "results.xlsx");

        ExcelWriterTest runner = new ExcelWriterTest(cases, basePath, outputPath);
        runner.runCases();
    }
}
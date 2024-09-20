package matpower;

import com.powsybl.commons.datasource.FileDataSource;
import com.powsybl.iidm.network.Network;
import com.powsybl.iidm.network.NetworkFactory;
import com.powsybl.loadflow.LoadFlow;
import com.powsybl.loadflow.LoadFlowParameters;
import com.powsybl.loadflow.LoadFlowResult;
import com.powsybl.math.matrix.DenseMatrixFactory;
import com.powsybl.matpower.converter.MatpowerImporter;
import com.powsybl.openloadflow.OpenLoadFlowParameters;
import com.powsybl.openloadflow.OpenLoadFlowProvider;
import com.powsybl.openloadflow.ac.solver.AcSolverType;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.nio.file.Path;
import java.util.Properties;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * @author Jeanne Archambault {@literal <jeanne.archambault at artelys.com>}
 */

public class MatPowerTest {

    private LoadFlow.Runner loadFlowRunner;

    private LoadFlowParameters parameters;

    private OpenLoadFlowParameters parametersExt;

    @BeforeEach
    void setUp() {
        parameters = new LoadFlowParameters();
        parametersExt = OpenLoadFlowParameters.create(parameters)
                .setAcSolverType(AcSolverType.KNITRO);
        // Sparse matrix solver only
        loadFlowRunner = new LoadFlow.Runner(new OpenLoadFlowProvider(new DenseMatrixFactory()));
        // No OLs
        parameters.setBalanceType(LoadFlowParameters.BalanceType.PROPORTIONAL_TO_LOAD);
        parameters.setDistributedSlack(false)
                .setUseReactiveLimits(false);
        parameters.getExtension(OpenLoadFlowParameters.class)
                .setSvcVoltageMonitoring(false);
    }

    @Test
    void case1951Rte() {
        // Load network from .mat file
        Properties properties = new Properties();
        // We want base voltages to be taken into account
        properties.put("matpower.import.ignore-base-voltage", false);
        Network network = new MatpowerImporter().importData(
                new FileDataSource(Path.of("C:", "Users", "jarchambault", "Downloads"), "case1951rte"),
                NetworkFactory.findDefault(), properties);
        network.write("XIIDM", new Properties(), Path.of("C:", "Users", "jarchambault", "Downloads", "case1951rte"));
        LoadFlowResult knitroResult = loadFlowRunner.run(network, parameters);

        assertEquals(LoadFlowResult.ComponentResult.Status.CONVERGED, knitroResult.getComponentResults().get(0).getStatus());
    }

    @Test
    void case57() {
        // Load network from .mat file
        Properties properties = new Properties();
        // We want base voltages to be taken into account
        properties.put("matpower.import.ignore-base-voltage", false);
        Network network = new MatpowerImporter().importData(
                new FileDataSource(Path.of("C:", "Users", "jarchambault", "Downloads"), "case57"),
                NetworkFactory.findDefault(), properties);
        network.write("XIIDM", new Properties(), Path.of("C:", "Users", "jarchambault", "Downloads", "case57"));
        LoadFlowResult knitroResult = loadFlowRunner.run(network, parameters);

        assertEquals(LoadFlowResult.ComponentResult.Status.CONVERGED, knitroResult.getComponentResults().get(0).getStatus());
    }

    @Test
    void case1888Rte() {
        // Load network from .mat file
        Properties properties = new Properties();
        // We want base voltages to be taken into account
        properties.put("matpower.import.ignore-base-voltage", false);
        Network network = new MatpowerImporter().importData(
                new FileDataSource(Path.of("C:", "Users", "jarchambault", "Downloads"), "case1888rte"),
                NetworkFactory.findDefault(), properties);
        network.write("XIIDM", new Properties(), Path.of("C:", "Users", "jarchambault", "Downloads", "case1888rte"));
        LoadFlowResult knitroResult = loadFlowRunner.run(network, parameters);

        assertEquals(LoadFlowResult.ComponentResult.Status.CONVERGED, knitroResult.getComponentResults().get(0).getStatus());
    }

    @Test
    void caseImported() {
        // Load network from .mat file
        Properties properties = new Properties();
        // We want base voltages to be taken into account
        properties.put("matpower.import.ignore-base-voltage", false);
        Network network = new MatpowerImporter().importData(
                new FileDataSource(Path.of("C:", "Users", "jarchambault", "Downloads"), "case14"),
                NetworkFactory.findDefault(), properties);
        network.write("XIIDM", new Properties(), Path.of("C:", "Users", "jarchambault", "Downloads", "case14" +
                ""));
        LoadFlowResult knitroResult = loadFlowRunner.run(network, parameters);

        assertEquals(LoadFlowResult.ComponentResult.Status.CONVERGED, knitroResult.getComponentResults().get(0).getStatus());
    }
}

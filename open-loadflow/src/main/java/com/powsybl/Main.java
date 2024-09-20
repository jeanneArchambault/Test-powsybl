package com.powsybl;

import com.powsybl.ieeecdf.converter.IeeeCdfNetworkFactory;
import com.powsybl.iidm.network.Network;
import com.powsybl.loadflow.LoadFlow;
import com.powsybl.loadflow.LoadFlowParameters;
import com.powsybl.math.matrix.DenseMatrixFactory;
import com.powsybl.openloadflow.OpenLoadFlowParameters;
import com.powsybl.openloadflow.OpenLoadFlowProvider;
import com.powsybl.openloadflow.ac.solver.AcSolverType;

public class Main {
    public static void main(String[] args) {
        a();
    }

    static void a() {
        Network network = IeeeCdfNetworkFactory.create9();
        LoadFlow.Runner runner = new LoadFlow.Runner(new OpenLoadFlowProvider(new DenseMatrixFactory()));
        LoadFlowParameters lfParameters = new LoadFlowParameters();
        OpenLoadFlowParameters olfParameters = OpenLoadFlowParameters.create(lfParameters).setAcSolverType(AcSolverType.KNITRO);
        runner.run(network, lfParameters);
    }
}

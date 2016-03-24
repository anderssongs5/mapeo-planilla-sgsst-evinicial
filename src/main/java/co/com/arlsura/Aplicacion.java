package co.com.arlsura;

import java.io.IOException;
import java.net.URISyntaxException;
import java.util.Date;
import java.util.List;

import co.com.arlsura.archivoexcel.Lector;
import co.com.arlsura.archivoexcel.Procesador;
import co.com.arlsura.dto.PlanillaOriginal;

public class Aplicacion {

    public static void main(String args[]) throws IOException, URISyntaxException {
        System.out.println(new Date());
        Lector archivoExcelLector = new Lector();
        // List<PlanillaOriginal> planillas = archivoExcelLector
        // .leerArchivos("C:\\Users\\Andersson\\Desktop\\Temporales\\EvaluacionesIniciales\\Parte1");

        List<PlanillaOriginal> planillasCarpetas = archivoExcelLector
                .leerCarpetas("C:\\Users\\Andersson\\Desktop\\Temporales\\EvaluacionesIniciales\\Parte4\\NUEVAS");

        Procesador procesador = new Procesador();
        // procesador.mapearPlanilla(planillas);
        procesador.mapearPlanilla(planillasCarpetas);

        System.out.println(new Date());

        // System.out.println(planillas.size());
        System.out.println(planillasCarpetas.size());
    }
}

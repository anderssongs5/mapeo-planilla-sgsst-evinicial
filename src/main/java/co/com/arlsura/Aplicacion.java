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
        try {
            String carpetaPlanillas = args[0];
            String rutaDestino = args[1];
            String nombreArchivoSalida = args[2];

            System.out.println(new Date());
            Lector archivoExcelLector = new Lector();
            List<PlanillaOriginal> planillas = archivoExcelLector.leerArchivos(carpetaPlanillas);

            Procesador procesador = new Procesador();
            procesador.mapearPlanilla(planillas, rutaDestino, nombreArchivoSalida);

            System.out.println(new Date());

            System.out.println(planillas.size());
        } catch (Exception e) {
            System.out.println("Por favor verifique los parámetros iniciales para ejecutar la aplicación correctamente.");
            e.printStackTrace();
        }
    }
}

package co.com.arlsura.archivoexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import co.com.arlsura.dto.PlanillaOriginal;

public class Lector {

    // private static final String SIN_IMPLEMENTAR = "0.0";
    // private static final String SIN_IMPLEMENTAR0 = "0";
    private static final String IMPLEMENTADO = "2.0";
    private static final String IMPLEMENTADO2 = "2";
    private int si = 0;
    private List<String> nombresHojaDecreto = Arrays.asList("Decreto 1443", "Decreto 1072", "Decreto1443",
            "Decreto1072", "Decreto 1443 ", "Decreto 1072 ", "decreto 1443", "decreto1443", "decreto 1072",
            "decreto1072");

    public List<PlanillaOriginal> leerArchivos(String path) throws IOException {
        List<PlanillaOriginal> datosPlanillas = new ArrayList<>();
        List<File> archivos = this.obtenerListaArchivos(new File(path));

        datosPlanillas = this.leerContenidoPlanillas(archivos);

        return datosPlanillas;
    }

    public List<PlanillaOriginal> leerCarpetas(String path) throws IOException {
        List<PlanillaOriginal> datosPlanillas = new ArrayList<>();
        List<File> carpetas = this.obtenerListaArchivos(new File(path));
        List<File> archivos = new ArrayList<>();
        for (File f : carpetas) {
            File[] files = f.listFiles();
            archivos.addAll(Arrays.asList(files));
        }

        datosPlanillas = this.leerContenidoPlanillas(archivos);
        return datosPlanillas;
    }

    private List<PlanillaOriginal> leerContenidoPlanillas(final List<File> archivos) throws IOException {
        List<PlanillaOriginal> datosPlanillas = new ArrayList<>();
        System.out.println("Archivos para leer: " + archivos.size());
        for (int i = 0; i < archivos.size(); i++) {
            try {
                PlanillaOriginal planillaOriginal = new PlanillaOriginal();
                FileInputStream fileInputStream = new FileInputStream(archivos.get(i));
                Workbook workbook = new XSSFWorkbook(fileInputStream);
                FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();

                Sheet portada = workbook.getSheet("Portada");
                if (portada != null) {

                    String contrato = this.obtenerValorCelda(portada.getRow(2).getCell(0),
                            portada.getRow(2).getCell(0).getCellType(), formulaEval);
                    String nit = this.obtenerValorCelda(portada.getRow(2).getCell(1),
                            portada.getRow(2).getCell(1).getCellType(), formulaEval);
                    String nombreArchivo = archivos.get(i).getName();
                    if (null == contrato || contrato.trim().isEmpty()) {
                        planillaOriginal.setSinContrato(true);
                        contrato = nit;
                        if (contrato != null && !contrato.trim().isEmpty()) {
                            planillaOriginal.setNoContrato("NIT");
                        } else {
                            contrato = nombreArchivo;
                            planillaOriginal.setNoContrato("ARCHIVO");
                        }
                    } else {
                        if (!contrato.startsWith("0")) {
                            contrato = "0".concat(contrato);
                        }
                    }
                    planillaOriginal.setContrato(contrato);
                    planillaOriginal.setNit(nit);
                    planillaOriginal.setNombreArchivo(nombreArchivo);

                    this.si = 0;
                    Sheet decreto1072 = null;
                    for (int d = 0; d < this.nombresHojaDecreto.size(); d++) {
                        decreto1072 = workbook.getSheet(this.nombresHojaDecreto.get(d));
                        if (null != decreto1072) {
                            break;
                        }
                    }

                    if (decreto1072 != null) {
                        for (int j = Numeracion.mapaNumeros.get("VALOR_2"); j <= Numeracion.mapaNumeros
                                .get("VALOR_78"); j++) {
                            Row row = decreto1072.getRow(j);
                            this.setDatosPlanillaOriginal(j, row, planillaOriginal, formulaEval);
                        }
                        planillaOriginal.setSi(this.si);

                        workbook.close();
                        fileInputStream.close();

                        datosPlanillas.add(planillaOriginal);

                        System.out.println("Archivo " + (i + 1) + " | " + archivos.get(i).getName() + " | "
                                + archivos.get(i).getAbsolutePath() + " | " + new Date() + " | "
                                + planillaOriginal.getContrato());
                    } else {
                        System.out.println("No se pudo procesar el archivo: " + archivos.get(i).getName() + " | "
                                + archivos.get(i).getAbsolutePath()
                                + " porque la hoja Decreto 1443 o Decreto 1072 no existe");
                    }

                } else {
                    System.out.println("No se pudo procesar el archivo: " + archivos.get(i).getName() + " | "
                            + archivos.get(i).getAbsolutePath() + " porque la hoja Portada no existe");
                }
            } catch (Exception e) {
                System.out.println("Se presentó error leyendo el archivo " + archivos.get(i).getName() + " | "
                        + archivos.get(i).getAbsolutePath());
                e.printStackTrace();
            }
        }
        return datosPlanillas;
    }

    private void setDatosPlanillaOriginal(int posicion, Row row, PlanillaOriginal planillaOriginal,
            FormulaEvaluator formulaEvaluator) {
        String campo = this.obtenerValorCelda(row.getCell(10), row.getCell(10).getCellType(), formulaEvaluator);
        String m = this.mapearValorCampo(campo);
        if ("Si".equals(m)) {
            this.si += 1;
        }
        switch (posicion) {
        case 2:
            planillaOriginal.setDocumentoEscritoPolitica(m);
            break;
        case 3:
            planillaOriginal.setComunicacionPolitica(m);
            break;
        case 4:
            planillaOriginal.setDefinicionResponsabilidades(m);
            break;
        case 5:
            planillaOriginal.setComunicacionResponsabilidades(m);
            break;
        case 6:
            planillaOriginal.setRendicionCuentas(m);
            break;
        case 7:
            planillaOriginal.setPresupuestoSST(m);
            break;
        case 8:
            planillaOriginal.setDefTHSST(m);
            break;
        case 9:
            planillaOriginal.setRecursosTecnicos(m);
            break;
        case 10:
            planillaOriginal.setReqLegalesMatReqLegales(m);
            break;
        case 11:
            planillaOriginal.setPlanTrabAnualCronogramaA(m);
            break;
        case 12:
            planillaOriginal.setPlanTrabAnualCronogramaB(m);
            break;
        case 13:
            planillaOriginal.setCopasst(m);
            break;
        case 14:
            planillaOriginal.setDireccionSST(m);
            break;
        case 15:
            planillaOriginal.setIntegracionOtrosSistmasGestion(m);
            break;
        case 16:
            planillaOriginal.setCapacitacionSSTPersonalCompetencias(m);
            break;
        case 17:
            planillaOriginal.setSocializacionCopasstPC(m);
            break;
        case 18:
            planillaOriginal.setInduccionReinduccionSST1(m);
            break;
        case 19:
            planillaOriginal.setIpevrA(m);
            break;
        case 20:
            planillaOriginal.setIpevrB(m);
            break;
        case 21:
            planillaOriginal.setCondSaludPerfilSocio(m);
            break;
        case 22:
            planillaOriginal.setEstandaresSeguridadOperacionSegura(m);
            break;
        case 23:
            planillaOriginal.setRegistroEntregaEPP(m);
            break;
        case 24:
            planillaOriginal.setReporteInvesATEL(m);
            break;
        case 25:
            planillaOriginal.setIdentificacionAmenazaVulnerabilidad(m);
            break;
        case 26:
            planillaOriginal.setProcOperativosNormalizados1(m);
            break;
        case 27:
            planillaOriginal.setPlanEvacuacionEvalSimuDisPE(m);
            break;
        case 28:
            planillaOriginal.setSve(m);
            break;
        case 29:
            planillaOriginal.setEvalAmbientales(m);
            break;
        case 30:
            planillaOriginal.setPerfilEpiSVE(m);
            break;
        case 31:
            planillaOriginal.setFormatoRegistroInspecciones(m);
            break;
        case 32:
            planillaOriginal.setRegistrosGestionRiesgos(m);
            break;
        case 33:
            planillaOriginal.setConservacionDocumentos(m);
            break;
        case 34:
            planillaOriginal.setComunInternaExternaCanales(m);
            break;
        case 35:
            planillaOriginal.setComunEvaluacionAmbiental(m);
            break;
        case 36:
            planillaOriginal.setAutoevaluacion(m);
            break;
        case 37:
            planillaOriginal.setCumplLegalFortCompoSistemaMejoraContinua(m);
            break;
        case 38:
            planillaOriginal.setObjetivosControlRiesgos(m);
            break;
        case 39:
            planillaOriginal.setIndicadoresEstructuraProcesoResultado(m);
            break;
        case 40:
            planillaOriginal.setMetasAnuales(m);
            break;
        case 41:
            planillaOriginal.setComunicacionObjetivos(m);
            break;
        case 42:
            planillaOriginal.setFichaIndicadoresMatrizIndicadores(m);
            break;
        case 43:
            planillaOriginal.setIndicadoresEstructura(m);
            break;
        case 44:
            planillaOriginal.setIndicadoresProceso(m);
            break;
        case 45:
            planillaOriginal.setIndicadoresResultado(m);
            break;
        case 46:
            planillaOriginal.setProcedGestionPeligrosRiesgos(m);
            break;
        case 47:
            planillaOriginal.setTratamientoRiesgos(m);
            break;
        case 48:
            planillaOriginal.setAdminEPPParteMatrizEPP(m);
            break;
        case 49:
            planillaOriginal.setSocializacionPartesInteresadas(m);
            break;
        case 50:
            planillaOriginal.setPlanMantCorrectivoPreventivo(m);
            break;
        case 51:
            planillaOriginal.setEvalMedicasOcupacionales(m);
            break;
        case 52:
            planillaOriginal.setIdentifAmenaVulneXCT(m);
            break;
        case 53:
            planillaOriginal.setValoracionRiesgosAsociadoAmenazas(m);
            break;
        case 54:
            planillaOriginal.setProcOperativosNormalizados2(m);
            break;
        case 55:
            planillaOriginal.setPlanRespuestaEvenPotencDesastrosos(m);
            break;
        case 56:
            planillaOriginal.setEvaluacionSimulacros(m);
            break;
        case 57:
            planillaOriginal.setCapaciEntrenaPlanEmergencias(m);
            break;
        case 58:
            planillaOriginal.setRealizaEvalSimulacrosAnuales(m);
            break;
        case 59:
            planillaOriginal.setConformacionFuncionamientoBrigadasEmergencia(m);
            break;
        case 60:
            planillaOriginal.setInspeccionEquiposEmergencia(m);
            break;
        case 61:
            planillaOriginal.setPlanAyudaMutua(m);
            break;
        case 62:
            planillaOriginal.setGestionCambio(m);
            break;
        case 63:
            planillaOriginal.setIntegrRequisitosSSTCompras(m);
            break;
        case 64:
            planillaOriginal.setProcedSeleccionEvalContratistas(m);
            break;
        case 65:
            planillaOriginal.setSeguimientoContratistas(m);
            break;
        case 66:
            planillaOriginal.setVerificacionAfiliacionSS(m);
            break;
        case 67:
            planillaOriginal.setInduccionReinduccionSST2(m);
            break;
        case 68:
            planillaOriginal.setInduccionReinduccionContratistas(m);
            break;
        case 69:
            planillaOriginal.setProgramaAuditoriaAnual(m);
            break;
        case 70:
            planillaOriginal.setInformeResultadosAuditoria(m);
            break;
        case 71:
            planillaOriginal.setAlcanceAuditoria(m);
            break;
        case 72:
            planillaOriginal.setRevisionGerenciaAnual1(m);
            break;
        case 73:
            planillaOriginal.setSocializacionCopasst(m);
            break;
        case 74:
            planillaOriginal.setProcInvestInciAccEL(m);
            break;
        case 75:
            planillaOriginal.setSocializacionLeccionesAprendidas(m);
            break;
        case 76:
            planillaOriginal.setInformesPeriodicosGerencia(m);
            break;
        case 77:
            planillaOriginal.setSeguiAccionesCorrectivasPreventivas(m);
            break;
        case 78:
            planillaOriginal.setRevisionGerenciaAnual2(m);
        }
    }

    private String mapearValorCampo(String campo) {
        if (campo != null && (IMPLEMENTADO.equals(campo) || IMPLEMENTADO2.equals(campo))) {
            return "Si";
        } else {
            return "No";
        }
    }

    private String obtenerValorCelda(Cell celda, int tipo, FormulaEvaluator formulaEvaluator) {
        if (null == celda) {
            return null;
        } else if (Cell.CELL_TYPE_NUMERIC == tipo) {
            celda.setCellType(Cell.CELL_TYPE_STRING);
            return celda.getStringCellValue();
        } else if (Cell.CELL_TYPE_STRING == tipo) {
            return celda.getStringCellValue().replace(",", ".");
        } else if (Cell.CELL_TYPE_FORMULA == tipo) {
            String valor = null;
            switch (celda.getCachedFormulaResultType()) {
            case Cell.CELL_TYPE_NUMERIC:
                valor = String.valueOf(celda.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING:
                valor = celda.getRichStringCellValue().toString();
                break;
            }
            return valor;
        } else {
            return "";
        }

    }

    private List<File> obtenerListaArchivos(final File directorio) {
        List<File> archivos = new ArrayList<>();
        File[] archivosPlanillas = directorio.listFiles();
        archivos = Arrays.asList(archivosPlanillas);

        return archivos;
    }
}

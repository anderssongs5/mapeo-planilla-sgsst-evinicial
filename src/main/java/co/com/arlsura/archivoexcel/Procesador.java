package co.com.arlsura.archivoexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import co.com.arlsura.dto.PlanillaOriginal;

public class Procesador {

    public void mapearPlanilla(List<PlanillaOriginal> planillas, String rutaDestino, String nombreArchivo) throws IOException, URISyntaxException {
        System.out.println("Archivos para procesar: " + planillas.size());

        URL url = Procesador.class.getResource("/co/com/arlsura/archivoexcel/resources/NuevaPlanilla.xlsx");
        File file = new File(url.toURI());
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        CellStyle estiloContrato = workbook.createCellStyle();
        estiloContrato.setAlignment(CellStyle.ALIGN_CENTER);
        CellStyle estiloOtrasCeldas = workbook.createCellStyle();
        estiloOtrasCeldas.setAlignment(CellStyle.ALIGN_CENTER);
        CellStyle estiloPorcentaje = workbook.createCellStyle();
        estiloPorcentaje.setAlignment(CellStyle.ALIGN_CENTER);
        Font fontOtros = workbook.createFont();
        Font fontContratoVerde = workbook.createFont();
        fontContratoVerde.setBold(true);
        fontContratoVerde.setColor(IndexedColors.RED.getIndex());
        Font fontContratoRojo = workbook.createFont();
        fontContratoRojo.setBold(true);
        fontContratoRojo.setColor(IndexedColors.RED.getIndex());
        Font fontPorcentaje = workbook.createFont();
        fontPorcentaje.setColor(IndexedColors.RED.getIndex());
        fontPorcentaje.setBold(true);
        estiloPorcentaje.setDataFormat(workbook.createDataFormat().getFormat("0%"));
        estiloPorcentaje.setFont(fontPorcentaje);
        for (int i = 0; i < planillas.size(); i++) {
            this.setDatosMapeoPlanilla(i, sheet, planillas.get(i), workbook, estiloContrato, estiloOtrasCeldas,
                    fontOtros, estiloPorcentaje, fontContratoRojo, fontContratoVerde);
        }

        fileInputStream.close();
        FileOutputStream fileOutputStream = new FileOutputStream(
                new File(rutaDestino + "/" + nombreArchivo + ".xlsx"));

        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    private void setDatosMapeoPlanilla(int i, Sheet sheet, PlanillaOriginal planillaOriginal, Workbook workbook,
            CellStyle estiloContrato, CellStyle estiloOtrasCeldas, Font f, CellStyle estiloPorcentaje,
            Font contratoRojo, Font contratoVerde) {
        try {
            System.out.println("Archivo Guardado: " + (i + 1) + " | " + planillaOriginal.getNombreArchivo());
            Row row = sheet.createRow(i + 1);

            Cell cell = row.createCell(0);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(planillaOriginal.getContrato());
            if (planillaOriginal.getSinContrato()) {
                if (planillaOriginal.getNoContrato().equals("NIT")) {
                    estiloContrato.setFont(contratoVerde);
                } else {
                    estiloContrato.setFont(contratoRojo);
                }
            } else {
                estiloContrato.setFont(f);
            }
            cell.setCellStyle(estiloContrato);

            for (int j = Numeracion.mapaNumeros.get("VALOR_2"); j <= Numeracion.mapaNumeros.get("VALOR_83"); j++) {
                String valor = this.retornarValorCelda(j, planillaOriginal, i + 1);
                Cell celda = row.createCell(j);

                if (j == Numeracion.mapaNumeros.get("VALOR_81")) {
                    celda.setCellStyle(estiloPorcentaje);
                } else {
                    celda.setCellStyle(estiloOtrasCeldas);
                }

                if (j >= Numeracion.mapaNumeros.get("VALOR_79") && j <= Numeracion.mapaNumeros.get("VALOR_81")) {
                    celda.setCellType(Cell.CELL_TYPE_FORMULA);
                    celda.setCellFormula(valor);
                } else {
                    celda.setCellType(Cell.CELL_TYPE_STRING);
                    celda.setCellValue(valor);
                }
            }
        } catch (Exception e) {
            System.out.println("Se presentó error procesando archivo " + planillaOriginal.getNombreArchivo());
            e.printStackTrace();
        }
    }

    private String retornarValorCelda(int j, PlanillaOriginal planillaOriginal, int fila) {
        String valor = null;
        switch (j) {
        case 2:
            valor = planillaOriginal.getDocumentoEscritoPolitica();
            break;
        case 3:
            valor = planillaOriginal.getComunicacionPolitica();
            break;
        case 4:
            valor = planillaOriginal.getDefinicionResponsabilidades();
            break;
        case 5:
            valor = planillaOriginal.getComunicacionResponsabilidades();
            break;
        case 6:
            valor = planillaOriginal.getRendicionCuentas();
            break;
        case 7:
            valor = planillaOriginal.getPresupuestoSST();
            break;
        case 8:
            valor = planillaOriginal.getDefTHSST();
            break;
        case 9:
            valor = planillaOriginal.getRecursosTecnicos();
            break;
        case 10:
            valor = planillaOriginal.getReqLegalesMatReqLegales();
            break;
        case 11:
            valor = planillaOriginal.getPlanTrabAnualCronogramaA();
            break;
        case 12:
            valor = planillaOriginal.getPlanTrabAnualCronogramaB();
            break;
        case 13:
            valor = planillaOriginal.getCopasst();
            break;
        case 14:
            valor = planillaOriginal.getDireccionSST();
            break;
        case 15:
            valor = planillaOriginal.getIntegracionOtrosSistmasGestion();
            break;
        case 16:
            valor = planillaOriginal.getCapacitacionSSTPersonalCompetencias();
            break;
        case 17:
            valor = planillaOriginal.getSocializacionCopasstPC();
            break;
        case 18:
            valor = planillaOriginal.getInduccionReinduccionSST1();
            break;
        case 19:
            valor = planillaOriginal.getIpevrA();
            break;
        case 20:
            valor = planillaOriginal.getIpevrB();
            break;
        case 21:
            valor = planillaOriginal.getCondSaludPerfilSocio();
            break;
        case 22:
            valor = planillaOriginal.getEstandaresSeguridadOperacionSegura();
            break;
        case 23:
            valor = planillaOriginal.getRegistroEntregaEPP();
            break;
        case 24:
            valor = planillaOriginal.getReporteInvesATEL();
            break;
        case 25:
            valor = planillaOriginal.getIdentificacionAmenazaVulnerabilidad();
            break;
        case 26:
            valor = planillaOriginal.getProcOperativosNormalizados1();
            break;
        case 27:
            valor = planillaOriginal.getPlanEvacuacionEvalSimuDisPE();
            break;
        case 28:
            valor = planillaOriginal.getSve();
            break;
        case 29:
            valor = planillaOriginal.getEvalAmbientales();
            break;
        case 30:
            valor = planillaOriginal.getPerfilEpiSVE();
            break;
        case 31:
            valor = planillaOriginal.getFormatoRegistroInspecciones();
            break;
        case 32:
            valor = planillaOriginal.getRegistrosGestionRiesgos();
            break;
        case 33:
            valor = planillaOriginal.getConservacionDocumentos();
            break;
        case 34:
            valor = planillaOriginal.getComunInternaExternaCanales();
            break;
        case 35:
            valor = planillaOriginal.getComunEvaluacionAmbiental();
            break;
        case 36:
            valor = planillaOriginal.getAutoevaluacion();
            break;
        case 37:
            valor = planillaOriginal.getCumplLegalFortCompoSistemaMejoraContinua();
            break;
        case 38:
            valor = planillaOriginal.getObjetivosControlRiesgos();
            break;
        case 39:
            valor = planillaOriginal.getIndicadoresEstructuraProcesoResultado();
            break;
        case 40:
            valor = planillaOriginal.getMetasAnuales();
            break;
        case 41:
            valor = planillaOriginal.getComunicacionObjetivos();
            break;
        case 42:
            valor = planillaOriginal.getFichaIndicadoresMatrizIndicadores();
            break;
        case 43:
            valor = planillaOriginal.getIndicadoresEstructura();
            break;
        case 44:
            valor = planillaOriginal.getIndicadoresProceso();
            break;
        case 45:
            valor = planillaOriginal.getIndicadoresResultado();
            break;
        case 46:
            valor = planillaOriginal.getProcedGestionPeligrosRiesgos();
            break;
        case 47:
            valor = planillaOriginal.getTratamientoRiesgos();
            break;
        case 48:
            valor = planillaOriginal.getAdminEPPParteMatrizEPP();
            break;
        case 49:
            valor = planillaOriginal.getSocializacionPartesInteresadas();
            break;
        case 50:
            valor = planillaOriginal.getPlanMantCorrectivoPreventivo();
            break;
        case 51:
            valor = planillaOriginal.getEvalMedicasOcupacionales();
            break;
        case 52:
            valor = planillaOriginal.getIdentifAmenaVulneXCT();
            break;
        case 53:
            valor = planillaOriginal.getValoracionRiesgosAsociadoAmenazas();
            break;
        case 54:
            valor = planillaOriginal.getProcOperativosNormalizados2();
            break;
        case 55:
            valor = planillaOriginal.getPlanRespuestaEvenPotencDesastrosos();
            break;
        case 56:
            valor = planillaOriginal.getEvaluacionSimulacros();
            break;
        case 57:
            valor = planillaOriginal.getCapaciEntrenaPlanEmergencias();
            break;
        case 58:
            valor = planillaOriginal.getRealizaEvalSimulacrosAnuales();
            break;
        case 59:
            valor = planillaOriginal.getConformacionFuncionamientoBrigadasEmergencia();
            break;
        case 60:
            valor = planillaOriginal.getInspeccionEquiposEmergencia();
            break;
        case 61:
            valor = planillaOriginal.getPlanAyudaMutua();
            break;
        case 62:
            valor = planillaOriginal.getGestionCambio();
            break;
        case 63:
            valor = planillaOriginal.getIntegrRequisitosSSTCompras();
            break;
        case 64:
            valor = planillaOriginal.getProcedSeleccionEvalContratistas();
            break;
        case 65:
            valor = planillaOriginal.getSeguimientoContratistas();
            break;
        case 66:
            valor = planillaOriginal.getVerificacionAfiliacionSS();
            break;
        case 67:
            valor = planillaOriginal.getInduccionReinduccionSST2();
            break;
        case 68:
            valor = planillaOriginal.getInduccionReinduccionContratistas();
            break;
        case 69:
            valor = planillaOriginal.getProgramaAuditoriaAnual();
            break;
        case 70:
            valor = planillaOriginal.getInformeResultadosAuditoria();
            break;
        case 71:
            valor = planillaOriginal.getAlcanceAuditoria();
            break;
        case 72:
            valor = planillaOriginal.getRevisionGerenciaAnual1();
            break;
        case 73:
            valor = planillaOriginal.getSocializacionCopasst();
            break;
        case 74:
            valor = planillaOriginal.getProcInvestInciAccEL();
            break;
        case 75:
            valor = planillaOriginal.getSocializacionLeccionesAprendidas();
            break;
        case 76:
            valor = planillaOriginal.getInformesPeriodicosGerencia();
            break;
        case 77:
            valor = planillaOriginal.getSeguiAccionesCorrectivasPreventivas();
            break;
        case 78:
            valor = planillaOriginal.getRevisionGerenciaAnual2();
            break;
        case 79:
            valor = "COUNTIF(C" + (fila + 1) + ":CA" + (fila + 1) + ",\"Si\")";
            break;
        case 80:
            valor = "COUNTIF(C" + (fila + 1) + ":CA" + (fila + 1) + ",\"No\")";
            break;
        case 81:
            valor = "CB" + (fila + 1) + "/(CB" + (fila + 1) + " + CC" + (fila + 1) + ")";
            break;
        case 82:
            valor = planillaOriginal.getNit();
            break;
        case 83:
            valor = planillaOriginal.getNombreArchivo();
        }
        return valor;
    }
}

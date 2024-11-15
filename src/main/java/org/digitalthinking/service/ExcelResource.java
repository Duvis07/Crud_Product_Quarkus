package org.digitalthinking.service;


import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.Response;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Path("/export")
public class ExcelResource {

    @GET
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Response generateExcelFile() {
        // Crear un flujo de salida de bytes para el archivo Excel
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            // Crear un nuevo archivo de Excel (Workbook) y una hoja (Worksheet)
            Workbook workbook = new Workbook(outputStream, "ExampleWorkbook", "1.0");
            Worksheet sheet = workbook.newWorksheet("Datos");

            // Escribir encabezados en la primera fila
            sheet.value(0, 0, "ID");
            sheet.value(0, 1, "Nombre");
            sheet.value(0, 2, "Edad");

            // Agregar algunas filas de datos de ejemplo
            sheet.value(1, 0, 1);
            sheet.value(1, 1, "Juan Pérez");
            sheet.value(1, 2, 30);

            sheet.value(2, 0, 2);
            sheet.value(2, 1, "Ana Gómez");
            sheet.value(2, 2, 25);

            sheet.value(3, 0, 3);
            sheet.value(3, 1, "Carlos Ruiz");
            sheet.value(3, 2, 40);

            // Finalizar la creación del archivo Excel
            workbook.finish();

            // Retornar el archivo como respuesta HTTP para su descarga
            return Response.ok(outputStream.toByteArray())
                    .header("Content-Disposition", "attachment; filename=\"generated.xlsx\"")
                    .build();
        } catch (IOException e) {
            // Manejo de errores en caso de que falle la generación del archivo
            return Response.serverError().entity("Error al generar el archivo de Excel").build();
        }
    }
}

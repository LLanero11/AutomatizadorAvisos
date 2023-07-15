import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
public class Main {
    //Proceso INICIAL
    public static void main(String[] args) {
        Timer timerInicio = new Timer();
        Timer timerFin = new Timer();
        //El Timer es un componente de java que nos ayuda a ejecutar una funcion cada cierto tiempo
        timerInicio.schedule(new TareaProgramadaInicio(), calculaDelayInicial(), 24 * 60 * 60 * 1000);
        timerFin.schedule(new TareaProgramadaFin(), calculaDelayFinal(), 24 * 60 * 60 * 1000);
    }
    // 1. Metodo Para obtener la hora configurada
    private static long calculaDelayInicial() {
        Calendar calendario = Calendar.getInstance();
        int horaActual = calendario.get(Calendar.HOUR_OF_DAY);
        int minutosActuales = calendario.get(Calendar.MINUTE);
        int segundosActuales = calendario.get(Calendar.SECOND);
        long tiempoEstimado = (20 - horaActual) * 60 * 60 * 1000 + (34 - minutosActuales) * 60 * 1000 - segundosActuales * 1000;
        if (tiempoEstimado < 0) {
            tiempoEstimado += 24 * 60 * 60 * 1000;
        }
        System.out.println(tiempoEstimado);
        return tiempoEstimado;
    }
    private static long calculaDelayFinal() {
        Calendar calendario = Calendar.getInstance();
        int horaActual = calendario.get(Calendar.HOUR_OF_DAY);
        int minutosActuales = calendario.get(Calendar.MINUTE);
        int segundosActuales = calendario.get(Calendar.SECOND);
        long tiempoEstimado = (20 - horaActual) * 60 * 60 * 1000 + (34 - minutosActuales) * 60 * 1000 - segundosActuales * 1000;
        if (tiempoEstimado < 0) {
            tiempoEstimado += 24 * 60 * 60 * 1000;
        }
        System.out.println(tiempoEstimado);
        return tiempoEstimado;
    }

    //2. Si cumple con la ejecucion hora programa se ejecuta el proceso
    static class TareaProgramadaInicio extends TimerTask {
        public void run() {
            List<Alumno> listaAlumnos = leerArchivoExcel("C:/Parcial/DATOS.xlsx");
            ValidaProcesoEnvioSaludo(listaAlumnos,1);



        }
    }

    static class TareaProgramadaFin extends TimerTask {
        public void run() {
            List<Alumno> listaAlumnos = leerArchivoExcel("C:/Parcial/DATOS.xlsx");
            ValidaProcesoEnvioSaludo(listaAlumnos,2);



        }
    }

    //3. Lee el archivo excel
    // Se utilizo la libreria Apache Poi
    public static List<Alumno> leerArchivoExcel(String archivo) {
        // Declara el objeto en una lista
        List<Alumno> listaEmpleados = new ArrayList<Alumno>();
        //Obtiene el archivo excel.
        try (FileInputStream archivoEntrada = new FileInputStream(archivo);
             Workbook libroExcel = WorkbookFactory.create(archivoEntrada)) {
            Sheet hoja = libroExcel.getSheetAt(0);
            int filaInicial = 2; // la fila B3 es la fila 2
            int filaFinal = hoja.getLastRowNum(); // última fila de la hoja
            int columnaInicial = 1; // columna B
            int columnaFinal = 4; // columna E

            boolean salida = false;
            //Obtiene la informacion desde un rango de celdas configurado.
            CellRangeAddress rangoCeldas = new CellRangeAddress(filaInicial, filaFinal, columnaInicial, columnaFinal);
            // Iteración de las filas.
            for (int i = rangoCeldas.getFirstRow(); i <= rangoCeldas.getLastRow(); i++) {
                Row fila = hoja.getRow(i);
                if (fila != null) {
                    String nombre = "";
                    Date fecha = null;
                    String area = "";
                    String correo = "";
                    // Iteración de las columnas de cada fila : Para esta caso son 4 iteraciones por fila.
                    for (int j = rangoCeldas.getFirstColumn(); j <= rangoCeldas.getLastColumn(); j++) {
                        Cell celda = fila.getCell(j);
                        if (celda != null) {
                            switch (j) {
                                case 1:
                                    nombre = celda.getStringCellValue();
                                    break;
                                case 2:
                                    fecha = celda.getDateCellValue();
                                    break;
                                case 3:
                                    area = celda.getStringCellValue();
                                case 4:
                                    correo = celda.getStringCellValue();
                                    break;
                            }
                        }else{
                            salida = true;
                            break;
                        }
                    }

                    if(!salida)
                    {
                        //Alimenta el objeto a las lista con la informacion de los Alumnos
                        listaEmpleados.add(new Alumno(nombre, fecha, area, correo));
                    }else{
                        break;
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return listaEmpleados;
    }

    //4. Valida la informacion
    /*
       Criterios son los siguiente:
       enviar correos
    * */
    private static void ValidaProcesoEnvioSaludo(List<Alumno> listaAlumnos,int tipo)
    {
        Iterator<Alumno> iter = listaAlumnos.iterator();
        while (iter.hasNext()) {
            Alumno emp = iter.next();




            enviarCorreoDeSesion(emp,tipo);


        }

    }

    // 5. Envia el correo a los Alumnos Para que ingresen a las clases
    // Se utilizo la libreria Java Mail
    private static void enviarCorreoDeSesion(Alumno alumno,int tipo) {
        // Configura las propiedades de la sesión de correo.
        Properties propiedades = new Properties();
        propiedades.put("mail.smtp.auth", "true");
        propiedades.put("mail.smtp.starttls.enable", "true");
        propiedades.put("mail.smtp.host", "smtp.gmail.com");
        propiedades.put("mail.smtp.port", "587");
        // Ingresa tus credenciales de correo electrónico de Gmail aquí.
        final String correoUsuario = "peyelomax@gmail.com";
        final String contrasenia = "lggwauysdmpqgfeh";
        String Asunto;
        String Mensaje1;
        String Mensaje2;
        if (tipo==1){
            Asunto ="Es hora de clases, " + alumno.getNombre() + "!";
            Mensaje1 =     "Es hora de clases";
            Mensaje2 = "\n Ingrese a clases " + alumno.getNombre();
        }else {
            Asunto ="Es hora salida, " + alumno.getNombre() + "!";
            Mensaje1 =     "Es hora de salida";
            Mensaje2 = "\n Puede retirarse " + alumno.getNombre();

        }

        // Inicia sesión en el servidor de correo.
        Session sesion = Session.getInstance(propiedades, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(correoUsuario, contrasenia);
            }
        });
        try {
            // Construye el mensaje de correo electrónico.
            Message mensaje = new MimeMessage(sesion);
            mensaje.setFrom(new InternetAddress(correoUsuario));
            mensaje.setRecipients(Message.RecipientType.TO, InternetAddress.parse(alumno.getCorreo()));

            mensaje.setSubject(Asunto);
            // .
            BodyPart mensajeParte0 = new MimeBodyPart();
            mensajeParte0.setText(Mensaje1);
            BodyPart mensajeParte1 = new MimeBodyPart();
            mensajeParte1.setText(Mensaje2);
            // Crea un mensaje compuesto que incluye el texto.
            Multipart mensajeCompuesto = new MimeMultipart();
            mensajeCompuesto.addBodyPart(mensajeParte0);
            mensajeCompuesto.addBodyPart(mensajeParte1);
            mensaje.setContent(mensajeCompuesto);
            // Envía el mensaje de correo electrónico.
            Transport.send(mensaje);
            System.out.println("Mensaje enviado a " + alumno.getCorreo());
        } catch (MessagingException e) {
            e.printStackTrace();
        }
    }
}

package dev.mamasoyjuanito.handling;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Collections;
import java.util.Iterator;
import java.util.Map;
import java.util.UUID;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.github.cliftonlabs.json_simple.JsonObject;

import dev.mamasoyjuanito.utils.Constantes;

public class DocxHandling {

	private Map<String, String> textsToReplace;
	private String originalFilePath;

	private String temporalFileName;

	private XWPFDocument doc;

	public DocxHandling() {
	}

	public DocxHandling(Map<String, String> textsToReplace, String originalFilePath, String destinationPath) {
		this.textsToReplace = textsToReplace;
		this.originalFilePath = originalFilePath;
	}

	public InputStream replaceTexts() throws Exception {
		System.out.println("se inicia replaceText");
		// Validamos que hayamos recibido el path del documento original
		if (originalFilePath == null)
			throw new NullPointerException("Debe proporcionar el path del archivo original");

		// Validamos que hayamos recibido los textos a reemplazar
		if (textsToReplace == null || textsToReplace.isEmpty())
			throw new Exception("Debe proporcionar los textos a reemplazar");

		// Generamos un nombre para nuestro archivo temporal
		temporalFileName = UUID.randomUUID().toString();

		// Abrimos nuestro documento
		doc = new XWPFDocument(new FileInputStream(originalFilePath));
		// System.out.println("estamos en el metodo que reemplaza");
		// comenzamos con la iteracion de los textos a reemplazar, se modifica para
		// utilizar con un arraylist o json
		// Desde base de datos se debe obtener esta informacion y hacerlo dinamico
		// cuando busque valores
		System.out.println("se debe tomar la desicion de que pdf se va a generar");
		Integer tipoReporte = 2;
		// variables globales
		JsonObject json = new JsonObject();
		String[] variablesWord = null;

		if (tipoReporte == 1) {
			json.put("nombre", "Luis Mario 1");
			json.put("edad", "33");
			json.put("prof", "programador");
			json.put("domicilio", "Santa Ana");
			json.put("departamento", "Santa Ana");
			json.put("dui", "041963089");
			json.put("montoAceptado", "1000");
			json.put("cuentaParaCargo", "0355588777444");
			json.put("plazo", "35");
			json.put("cuotaTotal", "40");
			json.put("cuotaCapitalIntereses", "50");
			json.put("seguroVidaD", "200");
			json.put("destino", "reemplazo");
			json.put("dia", "varDia");
			json.put("tasaNominalEnLetras", "Veinte");
			json.put("tasaNominalNumeros", "20");
			json.put("tasaEfectivaLetras", "Cincuenta");
			json.put("tasaENum", "50");
			json.put("tasaReferencia", "Diez");
			json.put("tasaMora", "Diez");
			json.put("montoComEstru", "100");
			json.put("ivaComEstru", "13");
			json.put("lugarFormalizacion", "San Salvador");
			json.put("fDia", "cinco");
			json.put("fMes", "Abril");
			json.put("fAn", "Veintitres");

			String nameWord = (String) json.get("nombre");
			// System.out.println("tamaño del json "+json.size());

			// variables se van a llenar desde base de datos
			// crear captura de variables que no hagan match, bitacora
			variablesWord = new String[] { "nombre", "edad", "prof", "domicilio", "departamento", "dui",
					"montoAceptado", "cuentaParaCargo", "plazo", "cuotaTotal", "cuotaCapitalIntereses", "destino",
					"seguroVidaD", "dia", "tasaNominalEnLetras", "tasaNominalNumeros", "tasaEfectivaLetras", "tasaENum",
					"tasaReferencia", "tasaMora", "montoComEstru", "ivaComEstru", "lugarFormalizacion", "fDia", "fMes",
					"fAn" };

//	String[] valores={"Luis Lemus","33","programador","tasaEfectivaLetras","tasaEfectivaNumeros","tasaReferencia","tasaMora",
//			"montoComisionEstructuracion"};	
		}
		if (tipoReporte == 2) {
			json.put("fDia", "cinco");
			json.put("fMes", "Abril");
			json.put("fAn", "Veintitres");
			json.put("nombreDui", "Luis Mario2");
			json.put("montoAceptado", "2000");
			json.put("capitalDeTrabajo", "Destino de capital");
			json.put("plazo", "23");
			json.put("tasaNominalEnLetras", "Veinte");
			json.put("tasaNominalNumeros", "20");
			json.put("tasaEfectivaLetras", "Cincuenta");
			json.put("tasaENum", "50");
			json.put("tasaMora", "10");
			json.put("cuentaParaCargo", "0355588777444");
			json.put("plazoPago", "0355588777444");
			json.put("cuotaCapitalIntereses", "50");
			json.put("seguroVidaD", "200");
			json.put("cuotaTotal", "40");
			json.put("docPrivAcepDigital", "reemplazo");
			json.put("interesCancelarTeorico", "5");
			json.put("totalCancelarTeorico", "Veinte");
			json.put("montoComisionEstructu", "20");
			json.put("ivaComisionEstructu", "6");
					
			json.put("nombreClienteCuenta", "Luis Mario Lemus");
			
			json.put("totalComisionEstructuracion", "13");
			json.put("remanente", "San Salvador");

			String nameWord = (String) json.get("nombre");
			// System.out.println("tamaño del json "+json.size());

			// variables se van a llenar desde base de datos
			// crear captura de variables que no hagan match, bitacora
			variablesWord = new String[] { "fDia", "fMes", "fAn", "nombreDui", "montoAceptado", "capitalDeTrabajo",
					"plazo", "tasaNominalEnLetras","tasaNominalNumeros","tasaEfectivaLetras","tasaENum", "tasaMora", "cuentaParaCargo",
					"cuotaCapitalIntereses", "seguroVidaD", "cuotaTotal", "docPrivAcepDigital", "interesCancelarTeorico","totalCancelarTeorico",
					"montoComisionEstructu","plazoPago",
					 "ivaComisionEstructu", "nombreClienteCuenta", "montoAceptado", 
					"totalComisionEstructuracion", "remanente" };

//	String[] valores={"Luis Mario2","20","programador","tasaEfectivaLetras","tasaEfectivaNumeros","tasaReferencia","tasaMora",
//			"montoComisionEstructuracion"};	
		}

		// doc = replaceText(doc, "<<nombre>>", "luis mario lemus");

		for (int i = 0; i < json.size(); i++) {

			String replace = (String) json.get(variablesWord[i]);
			 System.out.println("vars:"+variablesWord[i]+" json "+replace);
			doc = replaceText(doc, "{" + variablesWord[i] + "}", replace);
		}
		for (int i = 0; i < variablesWord.length; i++) {
			String replace = (String) json.get(variablesWord[i]);
			// System.out.println("llegamos aqui");

		}
		// metodo que utiliza collection, pero se descarta por falta de costumbre
		Iterator it = textsToReplace.keySet().iterator();
		while (it.hasNext()) {
			String item = (String) it.next();
			// Enviamos a reemplazar el texto y guardamos sustituimos el documento en la
			// misma variable
			// doc = replaceText(doc, item, textsToReplace.get(item));
		}
		System.out.println("archivoTemporal " + temporalFileName);
		// Una vez se haya terminado de reemplazar, guardamos el documento en un archivo
		// temporal
		// Para ello general el archivo temporal que pasaremos a nuestro metodo que se
		// encarga de realizar la escritura
		String tmpDestinationPath = Files.createTempFile("temporalFileName", "." + Constantes.EXTENSION_DOCX)
				.toAbsolutePath().toString();
		System.out.println("tmpDestinationPath " + tmpDestinationPath);
		// Guardamos el documento en el archivo
		saveWord(tmpDestinationPath, doc);

		// Retornamos un ImputStream por si el usuario va a trabajar con el
		return new FileInputStream(tmpDestinationPath);
	}

	private InputStream replaceTextsAndConvertToPDF() throws Exception {
		try {
			System.out.println("llamamos el metodo");
			InputStream in = replaceTexts();
			DocxToPdfConverter cwoWord = new DocxToPdfConverter();
			System.out.println("que es in " + in);
			return cwoWord.convert(in);
		} catch (Exception e) {
			throw e;
		}
	}

	// En el se carga el proceso de sustituir las variables de word y convertir a
	// pdf
	public static void main(String[] args) {
		System.out.println(UUID.randomUUID().toString());
		System.out.println("1");

		// Archivo que indicamos la ruta a cargar para convertir a pdf
		Integer tipoReporte = 2;
		String filePath="";
		if (tipoReporte==1) {
			 filePath = "C:\\Users\\567\\Desktop\\importarPyme.docx";
			// String filePathDestino = "C\\Users\\567\\Desktop\\pruebaConvertir.pdf";	
		}
		if (tipoReporte==2) {
			 filePath = "C:\\Users\\567\\Desktop\\CARTAMOVILPYME.docx";
			// String filePathDestino = "C\\Users\\567\\Desktop\\pruebaConvertir.pdf";	
		}
		

		try {
			DocxHandling handling = new DocxHandling();
			// (Collections.singletonMap("{nombre}", "Rolando Mota Del Campo"), filePath,
			// filePathDestino);
			handling.setOriginalFilePath(filePath);
			System.out.println("iniciamos ");

			handling.setTextsToReplace(Collections.singletonMap("<<nombre>>", "Luis Lemus"));
			InputStream in = handling.replaceTextsAndConvertToPDF();
			// File destino = new File(filePathDestino);
			// System.out.println("file destino "+destino);
			// Files.copy(in, destino.toPath(), StandardCopyOption.REPLACE_EXISTING);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static XWPFDocument replaceText(XWPFDocument doc, String findText, String replaceText) {
		// Realizamos el recorrido de todos los parrafos
		//System.out.println("replaceText " + findText + "  " + replaceText + "  " + doc.getParagraphs().size());
		for (int i = 0; i < doc.getParagraphs().size(); ++i) {
			// Asignamos el parrafo a una variable para trabajar con ella
			XWPFParagraph s = doc.getParagraphs().get(i);

			// De ese parrafo recorremos todos los Runs
			for (int x = 0; x < s.getRuns().size(); x++) {
				// Asignamos el run en turno a una varibale
				XWPFRun run = s.getRuns().get(x);
				// Obtenemos el texto
				String text = run.text();
				// Validamos si el texto contiene el key a sustituir

				if (text.contains(findText)) {
					// System.out.println("replaceText "+findText + " "+replaceText);
					// Si lo contiene lo reemplazamos y guardamos en una variable
					String replaced = run.getText(run.getTextPosition()).replace(findText, replaceText);
					// Pasamos el texto nuevo al run
					run.setText(replaced, 0);
				}
			}
		}
		// Retornamos el documento con los textos ya reemplazados
		return doc;
	}

	private static void saveWord(String filePath, XWPFDocument doc) throws FileNotFoundException, IOException {
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filePath);
			doc.write(out);
		} finally {
			out.close();
		}
	}

	public void setTextsToReplace(Map<String, String> textsToReplace) {
		this.textsToReplace = textsToReplace;
	}

	public void setOriginalFilePath(String originalFilePath) {
		this.originalFilePath = originalFilePath;
	}

}

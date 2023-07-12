package dev.mamasoyjuanito.handling;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.UUID;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import dev.mamasoyjuanito.utils.Constantes;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

public class DocxToPdfConverter {

	public static void main(String[] args) {
		System.out.println("Starting conversion!!!");
		DocxToPdfConverter cwoWord = new DocxToPdfConverter();
		cwoWord.convert("C:\\Users\\567\\AppData\\Local\\Temp\\temporalFileName16737802201674949491.docx",
				"C\\Users\\567\\Desktop\\pruebaConvertir.pdf");
		System.out.println("Ending conversion!!!");
	}

	public void convert(String docPath, String pdfPath) {
		try {
			InputStream doc = new FileInputStream(new File(docPath));
			XWPFDocument document = new XWPFDocument(doc);
			PdfOptions options = PdfOptions.create();
			OutputStream out = new FileOutputStream(new File(pdfPath));
			PdfConverter.getInstance().convert(document, out, options);
		} catch (IOException ex) {
			System.out.println(ex.getMessage());
		}
	}

	public void convert(InputStream in, String pdfPath) {
		try {
			XWPFDocument document = new XWPFDocument(in);
			PdfOptions options = PdfOptions.create();
			OutputStream out = new FileOutputStream(new File(pdfPath));
			PdfConverter.getInstance().convert(document, out, options);
		} catch (IOException ex) {
			System.out.println(ex.getMessage());
		}
	}
	
	
	public InputStream convert(String docPath) {
		try {
			InputStream doc = new FileInputStream(new File(docPath));
			File tmpFile = Files.createTempFile(UUID.randomUUID().toString(), "." + Constantes.EXTENSION_PDF).toFile();
			XWPFDocument document = new XWPFDocument(doc);
			PdfOptions options = PdfOptions.create();
			
			OutputStream out = new FileOutputStream(tmpFile);
			PdfConverter.getInstance().convert(document, out, options);
			return new FileInputStream(tmpFile);
		} catch (IOException ex) {
			System.out.println(ex.getMessage());
		}
		return null;
	}

	public InputStream convert(InputStream in) {
		try {
			System.out.println("esto es convert ");
			File tmpFile = Files.createTempFile(UUID.randomUUID().toString(), "." + Constantes.EXTENSION_PDF).toFile();
			XWPFDocument document = new XWPFDocument(in);
			PdfOptions options = PdfOptions.create();
			System.out.println("esto es tmpFile "+tmpFile);
			OutputStream out = new FileOutputStream(tmpFile);
			PdfConverter.getInstance().convert(document, out, options);
			System.out.println("Aqui convertimose a pdf");
			return new FileInputStream(tmpFile);
		} catch (IOException ex) {
			System.out.println(ex.getMessage());
		}
		return null;
	}
}
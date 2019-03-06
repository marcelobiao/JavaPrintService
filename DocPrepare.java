package model;

import java.awt.image.BufferedImage;
import java.awt.print.Book;
import java.awt.print.PageFormat;
import java.awt.print.Paper;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.ConnectException;
import java.util.List;

import javax.print.PrintService;
import javax.print.attribute.standard.Chromaticity;
import javax.print.attribute.standard.Copies;

import org.apache.http.annotation.Experimental;
import org.apache.log4j.Logger;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.printing.PDFPageable;
import org.apache.pdfbox.printing.PDFPrintable;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.ghost4j.Ghostscript;
import org.ghost4j.GhostscriptException;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

import controller.Controller;
import javafx.print.PageRange;
import util.Log;

/**
 * Classe responsável por formatar o documento antes da impressão.
 * @author Marcelo Bião
 */
public class DocPrepare {
	private PDDocument doc;
	private String filePath;
	private PrintService printService;
	private Copies copies;
	private PageRange pageRange;
	private Chromaticity color;
	
	static Logger log = Logger.getLogger(DocPrepare.class.getName());
	/**
	 * Formata documento para ser impresso, definindo cor da impressão, quantidade de cópias, range de páginas do documento. 
	 * No momento aceita apenas PDF e tamanho A4.
	 * @param filePath Define path do arquivo que será impresso. Path absoluto.
	 * @param copies Define quantidade de cópias. Pode ser nulo. Default 1. 
	 * @param pageRange Define range de impressão. Pode ser nulo. Default [1-N].
	 * @param color Define cor da impressão. Pode ser nulo. Default Chromaticity.MONOCHROME
	 */
	public DocPrepare(String filePath, PrintService printService, Copies copies, PageRange pageRange, Chromaticity color) throws Exception {
		
		//inicializando variáveis
		this.filePath = filePath;
		this.copies = copies;
		this.pageRange = pageRange;
		this.color = color;		
		this.printService = printService;
		log.info(this.toString());
		
		//Carregando documento
		String[] pathArray = this.filePath.split("\\.");
		if((pathArray.length - 1) <= 0)
			throw new Exception("Extensão não encontrada");
		String extension = pathArray[pathArray.length - 1].toLowerCase();
		if(extension.equals("pdf"))
			this.pdfLoad();
		else if(extension.equals("docx"))
			this.docLoad();
		else
			throw new Exception("Extensão não suportada");
		Log.docSave(this.doc,"doc1Load.pdf");
		
		//Converte documento para escala de cinza		
		if(this.color instanceof Chromaticity) {
			if(this.color.equals(Chromaticity.MONOCHROME)) {
				//this.gsDocToGray();
				this.docToGray();
			}
		}else {
			//this.gsDocToGray();
			this.docToGray();
		}
		Log.docSave(this.doc,"doc2Gray.pdf");
				
		//Recorta documento de acordo com o range
		if(this.pageRange instanceof PageRange)
			this.docSplit(this.pageRange.getStartPage(), this.pageRange.getEndPage());
		Log.docSave(this.doc,"doc3Split.pdf");
		
		//Duplica documento de acordo com a quantidade de copias
		if(this.copies instanceof Copies)
			this.docCopies(this.copies.getValue());
		else
			this.docCopies(1);
		Log.docSave(this.doc,"doc4Copies.pdf");
	}
	
	public void decodificar() {
		//TODO: Implementar
	}
	
	public void pdfLoad() throws InvalidPasswordException, IOException {
		this.doc = PDDocument.load(new File(filePath));
	}
	
	@Experimental
	public void docLoadJod() throws IOException {
		OpenOfficeConnection connection = null;
	    try {
	      File inputFile = new File(this.filePath);
	      String workingDirectory = System.getenv("HOME");
	      String theFile = workingDirectory + "/PrinterSev/saida.pdf";
	      File outputFile = new File(theFile);

	      connection = new SocketOpenOfficeConnection(8100);
	      connection.connect();

	      // convert
	      DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
	      converter.convert(inputFile, outputFile);
	      this.doc = PDDocument.load(new File(theFile));
	    }
	    finally {
	      // close the connection
	      if (connection.isConnected()) {
	        connection.disconnect();
	      }
	    }
	}
	
	@Experimental
	public void docLoad() throws Docx4JException, InvalidPasswordException, IOException {
		//TODO: Corrigir, carrega docx porém gera uma marca na impressão.
		InputStream is = new FileInputStream(new File(filePath));
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(is);
        List sections = wordMLPackage.getDocumentModel().getSections();   
        
        String workingDirectory = System.getenv("HOME");
        String theFile = workingDirectory + "/PrinterSev/saida.pdf";        
        Docx4J.toPDF(wordMLPackage, new FileOutputStream(theFile));
        
        this.doc = PDDocument.load(new File(theFile));
	}
	
	public void gsDocToGray() throws IOException, GhostscriptException {
		String workingDirectory = System.getenv("HOME");
        String theFile = workingDirectory + "/PrinterSev/saida.pdf";    
        
		Ghostscript gs = Ghostscript.getInstance();
		 
		String[] gsArgs = new String[9];
		gsArgs[1] = "-sDEVICE=pdfwrite";
		gsArgs[2] = "-sProcessColorModel=DeviceGray";
		gsArgs[3] = "-sColorConversionStrategy=Gray";
		gsArgs[4] = "-dOverrideICC";
		gsArgs[5] = "-o";
		gsArgs[6] = theFile;
		gsArgs[7] = "-f";
		gsArgs[8] = this.filePath;
 
        //execute and exit interpreter
        gs.initialize(gsArgs);
        gs.exit();
        this.doc = PDDocument.load(new File(theFile));
	}
	@Deprecated
	public void docToGray() throws IOException {
		PDFRenderer pdfRenderer = new PDFRenderer(this.doc);
		PDDocument docOut = new PDDocument();
		
		for (int pageNum = 0; pageNum < this.doc.getNumberOfPages(); ++pageNum)
		{ 
		    BufferedImage bim = pdfRenderer.renderImageWithDPI(pageNum, 300, ImageType.GRAY);
		    PDPage newPage = new PDPage(this.doc.getPage(0).getMediaBox());
		    docOut.addPage(newPage);
		    
		    PDImageXObject pdImage = JPEGFactory.createFromImage(docOut, bim);
		    PDPageContentStream content = new PDPageContentStream(docOut, newPage);
		    content.drawImage(pdImage, 0, 0,newPage.getMediaBox().getWidth(),newPage.getMediaBox().getHeight());
		    content.close();
		}
		this.doc = docOut;
	}
	
	public void docSplit(int lowerBound, int upperBound) throws IOException {
		PDDocument docOut = new PDDocument();
		for(int i = lowerBound; i <= upperBound; i++) {
			try{
				docOut.addPage(this.doc.getPage(i-1));
			}catch(IndexOutOfBoundsException e) {
				System.out.println("Página "+ i + " não existe");
				break;
			}
		}
		this.doc = docOut;
	}
	
	public void docCopies(int copies) throws Exception {
		if(copies > 100) {
			throw new Exception("Quantidade de cópias acima do permitido, máximo de 100 copias");
		}
		
		PDDocument docOut = new PDDocument();
		PDFMergerUtility PDFmerger = new PDFMergerUtility();
		for(int i =1;i<=copies; i++) {
			PDFmerger.appendDocument(docOut, this.doc);
		}
		this.doc = docOut;
	}
	
	/**
	 * Redimensiona o documento seguindo o padrão ISO.AN.
	 */
	public void docResize() {
		//TODO: Implementar
	}
	/**
	 * Novo método para impressão.
	 * @throws Exception 
	 */
	public void docPrintGS() throws Exception {
		//Verifica se doc possui páginas
	    if(this.doc.getNumberOfPages() == 0) {
	    	throw new Exception("Documento vazio");
	    }
	    
	    //Verifica se existe impressora
	    if(!(this.printService instanceof PrintService)) {
	    	throw new Exception("Impressora não encontrada");
	    }
	    
		PrinterJob job = PrinterJob.getPrinterJob();
	    //PDDocument doc = PDDocument.load(new File(this.doc));
	    job.setPageable(new PDFPageable(this.doc));
	    job.setPrintService(this.printService);
	    job.print();
	}
	
	@Deprecated
	public void docPrint() throws Exception {
		//Verifica se doc possui páginas
	    if(this.doc.getNumberOfPages() == 0) {
	    	throw new Exception("Documento vazio");
	    }
	    
	    if(!(this.printService instanceof PrintService)) {
	    	throw new Exception("Impressora não encontrada");
	    }
	    
	    //Configurando a folha de impressão
	    PDPage page = this.doc.getPage(0);
	    
	    Paper paper = new Paper();
	    if(page.getMediaBox().getWidth() > page.getMediaBox().getHeight())
	    	paper.setSize(PDRectangle.A4.getHeight(),PDRectangle.A4.getWidth());
	    else
	    	paper.setSize(PDRectangle.A4.getWidth(),PDRectangle.A4.getHeight());
	    
	    paper.setImageableArea(0, 0, paper.getWidth(), paper.getHeight());
	    PageFormat pageFormat = new PageFormat();
	    pageFormat.setPaper(paper);
	    
	    //Criando Book de impressão
	    //TODO: Configurar folha a folha, para impressões documentos com impressão retrato+paisagem
	    Book book = new Book();
	    book.append(new PDFPrintable(this.doc), pageFormat, this.doc.getNumberOfPages());
	      
	    //Gerando serviço de impressão
	    //TODO: RECEBER Dispositivo
	    PrinterJob job = PrinterJob.getPrinterJob();
	    job.setPrintService(this.printService);
	    job.setPageable(book);
		job.print();
	}
	
	@Override
	public String toString() {
		return ("DocPrepareConfig = {File Path: "+this.filePath+", \n"+
				"Copies: "+this.copies+", \n"+
				"Page Range: "+this.pageRange+", \n"+
				"Color: "+this.color+", \n"+
				"PrintService: "+this.printService+"}"
				);
	}
}
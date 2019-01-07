package model;

import java.awt.image.BufferedImage;
import java.awt.print.Book;
import java.awt.print.PageFormat;
import java.awt.print.Paper;
import java.awt.print.PrinterJob;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.print.PrintService;
import javax.print.attribute.standard.Chromaticity;
import javax.print.attribute.standard.Copies;

import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.printing.PDFPrintable;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import javafx.print.PageRange;

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
	
	/**
	 * Formata documento para ser impresso, definindo cor da impressão, quantidade de cópias, range de páginas do documento. 
	 * No momento aceita apenas PDF e tamanho A4.
	 * @param filePath Define path do arquivo que será impresso. Path absoluto.
	 * @param copies Define quantidade de cópias. Pode ser nulo. Default 1. 
	 * @param pageRange Define range de impressão. Pode ser nulo. Default [1-N].
	 * @param color Define cor da impressão. Pode ser nulo. Default Chromaticity.MONOCHROME
	 */
	public DocPrepare(String filePath, PrintService printService, Copies copies, PageRange pageRange, Chromaticity color) throws Exception {
		
		this.filePath = filePath;
		this.copies = copies;
		this.pageRange = pageRange;
		this.color = color;		
		this.printService = printService;
		
		//Carregando documento
		String[] pathArray = this.filePath.split("\\.");
		if((pathArray.length - 1) <= 0)
			throw new Exception("Extensão não encontrada");
		String extension = pathArray[pathArray.length - 1];
		if(extension.equals("pdf"))
			this.pdfLoad();
		else if(extension.equals("docx"))
			this.docLoad();
		else
			throw new Exception("Extensão não suportada");
		
		//Converte documento para escala de cinza
		if(this.color instanceof Chromaticity) {
			if(this.color.equals(Chromaticity.MONOCHROME))
				this.docToGray();
		}else
			this.docToGray();
			
		//Recorta documento de acordo com o range
		if(this.pageRange instanceof PageRange)
			this.docSplit(this.pageRange.getStartPage(), this.pageRange.getEndPage());
		
		//Duplica documento de acordo com a quantidade de copias
		if(this.copies instanceof Copies)
			this.docCopies(this.copies.getValue());
		else
			this.docCopies(1);
	}
	
	public void decodificar() {
		//TODO: Implementar
	}
	
	public void pdfLoad() throws InvalidPasswordException, IOException {
		this.doc = PDDocument.load(new File(filePath));
	}
	
	public void docLoad() throws Docx4JException, InvalidPasswordException, IOException {
		//TODO: Corrigir, não carrega .doc nem .docx
		InputStream is = new FileInputStream(this.filePath);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(is);
        //List sections = wordMLPackage.getDocumentModel().getSections();
        
        ByteArrayOutputStream outStream =  new ByteArrayOutputStream();
		Docx4J.toPDF(wordMLPackage, outStream);
		ByteArrayInputStream inStream = new ByteArrayInputStream(outStream.toByteArray());
		PDDocument outDoc = PDDocument.load(inStream);
		
		this.doc = outDoc;
	}
	
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
}
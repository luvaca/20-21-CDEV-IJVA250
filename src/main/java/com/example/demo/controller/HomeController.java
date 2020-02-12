package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.service.impl.ArticleServiceImpl;
import com.example.demo.service.impl.ClientServiceImpl;
import com.ibm.icu.impl.DateNumberFormat;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.bcel.Const;
import org.eclipse.datatools.modelbase.sql.query.ValueExpressionVariable;
import org.springframework.boot.autoconfigure.data.redis.RedisProperties.Lettuce;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.xml.ws.Response;

import java.util.Date;

/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController {

    private ArticleService articleService;
    private ClientServiceImpl clientServiceImpl;
    private FactureService factureService;

    public HomeController(ArticleService articleService, ClientServiceImpl clientService, FactureService factureService) {
        this.articleService = articleService;
        this.clientServiceImpl = clientService;
        this.factureService = factureService;
    }

    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Article> articles = articleService.findAll();
        modelAndView.addObject("articles", articles);

        List<Client> toto = clientServiceImpl.findAllClients();
        modelAndView.addObject("clients", toto);

        List<Facture> factures = factureService.findAllFactures();
        modelAndView.addObject("factures", factures);

        return modelAndView;
    }
    @GetMapping("/articles/csv")
    public void articlesCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
    		response.setContentType("text/csv");
    		response.setHeader("Content-Disposition","attachement; filename=\"export-articles.csv\"");
    		PrintWriter writer = response.getWriter();
    		
    		 List<Article> articles = articleService.findAll();
    		 writer.println("prix ; libelle");
    		 for (int i = 0; i < articles.size(); i++) {
    		
    			 Article article = articles.get(i);
    			
    			 writer.println(article.getPrix()+ ";" + article.getLibelle());
    			 
    		 	}
    		 }
    		 
    		   @GetMapping("/clients/csv")
    		    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
    		    		response.setContentType("text/csv");
    		    		response.setHeader("Content-Disposition","attachement; filename=\"export-clients.csv\"");
    		    		PrintWriter writer = response.getWriter();
    		    		
    		    		List<Client> clients = clientServiceImpl.findAllClients();
    		    		 writer.println("nom ; prenom ; age");
    		    	
    		    		 for (int i = 0; i < clients.size(); i++) {
    		    		
    		    			 Client client = clients.get(i);
    		    			 
    		    			 LocalDate now= LocalDate.now();   
        		    		 Integer age = client.getDateNaissance().until(now).getYears();
    		    			
    		    			 writer.println(client.getNom()+ ";" + client.getPrenom() + ";" + age);
    		    			 
    					}
    		
    }
    		   
    	
    		   
    		   // import xlsx client
    		   @GetMapping("/clients/xlsx")
    		   public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
  		    		response.setContentType("application/vnd.ms-excel");
  		    		response.setHeader("Content-Disposition","attachement; filename=\"export-clients.xlsx\"");
  		    		//ServletOutputStream os = response.getOutputStream();
	    		   
	    		   Workbook workbook = new XSSFWorkbook();
	    		   Sheet sheet = workbook.createSheet("Clients");
	    		   Row headerRow = sheet.createRow(0);
	    		   Cell cellPrenom = headerRow.createCell(0);
	    		   Cell cellName = headerRow.createCell(1);
	    		   Cell cellAge = headerRow.createCell(2);
	    		   cellPrenom.setCellValue("Prénom");
	    		   cellName.setCellValue("Nom");
	    		   cellAge.setCellValue("age");
	    		   
	    		   Integer rowNum = 1;
	    		   List<Client> clients = clientServiceImpl.findAllClients();
	    		   
	    		   
	    		   for(Client client : clients) {
	    			   
	    			 Row row = sheet.createRow(rowNum++);
	    			   
	    			 LocalDate now = LocalDate.now();   
		    		 Integer age = client.getDateNaissance().until(now).getYears();
		    		 
		    		 row.createCell(0)
		    		 	.setCellValue(client.getPrenom());
		    		 
		    		 row.createCell(1)
		    		 	.setCellValue(client.getNom());
		    		 
		    		 row.createCell(2)
		    		 	.setCellValue(age);
	    		   }
	    		   
	    		   //FileOutputStream fileOut = new FileOutputStream("export-clients.xlsx");
					workbook.write(response.getOutputStream());
				    
	    		   //workbook.write(os);
	    		   //fileOut.close();
	    		   workbook.close();
    		   
    		   }
    		   
    		   
    		   // import xlsx articles
    		   @GetMapping("/articles/xlsx")
    		   public void articlesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
  		    		response.setContentType("application/vnd.ms-excel");
  		    		response.setHeader("Content-Disposition","attachement; filename=\"export-articles.xlsx\"");
  		    		//ServletOutputStream os = response.getOutputStream();
	    		   
	    		   Workbook workbook = new XSSFWorkbook();
	    		   Sheet sheet = workbook.createSheet("Articles");
	    		   Row headerRow = sheet.createRow(0);
	    		   Cell cellLibelle = headerRow.createCell(0);
	    		   Cell cellPrix = headerRow.createCell(1);
	    		   cellLibelle.setCellValue("Libelle");
	    		   cellPrix.setCellValue("Prix");
	    		  
	    		   
	    		   Integer rowNum = 1;
	    		   List<Article> articles = articleService.findAll();
	    		   
	    		   
	    		   for(Article article : articles) {
	    			   
	    			 Row row = sheet.createRow(rowNum++);
	    			   
		    		 
		    		 row.createCell(0)
		    		 	.setCellValue(article.getLibelle());
		    		 
		    		 row.createCell(1)
		    		 	.setCellValue(article.getPrix());
		    		 
	    		   }
	    		   
	    		   //FileOutputStream fileOut = new FileOutputStream("export-clients.xlsx");
					workbook.write(response.getOutputStream());
				    
	    		   //workbook.write(os);
	    		   //fileOut.close();
	    		   workbook.close();
    		   
    		   }
    		   
    		   
    		   // import xlsx articles
    		   @GetMapping("/factures/xlsx")
    		   public void factureXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
  		    		response.setContentType("application/vnd.ms-excel");
  		    		response.setHeader("Content-Disposition","attachement; filename=\"export-factures.xlsx\"");
  		    		//ServletOutputStream os = response.getOutputStream();
	    		   
	    		   Workbook workbook = new XSSFWorkbook();
	    		   Sheet sheet = workbook.createSheet("facture");
	    		   Row headerRow = sheet.createRow(0);
	    		   Cell cellId = headerRow.createCell(0);
	    		   Cell cellClient = headerRow.createCell(1);
	    		   Cell cellTotal = headerRow.createCell(2);
	    		   Cell cellLigne = headerRow.createCell(3);
	    		   cellId.setCellValue("Id");
	    		   cellClient.setCellValue("Client");
	    		   cellTotal.setCellValue("Total");
	    		   cellLigne.setCellValue("Ligne");
	    		  
	    		   
	    		   Integer rowNum = 1;
	    		   List<Facture> factures = factureService.findAllFactures();
	    		   
	    		   
	    		   for(Facture facture : factures) {
	    			   
	    			 Row row = sheet.createRow(rowNum++);
	    			   
		    		 
		    		 row.createCell(0)
		    		 	.setCellValue(facture.getId());
		    		 
		    		 row.createCell(1)
		    		 	.setCellValue(facture.getClient().getNom());
		    		 
		    		 row.createCell(2)
		    		 	.setCellValue(facture.getTotal());
		    		 
		    		// row.createCell(3)
		    		 //	.setCellValue(facture.getLigneFactures());
		    		 
	    		   }
	    		   
	    		   //FileOutputStream fileOut = new FileOutputStream("export-clients.xlsx");
					workbook.write(response.getOutputStream());
				    
	    		   //workbook.write(os);
	    		   //fileOut.close();
	    		   workbook.close();
    		   
    		   }
    		   
    		   
    		   
 
}



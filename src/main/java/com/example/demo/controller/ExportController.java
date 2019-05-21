package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.ClientRepository;
import com.example.demo.repository.FactureRepository;
import com.example.demo.service.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private  FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id" + ";" + "Nom" + ";" + "Prenom" + ";" + "Date de Naissance" + ";" + "Age");
        for (Client liste : allClients) {
            writer.println(liste.getId() + ";"
            + liste.getNom() + ";"
            + liste.getPrenom() + ";"
            + liste.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/mm/yyyy")) + ";"
            + (now.getYear() - liste.getDateNaissance().getYear()));
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("Prénom");
        headerRow.createCell(2).setCellValue("Nom");
        headerRow.createCell(3).setCellValue("Date de naissance");
        headerRow.createCell(4).setCellValue("Age");

        List<Client> allClients = clientService.findAllClients();

        for (Client client: allClients
             ) {
            Row row = sheet.createRow(sheet.getLastRowNum()+1);
            row.createCell(0).setCellValue(client.getId());
            row.createCell(1).setCellValue(client.getPrenom());
            row.createCell(2).setCellValue(client.getNom());
            row.createCell(3).setCellValue(client.getDateNaissance().toString());
            row.createCell(4).setCellValue(now.getYear() - client.getDateNaissance().getYear() );


        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        Workbook workbook = new XSSFWorkbook();

        List<Client> allClients = clientService.findAllClients();

        //Pour chaque client, création d'une feuille client
        for (Client client: allClients) {
            Sheet clientSheet = workbook.createSheet(client.getNom());
            Row clientRow = clientSheet.createRow(0);
            clientRow.createCell(0).setCellValue(client.getNom());
            clientRow.createCell(1).setCellValue(client.getPrenom());

            List<Facture> facturesClient = factureService.findByClientId(client.getId());

            //Pour chaque facture, création d'une feuille facture
            for (Facture facture: facturesClient) {
                Set<LigneFacture> ligneFactures = facture.getLigneFactures();
                Sheet factureSheet = workbook.createSheet("Facture " + facture.getId());
                factureSheet.setColumnWidth(0, 5000);
                factureSheet.setColumnWidth(1, 5000);
                factureSheet.setColumnWidth(2, 5000);
                factureSheet.setColumnWidth(3, 5000);
                Row headerRow = factureSheet.createRow(0);
                headerRow.createCell(0).setCellValue("Article");
                headerRow.createCell(1).setCellValue("Quantité");
                headerRow.createCell(2).setCellValue("Prix unitaire");
                headerRow.createCell(3).setCellValue("Prix de la ligne");

                //Pour chaque ligne de facture, création d'une ligne
                for (LigneFacture ligne: ligneFactures) {
                    Row ligneRow = factureSheet.createRow(factureSheet.getLastRowNum()+1);
                    ligneRow.createCell(0).setCellValue(ligne.getArticle().getLibelle());
                    ligneRow.createCell(1).setCellValue(ligne.getQuantite());
                    ligneRow.createCell(2).setCellValue(ligne.getArticle().getPrix());
                    ligneRow.createCell(3).setCellValue(ligne.getArticle().getPrix() * ligne.getQuantite());
                }

            //Ligne de total général de la facture, style gras rouge avec bordure
            Row totalRow = factureSheet.createRow(factureSheet.getLastRowNum()+1);
            CellStyle totalRowStyle = workbook.createCellStyle();
            totalRowStyle.setBorderBottom(BorderStyle.MEDIUM);
            totalRowStyle.setBorderRight(BorderStyle.MEDIUM);
            totalRowStyle.setBorderLeft(BorderStyle.MEDIUM);
            totalRowStyle.setBorderTop(BorderStyle.MEDIUM);
            totalRowStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
            totalRowStyle.setTopBorderColor(IndexedColors.RED.getIndex());
            totalRowStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
            totalRowStyle.setRightBorderColor(IndexedColors.RED.getIndex());

            Font font = workbook.createFont();
            font.setBold(true);
            font.setColor(IndexedColors.RED.getIndex());
            totalRowStyle.setFont(font);

            CellRangeAddress cellRangeAddress = new CellRangeAddress(totalRow.getRowNum(), totalRow.getRowNum(), 0,2 );
            factureSheet.addMergedRegion(cellRangeAddress);

            Cell cell1 = totalRow.createCell(0);
            cell1.setCellValue("TOTAL GENERAL");
            cell1.setCellStyle(totalRowStyle);

            Cell cell2 = totalRow.createCell(3);
            cell2.setCellValue(facture.getTotal());
            cell2.setCellStyle(totalRowStyle);
            }
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{id}/factures/xlsx")
    public void facturesClientsXLSX(@PathVariable("id") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures-client-" + clientId + ".xlsx\"");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Factures de Monsieur");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("N° facture");
        headerRow.createCell(1).setCellValue("Montant total");

        List<Facture> allFactures = factureService.findByClientId(clientId);

        generateFactureWorkbook(response, workbook, sheet, allFactures);
    }

    private void generateFactureWorkbook(HttpServletResponse response, Workbook workbook, Sheet sheet, List<Facture> allFactures) throws IOException {
        for (Facture facture: allFactures) {
            Row row = sheet.createRow(sheet.getLastRowNum()+1);
            row.createCell(0).setCellValue(facture.getId());
            row.createCell(1).setCellValue(facture.getTotal());
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }

}

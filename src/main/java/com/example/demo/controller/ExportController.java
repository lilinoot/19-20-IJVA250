package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.ClientRepository;
import com.example.demo.service.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.loader.plan.spi.LoadPlan;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ContentDisposition;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.swing.plaf.synth.Region;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.Period;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Set;

/**
 * Controller pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");

        for (Client client : allClients) {
            writer.println(client.getId() + ";"
                    + "\"" + client.getNom() + "\";"
                    + "\"" + client.getPrenom() + "\";"
                    + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")) + ";"
                    + Period.between(client.getDateNaissance(), now).getYears());
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vmd-s-excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.PLUM.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        headerRow.createCell(0).setCellValue("Id");
        headerRow.getCell(0).setCellStyle(headerCellStyle);
        headerRow.createCell(1).setCellValue("Nom");
        headerRow.getCell(1).setCellStyle(headerCellStyle);
        headerRow.createCell(2).setCellValue("Prénom");
        headerRow.getCell(2).setCellStyle(headerCellStyle);
        headerRow.createCell(3).setCellValue("Date de naissance");
        headerRow.getCell(3).setCellStyle(headerCellStyle);
        headerRow.createCell(4).setCellValue("Âge");
        headerRow.getCell(4).setCellStyle(headerCellStyle);

        int num = 1;

        for (Client client : allClients) {
            Row rowClient = sheet.createRow(num);
            rowClient.createCell(0).setCellValue(client.getId());
            rowClient.createCell(1).setCellValue(client.getNom());
            rowClient.createCell(2).setCellValue(client.getPrenom());
            rowClient.createCell(3).setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));
            rowClient.createCell(4).setCellValue(Period.between(client.getDateNaissance(), now).getYears());
            num++;
        }

        workbook.write(response.getOutputStream()); // Pour écrire les données
        workbook.close();

    }

    @GetMapping("/clients/{clientId}/factures/xlsx")
    public void facturesEachClientXLSX(@PathVariable("clientId") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vmd-s-excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"facture_client" + clientId + ".xlsx\"");
        Client client = clientService.findById(clientId);
        Set<Facture> factures = client.getFactures();
        Workbook workbook = new XSSFWorkbook();


        for (Facture facture : factures) {
            Sheet sheet = workbook.createSheet("Facture" + facture.getId());
            Row headerRow = sheet.createRow(0);

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 11);
            headerFont.setColor(IndexedColors.PLUM.getIndex());

            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            headerRow.createCell(0).setCellValue("Id");
            headerRow.getCell(0).setCellStyle(headerCellStyle);
            headerRow.createCell(1).setCellValue("Total");
            headerRow.getCell(1).setCellStyle(headerCellStyle);

            int num = 1;

            for (LigneFacture ligneFacture : facture.getLigneFactures()) {
                Row rowFacture = sheet.createRow(num);
                rowFacture.createCell(0).setCellValue(ligneFacture.getArticle().getId());
                rowFacture.createCell(1).setCellValue(ligneFacture.getSousTotal());
                num++;
            }
        }
        workbook.write(response.getOutputStream()); // Pour écrire les données
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vmd-s-excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.PLUM.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);


        for (Client client : allClients) {
            Sheet sheetClient = workbook.createSheet(client.getNom().toUpperCase());
            Row row = sheetClient.createRow(0);


            row.createCell(0).setCellValue("Nom");
            row.getCell(0).setCellStyle(headerCellStyle);
            row.createCell(1).setCellValue(client.getNom());

            row = sheetClient.createRow(row.getRowNum() + 1);
            row.createCell(0).setCellValue("Prénom");
            row.getCell(0).setCellStyle(headerCellStyle);
            row.createCell(1).setCellValue(client.getPrenom());

            row = sheetClient.createRow(row.getRowNum() + 1);
            row.createCell(0).setCellValue("Date de naissance");
            row.getCell(0).setCellStyle(headerCellStyle);
            row.createCell(1).setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));

            row = sheetClient.createRow(row.getRowNum() + 1);
            row.createCell(0).setCellValue("Âge");
            row.getCell(0).setCellStyle(headerCellStyle);
            row.createCell(1).setCellValue((Period.between(client.getDateNaissance(), now).getYears()));

            sheetClient.autoSizeColumn(0);
            sheetClient.autoSizeColumn(1);


            for (Facture facture : client.getFactures()) {
                Sheet sheetFacture = workbook.createSheet("Facture" + facture.getId());
                Row headerRow = sheetFacture.createRow(0);

                headerRow.createCell(0).setCellValue("Nom de l'article");
                headerRow.getCell(0).setCellStyle(headerCellStyle);
                headerRow.createCell(1).setCellValue("Quantité");
                headerRow.getCell(1).setCellStyle(headerCellStyle);
                headerRow.createCell(2).setCellValue("Prix unitaire");
                headerRow.getCell(2).setCellStyle(headerCellStyle);
                headerRow.createCell(3).setCellValue("Sous total");
                headerRow.getCell(3).setCellStyle(headerCellStyle);

                sheetFacture.autoSizeColumn(0);
                sheetFacture.autoSizeColumn(1);
                sheetFacture.autoSizeColumn(2);
                sheetFacture.autoSizeColumn(3);

                int num = 1;

                for (LigneFacture ligneFacture : facture.getLigneFactures()) {
                    Row rowFacture = sheetFacture.createRow(num);
                    rowFacture.createCell(0).setCellValue(ligneFacture.getArticle().getLibelle());
                    rowFacture.createCell(1).setCellValue(ligneFacture.getQuantite());
                    rowFacture.createCell(2).setCellValue(ligneFacture.getArticle().getPrix());
                    rowFacture.createCell(3).setCellValue(ligneFacture.getSousTotal());
                    num++;
                }

                CellStyle totalCellStyle = workbook.createCellStyle();
                totalCellStyle.setAlignment(HorizontalAlignment.RIGHT);

                Font totalFont = workbook.createFont();
                totalFont.setBold(true);
                totalFont.setColor(IndexedColors.PLUM.getIndex());
                totalCellStyle.setFont(totalFont);

                totalCellStyle.setBorderTop(BorderStyle.MEDIUM);
                totalCellStyle.setBorderBottom(BorderStyle.MEDIUM);
                totalCellStyle.setBorderLeft(BorderStyle.MEDIUM);
                totalCellStyle.setBorderRight(BorderStyle.MEDIUM);


                Row rowTotal = sheetFacture.createRow(num);
                CellRangeAddress cellRangeAddress = new CellRangeAddress(num, num , 0, 2);
                sheetFacture.addMergedRegion(cellRangeAddress);

                RegionUtil.setBorderTop(BorderStyle.MEDIUM, cellRangeAddress, sheetFacture);
                RegionUtil.setBorderBottom(BorderStyle.MEDIUM, cellRangeAddress, sheetFacture);
                RegionUtil.setBorderLeft(BorderStyle.MEDIUM, cellRangeAddress, sheetFacture);
                RegionUtil.setBorderRight(BorderStyle.MEDIUM, cellRangeAddress, sheetFacture);

                rowTotal.createCell(0).setCellValue("TOTAL");
                rowTotal.getCell(0).setCellStyle(totalCellStyle);

                rowTotal.createCell(3).setCellValue(facture.getTotal());
                rowTotal.getCell(3).setCellStyle(totalCellStyle);

            }


        }

        workbook.write(response.getOutputStream()); // Pour écrire les données
        workbook.close();


    }
}

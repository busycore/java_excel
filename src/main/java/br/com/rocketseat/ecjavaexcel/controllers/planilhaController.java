package br.com.rocketseat.ecjavaexcel.controllers;


import br.com.rocketseat.ecjavaexcel.models.User;
import br.com.rocketseat.ecjavaexcel.repositories.UsersRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;

import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

@RestController
@RequestMapping("/api/users")
public class planilhaController {

    @Autowired
    UsersRepository usersRepository;


    @PostMapping("/")
    public List<User> upload(@RequestParam("file")MultipartFile file) throws IOException {

        Workbook planilha = new XSSFWorkbookFactory().create(file.getInputStream());

        Sheet sheet = planilha.getSheetAt(0);
        //Sheet sheet = planilha.getSheet("Nome");


        List<User> listaUsuariosNoExcel = new ArrayList<>();

        for(Row row : sheet){
            //Skipar o Header

            if(row.getRowNum()==0){
                continue;
            }

            //Coletar as informações da linha
            String nome = row.getCell(0).getStringCellValue();
            Integer idade = (int)row.getCell(1).getNumericCellValue();
            LocalDate data_nascimento = row.getCell(2).getLocalDateTimeCellValue().toLocalDate();
            BigDecimal saldo = BigDecimal.valueOf(row.getCell(3).getNumericCellValue());

            //Fazer alguma trativa
            nome = nome.toUpperCase();

            //Criar um objeto de cliente
            User novoUser = new User();

            novoUser.setNome(nome);
            novoUser.setIdade(idade);
            novoUser.setData_nascimento(data_nascimento);
            novoUser.setSaldo(saldo);

            listaUsuariosNoExcel.add(novoUser);
        }
        //Persistir

        return usersRepository.saveAll(listaUsuariosNoExcel);

    }

    @ResponseBody
    @GetMapping("/")
    public ResponseEntity<ByteArrayResource> export() throws IOException {
        List<User> usersList =  usersRepository.findAll();

        Workbook nossaPlanilha = new XSSFWorkbook();
        String nomeSeguro = WorkbookUtil.createSafeSheetName("[~Relatório~]");
        Sheet nossaAba = nossaPlanilha.createSheet(nomeSeguro);

        //Criar header
        CellStyle headerStyle =  nossaPlanilha.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());




        var headerFields = Arrays.asList("Nome","Idade","Data de Nascimento","Saldo","Taxa",
                "Saldo-Taxa Formula","Saldo-Taxa no Java");


        Row headerRow = nossaAba.createRow(0);


        headerFields.forEach(eachField->{

            Cell headerCell =  headerRow.createCell(headerFields.indexOf(eachField), CellType.STRING);
            headerCell.setCellValue(eachField);
            headerCell.setCellStyle(headerStyle);

        });


        usersList.forEach(eachUser->{
            var currentUser = eachUser;
            int i =  usersList.indexOf(eachUser)+1;
            Row row = nossaAba.createRow(i);

            Cell colunaNome = row.createCell(0);
            colunaNome.setCellValue(currentUser.getNome());

            Cell colunaIdade = row.createCell(1);
            colunaIdade.setCellValue(currentUser.getIdade());

            Cell colunaNascimento = row.createCell(2);
            colunaNascimento.setCellValue(currentUser.getData_nascimento());

            Cell colunaSaldo = row.createCell(3);
            colunaSaldo.setCellValue(currentUser.getSaldo().doubleValue());
            String colunaSaldoLetra = CellReference.convertNumToColString(3);

            Cell colunaTaxa = row.createCell(4);
            colunaTaxa.setCellValue(250.00);
            String colunaTaxaLetra = CellReference.convertNumToColString(4);

            Cell colunaValorDeduzido = row.createCell(5);
            int linhaAtual = i+1;

            colunaValorDeduzido.setCellFormula("(D"+(i+1)+"-"+"E"+(i+1)+")");
            //colunaValorDeduzido.setCellFormula("("+colunaSaldoLetra+linhaAtual+"-"+colunaTaxaLetra+linhaAtual+")");

            Cell colunaValorDeduzidoJava = row.createCell(6);
            colunaValorDeduzidoJava.setCellValue(colunaSaldo.getNumericCellValue()-colunaTaxa.getNumericCellValue());
        });


        //Totalizador
        Row linhaTotalizadora =  nossaAba.createRow(usersList.size()+1);
        String formulaSomatoria = "SUM("+CellReference.convertNumToColString(5)+2+":"+CellReference.convertNumToColString(5)+usersList.size()+")";
        linhaTotalizadora.createCell(5).setCellFormula(formulaSomatoria);


        HttpHeaders responseHeaders = new HttpHeaders();
        responseHeaders.setContentType(new MediaType("application", "force-download"));
        responseHeaders.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ProductTemplate.xlsx");


        ByteArrayOutputStream stream = new ByteArrayOutputStream();

//        try( OutputStream fileOutput = new FileOutputStream("workbook.xlsx")) {
//            nossaPlanilha.write(fileOutput);
//            fileOutput.close();
//        }
//
//
//        final InputStream fileInputStream = new FileInputStream("workbook.xlsx");
//
//        return fileInputStream.readAllBytes();


          nossaPlanilha.write(stream);
          nossaPlanilha.close();

        return new ResponseEntity<ByteArrayResource>(new ByteArrayResource(stream.toByteArray()),
                responseHeaders, HttpStatus.OK);

    }


}

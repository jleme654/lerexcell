package br.com.missaocena;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import br.com.vo.Assistido;


public class LePlanilha {
	
	private static final String fileName = "/home/julio/arquivos/Lista-pop-rua-2018.xls";

	public static void main(String[] args) throws IOException {

		List<Assistido> listaAlunos = new ArrayList<Assistido>();

		try {
			FileInputStream arquivo = new FileInputStream(new File(LePlanilha.fileName));

			HSSFWorkbook workbook = new HSSFWorkbook(arquivo);

			HSSFSheet sheetAlunos = workbook.getSheetAt(0);

			Iterator<Row> rowIterator = sheetAlunos.iterator();
			
			int count = 0;
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				Assistido aluno = new Assistido();
				listaAlunos.add(aluno);
				Timestamp dataDeHoje = new Timestamp(System.currentTimeMillis());
				while (cellIterator.hasNext()) {
					count++;
					Cell cell = cellIterator.next();
					switch (cell.getColumnIndex()) {
					case 0:
						aluno.setId(count++);
						break;
					case 1:
						aluno.setNome(cell.getStringCellValue());
						break;
					case 2:
						aluno.setDataVisita(dataDeHoje);
						break;
					case 3:
						aluno.setRg(cell.getStringCellValue());
						break;
					case 4:
						aluno.setCpf(cell.getStringCellValue());
						break;
					}
				}
			}
			arquivo.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo Excel não encontrado!");
		}

		for (Assistido assistido : listaAlunos) {
			System.out.println(assistido);
		}
	}
}
package br.com.amazon;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import br.com.vo.ClienteVO;

public class GeraScript {

	private static final String arquivoclientes = "/home/julio/arquivos/clientes_version.xls";//Lista-pop-rua-2018.xls";
	public static ArrayList<ClienteVO> listaClientes = new ArrayList<ClienteVO>();

	static ArrayList<ClienteVO> loadLista() throws FileNotFoundException {
		try {
			FileInputStream arquivo = new FileInputStream(new File(GeraScript.arquivoclientes));
			HSSFWorkbook workbook = new HSSFWorkbook(arquivo);
			HSSFSheet sheetAlunos = workbook.getSheetAt(0);

			Iterator<Row> rowIterator = sheetAlunos.iterator();

			int id =0;
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				ClienteVO aluno = new ClienteVO();

				int count = 0;
				while (cellIterator.hasNext()) {
					count++;
					System.out.print(count+"\t");
					aluno.setId(count);
					Cell cell = cellIterator.next();
					
				//*****	System.out.println(cell + "\t");
				
					int index = count;//cell.getColumnIndex();
					if (index == 1) {
						aluno.setNome(cell.getStringCellValue());
					}else if (index == 2) {
						aluno.setCnpjCpf(cell.getStringCellValue());
					}else if(index == 3) {
						aluno.setEndereco(cell.getStringCellValue() == null ? "" : cell.getStringCellValue());
					}else if(index == 4) {
						aluno.setCidade(cell.getStringCellValue() == null ? "" : cell.getStringCellValue());
					}else if(index == 5) {
						aluno.setTelefone(String.valueOf(cell.getNumericCellValue() == 0 ? "" : cell.getNumericCellValue()));
					}else if(index ==6) {
						aluno.setContato(cell.getStringCellValue()== null ? "" : cell.getStringCellValue());
					}else if(index == 7) {
						aluno.setEmail(cell.getStringCellValue() == null ? "" : cell.getStringCellValue());
					}else if(index ==8) {
						aluno.setVendedor(cell.getStringCellValue() == null ? "" : cell.getStringCellValue());
					}
				}
				id++;
				aluno.setId(id);
				listaClientes.add(aluno);
				if (id == 343) {
					break;
				}
				System.out.println();
			}
			arquivo.close();
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("Arquivo Excel não encontrado!");
		}
		return listaClientes;
	}

	/*private static String getStringValue(Cell cell) {
		if (cell.getStringCellValue().isEmpty()) {
			return "";
		}
		if (String.valueOf(cell.getNumericCellValue()).isEmpty()) {
			return "";
		}
		return null;
	}*/

	public static void main(String[] args) {
		try {
			listaClientes = loadLista();
			for (ClienteVO vo : listaClientes) {
				System.out.println(vo);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}

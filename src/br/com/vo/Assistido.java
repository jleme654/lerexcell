package br.com.vo;

import java.sql.Timestamp;

public class Assistido {

	private int id;
	private String nome;
	private Timestamp dataVisita;
	private String rg;
	private String cpf;

	public String getNome() {
		return nome;
	}


	public void setNome(String nome) {
		this.nome = nome;
	}

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public Timestamp getDataVisita() {
		return dataVisita;
	}

	public void setDataVisita(Timestamp dataVisita) {
		this.dataVisita = dataVisita;
	}

	public String getRg() {
		return rg;
	}

	public void setRg(String rg) {
		this.rg = rg;
	}

	public String getCpf() {
		return cpf;
	}

	public void setCpf(String cpf) {
		this.cpf = cpf;
	}
	@Override
	public String toString() {
		return "INSERT INTO `assistido`(`id`, `nome`, `dataṼisita`, `rg`, `cpf`) VALUES ("+ id +", '"+ nome +"', '"+dataVisita +"', '"+ rg + "', '" + cpf + "');";
	}

}

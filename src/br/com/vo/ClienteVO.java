package br.com.vo;

public class ClienteVO {

	private int id;
	private String nome;
	private String cnpjCpf;
	private String endereco;
	private String cidade;
	private String telefone;
	private String contato;
	private String email;
	private String vendedor;

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getNome() {
		return nome;
	}

	public void setNome(String nome) {
		this.nome = nome;
	}

	public String getCnpjCpf() {
		return cnpjCpf;
	}

	public void setCnpjCpf(String cnpjCpf) {
		this.cnpjCpf = cnpjCpf;
	}

	public String getEndereco() {
		return endereco;
	}

	public void setEndereco(String endereco) {
		this.endereco = endereco;
	}

	public String getCidade() {
		return cidade;
	}

	public void setCidade(String cidade) {
		this.cidade = cidade;
	}

	public String getTelefone() {
		return telefone;
	}

	public void setTelefone(String telefone) {
		this.telefone = telefone;
	}

	public String getContato() {
		return contato;
	}

	public void setContato(String contato) {
		this.contato = contato;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public String getVendedor() {
		return vendedor;
	}

	public void setVendedor(String vendedor) {
		this.vendedor = vendedor;
	}

	@Override
	public String toString() {
		return "ClienteVO [id=" + id + ", nome=" + nome + ", cnpjCpf=" + cnpjCpf + ", endereco=" + endereco
				+ ", cidade=" + cidade + ", telefone=" + telefone + ", contato=" + contato + ", email=" + email
				+ ", vendedor=" + vendedor + "]";
	}

}

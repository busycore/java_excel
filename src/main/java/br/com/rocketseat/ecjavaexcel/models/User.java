package br.com.rocketseat.ecjavaexcel.models;


import javax.persistence.*;
import java.math.BigDecimal;
import java.time.LocalDate;

@Entity
public class User {


    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Integer id;

    @Column
    private String nome;

    @Column
    private Integer idade;

    @Column
    private LocalDate data_nascimento;

    @Column
    private BigDecimal saldo;

    public User() {
    }

    public User(Integer id, String nome, Integer idade, LocalDate data_nascimento, BigDecimal saldo) {
        this.id = id;
        this.nome = nome;
        this.idade = idade;
        this.data_nascimento = data_nascimento;
        this.saldo = saldo;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getNome() {
        return nome;
    }

    public void setNome(String nome) {
        this.nome = nome;
    }

    public Integer getIdade() {
        return idade;
    }

    public void setIdade(Integer idade) {
        this.idade = idade;
    }

    public LocalDate getData_nascimento() {
        return data_nascimento;
    }

    public void setData_nascimento(LocalDate data_nascimento) {
        this.data_nascimento = data_nascimento;
    }

    public BigDecimal getSaldo() {
        return saldo;
    }

    public void setSaldo(BigDecimal saldo) {
        this.saldo = saldo;
    }
}

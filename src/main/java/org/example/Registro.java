package org.example;

public class Registro {
    private String agencia;
    private String intermediario;
    private String ramo;
    private int año;
    private double primaNeta;
    private String detalleAnexo;
    private double comision;

    public Registro(String agencia, String intermediario, String ramo, int año, double primaNeta, String detalleAnexo, double comision) {
        this.agencia = agencia;
        this.intermediario = intermediario;
        this.ramo = ramo;
        this.año = año;
        this.primaNeta = primaNeta;
        this.detalleAnexo = detalleAnexo;
        this.comision = comision;
    }

    @Override
    public String toString() {
        return "Agencia: " + agencia +
               ", Intermediario: " + intermediario +
               ", Ramo: " + ramo +
               ", Año: " + año +
               ", Prima Neta: " + primaNeta +
               ", Detalle Anexo: " + detalleAnexo +
               ", % Comisión: " + comision;
    }

    public String getAgencia() {
        return agencia;
    }

    public void setAgencia(String agencia) {
        this.agencia = agencia;
    }

    public String getIntermediario() {
        return intermediario;
    }

    public void setIntermediario(String intermediario) {
        this.intermediario = intermediario;
    }

    public String getRamo() {
        return ramo;
    }

    public void setRamo(String ramo) {
        this.ramo = ramo;
    }

    public int getAño() {
        return año;
    }

    public void setAño(int año) {
        this.año = año;
    }

    public double getPrimaNeta() {
        return primaNeta;
    }

    public void setPrimaNeta(double primaNeta) {
        this.primaNeta = primaNeta;
    }

    public String getDetalleAnexo() {
        return detalleAnexo;
    }

    public void setDetalleAnexo(String detalleAnexo) {
        this.detalleAnexo = detalleAnexo;
    }

    public double getComision() {
        return comision;
    }

    public void setComision(double comision) {
        this.comision = comision;
    }

    

    // Agrega los getters para acceder a los atributos
    // (omitiendo para mantener el código más breve)
}


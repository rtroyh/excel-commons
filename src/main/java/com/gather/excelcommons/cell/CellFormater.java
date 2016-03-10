package com.gather.excelcommons.cell;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excel-commons
 * User: rodrigotroy
 * Date: 10-03-16
 * Time: 14:50
 */
public class CellFormater {
    private boolean esColumnaVisible;
    private boolean esTexto;
    private boolean esNumerico;
    private boolean esPorcentual;
    private boolean esFecha;
    private boolean esImagen;
    private boolean noUsaDecimales;

    private List<Object> properties;

    public CellFormater(List<Object> properties) {
        this.properties = properties;
        this.configure();
    }

    private void configure() {
        esColumnaVisible = properties.get(4).equals(1) || properties.get(4).equals(3);
        esTexto = properties.get(1).equals(1);
        esNumerico = properties.get(1).equals(2);
        esPorcentual = properties.get(1).equals(3);
        esFecha = properties.get(1).equals(4);
        esImagen = properties.get(1).equals(5);
        noUsaDecimales = properties.get(2).equals(0);
    }

    public boolean isEsColumnaVisible() {
        return esColumnaVisible;
    }

    public void setEsColumnaVisible(boolean esColumnaVisible) {
        this.esColumnaVisible = esColumnaVisible;
    }

    public boolean isEsTexto() {
        return esTexto;
    }

    public void setEsTexto(boolean esTexto) {
        this.esTexto = esTexto;
    }

    public boolean isEsNumerico() {
        return esNumerico;
    }

    public void setEsNumerico(boolean esNumerico) {
        this.esNumerico = esNumerico;
    }

    public boolean isEsPorcentual() {
        return esPorcentual;
    }

    public void setEsPorcentual(boolean esPorcentual) {
        this.esPorcentual = esPorcentual;
    }

    public boolean isEsFecha() {
        return esFecha;
    }

    public void setEsFecha(boolean esFecha) {
        this.esFecha = esFecha;
    }

    public boolean isEsImagen() {
        return esImagen;
    }

    public void setEsImagen(boolean esImagen) {
        this.esImagen = esImagen;
    }

    public boolean isNoUsaDecimales() {
        return noUsaDecimales;
    }

    public void setNoUsaDecimales(boolean noUsaDecimales) {
        this.noUsaDecimales = noUsaDecimales;
    }
}

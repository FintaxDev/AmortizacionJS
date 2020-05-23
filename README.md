# Tabla de Amortización con Excel JavaScript API

Guía para crear una tabla de amortización en la que los pagos son iguales para cada periodo; los intereses, decrecientes; y las amortizaciones, crecientes. 

Para ello, se utilizará **Script Lab** la cual es una herramienta que permite el manejo de la API de **Office JavaScript** en Excel.

![excel-js](https://user-images.githubusercontent.com/65678421/82627833-eebaf280-9bb0-11ea-843f-ee5ba8f00589.png)

En la imagen anterior se observan dos paneles. Uno es el **Code**, un editor de código en donde se puede programar con TypeScript (idéntico a JavaScript), HTML y CSS. El otro es **Ejecutar**, aquí se muestra el resultado de lo que se escribe en **Code**.

## 1. Comenzando

### 1.1. Requisitos

* Conexión a internet (forzoso).

* Excel 2013 o superior.

* Tener [Script Lab](https://appsource.microsoft.com/es-ES/product/office/wa104380862) instalado.

```
En Excel: pestaña Insertar > Obtener complementos > en el cuadro de búsqueda pon Script Lab > Agregar
```

### 1.2. Importando el snippet a Excel

* Abrir el panel **Code**.

```
En Excel: pestaña Script Lab > Code
```

* En el panel haz clic en **Import** y pega el siguiente link:

```
https://gist.github.com/yairtestas/0c6e85d393cf8a17b612bd1e7da03a94
```

* Abrir el panel **Ejecutar**.

```
En Excel: pestaña Script Lab > Ejecutar
```

## 2. Caso Práctico

```
Suponga que usted va a adquirir un préstamo por $25,000 a pagar en 2 años.
Los pagos son mensuales a una tasa de interés del 6%. 
El pago de la primera mensualidad deberá realizarse al siguiente mes.
```

En el panel **Ejecutar** introduzca los siguientes valores:

| Campo | Valor | Comentarios |
| :--- | :---: | --- |
| Monto | 25000 | Insertar números sin comas|
| Tasa | 0.06 | Utilizar formato decimal |
| Periodos | 24 | Los pagos son mensuales (2 años = 24 meses) |
| Periodicidad | 12 | Mensual: 12, Bimestral: 6, Trimestral: 4, Cuatrimestral: 3, Anual: 1 |
| Primer pago | 1 | El periodo en el que se hará el primer pago, es decir, el mes 1 |

Al dar clic en el botón **`Ejecutar`** se creará la tabla de amortización.

![GIF Excel](https://user-images.githubusercontent.com/65678421/82718826-a06a2a00-9c6a-11ea-8aaa-63c104a6cebd.gif)

---

© **Yair Testas** \| Financial & Tax Technology

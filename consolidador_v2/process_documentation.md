# Documentación del Proceso

## Entrada

## Proceso

### Limpieza de los datos

### Obtención de datos básicos de la cuenta
- Saldo inicial de la cuenta
- Saldo final de la cuenta

- Ingresos totales
- Ingresos totales por Biopago

- Comisiones totales
- Retiros totales

### Validación de datos
En esta etapa, se verifica que todos los pagos para una liquidación determinada sean encontrados.

### Consolidación de Pagos Normales

### Consolidación de Biopago
- Pagos efectivos del mes anterior
- Pagos en el mes actual
- Pagos efectivos del próximo mes

- Comisión de este mes

- Pagos efectivos en otros meses

## Salida

### Pagos Pendientes
Una lista de todos los pagos pendientes por liquidar en el mes dado, con las siguientes columnas:
  - Fecha
  - Referencia => 'Número'
  - Monto

### El Libro Contable
Este libro contendrá todos los ingresos del mes dado. Los datos en el libro:

Para todas las entradas en este libro, tendrán las siguientes columnas:
- Fecha: Será la fecha para las comisiones y la fecha de liquidación para los ingresos.
- Referencia: Solo 6 dígitos.
- Descripción: Serán las descripciones del extracto bancario.
- Código de Liquidación => 'Código': Este será el código de liquidación.
- Débito => 'Debe': El monto para los ingresos.
- Crédito => 'Haber': El monto para las comisiones y transferencias de débito.

Todas las entradas se ordenan por fecha. En el momento de imprimirlo, se debe agregar una columna llamada 'Saldo' que contenga el balance de la cuenta.
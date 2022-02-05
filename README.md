# GeneradorComprobantes-Automatico

Generador de Comprobantes de AFIP Automatico, usando un Excel en donde:
   - La Columna "A" del Excel tiene las fechas de Periodo Desde.
   - La Columna "B" del Excel tiene el CUIL/CUIT del receptor. (El programa agarra la primer instancia que encuentra de un numero de 10 u 11 o 12 digitos)
   - La Columna "C" del Excel tiene el importe.

## Uso del programa
1. Se selecciona el archivo Excel donde estan los datos, luego puede elegir el número de fila del Excel que desea trabajar

2. Ahora debe llenar las demás celdas del programa (__CUIL/CUIT__, __Clave__, __Fecha de Comprobante HASTA__, __Empresa y Servicio__).
     - *Al seleccionar la fila, el programa autorellena con la fecha Periodo Desde presente del Excel, pero puede cambiarla simplemente escribiendo en la celda __Periodo Desde__.*

3. Tocar el botón de __Empezar__.

4. Una vez terminado de correr volverá a aparecer la ventana del programa con todos los datos previos ya cargados excepto el de Periodo Desde y el numero de fila, que habran cambiado (*el número de fila aumenta una unidad, con lo que se autorellena del Excel __Periodo Desde__*).

6. Generar otro comprobante o cerrar el programa.

> El programa cuenta con un Fail Safe, ante cualquier problema se puede llevar el mouse hasta la esquina izquierda superior de la computadora y el programa se detendrá. ;)

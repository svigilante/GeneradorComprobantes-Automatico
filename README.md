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

> El programa cuenta con un Fail Safe, ante cualquier problema se puede llevar el mouse hasta la esquina izquierda superior de la computadora y el programa se detendrá.

## ¡La App!
<img width="760" alt="image" src="https://user-images.githubusercontent.com/90156823/169200236-d55c9f74-22eb-4a13-b1f3-a1079d5c7bcf.png">

_Ventana Inicial._


<img width="675" alt="image" src="https://user-images.githubusercontent.com/90156823/169201455-0fec621d-3781-45d9-849d-62e9a003c9f6.png">

_Ventana que aparece al presionar "Buscar Archivo" (solo nos dejará elegir archivos Excel.)_


![ezgif com-video-to-gif](https://user-images.githubusercontent.com/90156823/169204814-1709a3e4-a3b8-4da8-ab67-29b814766ea7.gif)

_Podemos ver que la celda Periodo Facturado "Desde" se actualiza automáticamente con cada cambio de la celda en donde indicamos la primer fila del Excel (a la derecha de la ventana.)_

### Una vez llenadas todas las celdas, se debe tocar el botón "Empezar" para comenzar el proceso automático donde se abrirá el navegador y el programa completará los datos en la página de la AFIP a partir de lo extraido del Excel y los datos llenados.

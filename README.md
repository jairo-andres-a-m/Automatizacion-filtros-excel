# Automatizacion-filtros-excel 👨‍💻📃
  
Esta es una aplicacion hecha en Python, que ayuda a filtrar informacion de una "base de datos" de colegios. Filtra por Colegio, Id y Fechas para facilitar la aplicacion de modificaciones en las cantidades de suministros. Permite configurar si se filtra un dia o todos los dias del mes y además permite filtrar y desfiltrar los dias de entre semana.

Consiste en una ventana hecha con el modulo tkinter en la cual se ingresa un texto extraido de un PDF, el cual es un "documento oficial".  
  
 <img width="419" height="179" alt="image" src="https://github.com/user-attachments/assets/16de2d3f-970a-4fc9-9a15-dc3157400fed" />

  
La aplicacion se conecta a un Excel usando el modulo xlwings. Al hacer clic en filtrar, o usando la tecla <F9> (tecla asignada a esta accion para mayor facilidad en un proceso repetitivo), la aplicacion extrae los datos relevantes del texto y aplica el filtro a nuestro Excel. De esta manera se facilita la aplicacion manual de muchos cambios y se minimizan errores. 

Ejemplo de los "documentos oficiales" en PDF:  
![image](https://github.com/user-attachments/assets/1d4d55bc-728d-45a2-8169-d89381571b8a)
  
Ejemplo de la base en Excel filtrada:  
![image](https://github.com/user-attachments/assets/d0f4ef55-40d4-4d40-90c1-83827938f001)

Como ultima modificacion se ha cambiado la extension del script a .pyw para que ejecute de una vez como un programa, sin que el usuario vea el terminal corriendo python.

# Automatizacion-filtros-excel 👨‍💻📃
  
Esta es una aplicacion hecha en Python, que ayuda a filtrar informacion de una "base de datos" de colegios. Filtra por Colegio, Id y Fechas para facilitar la aplicacion de modificaciones en las cantidades de suministros de acuerdo a "documentos oficiales".

Consiste en una ventana hecha con el modulo tkinter en la cual se ingresa un texto extraido de un PDF, el "documento oficial". Este texto contiene la informacion de la modificacion a realizar.  
  
![image](https://github.com/user-attachments/assets/df5ab186-6bf4-40b0-8f1f-c5c31d134d54)  
  
La aplicacion se conecta a un Excel usando el modulo xlwings. Al hacer clic en filtrar, la aplicacion extrae los datos relevante del texto y aplica el filtro en Excel, esto para facilitar la aplicacion manual de cambios y minimizar errores. 

Ejemplo de los "documentos oficiales" en PDF:  
![image](https://github.com/user-attachments/assets/5b77698a-922d-485e-8269-97c4068e6cce)
  
Ejemplo de la base en Excel filtrada:  
![image](https://github.com/user-attachments/assets/d0f4ef55-40d4-4d40-90c1-83827938f001)

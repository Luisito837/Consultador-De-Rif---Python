
# Consultador de Python

Herramienta para hacer multiples consultas al seniat.

Esta herramienta pide el excel de las rutas y toma los rifs de los clientes para hacer peticiones mediante python a la pagina del SENIAT, con el fin de conocer que clientes son Contribuyentes Especiales o no.





## Uso (Aplicacion/Release)

Para usar la aplicacion se debe de seguir esta serie de pasos:

1 - Descargue la version mas reciente desde la ventana de Releases.

2 - Al abrir la Aplicacion cargue su excel mediante el boton o pulse y arrastre el excel hacia la aplicacion. De ambas maneras es valido.

3 - Al cargar el excel se desplegara una lista con todas las Rutas disponibles y el numero de clientes asociados a esa Ruta. Aqui podra seleccionar las Rutas a las cuales quiere consultar marcado o desmarcado la casilla de verificación (CheckBox).

4 - Para empezar a procesar solo presione el boton y el programa empezara a trabajar `El proceso varia segun la cantidad de peticiones / clientes.`

5 - Al finalizar se generara un nuevo excel con el mismo nombre del anterior. Y se agrega "_PROCESADO" al final del archivo.




## Uso (Codigo Fuente/Source Code)

Para usar el codigo fuente en python debe de descargar las siguentes librerias ejecutando los suguientes comandos:

CMD / TERMINAL:

```bash
  pip install pandas openpyxl requests beautifulsoup4 opencv-python numpy easyocr tkinterdnd2 pyinstaller
```

Ejecute el Visual Studio Code y ya podra ejecutarlo
